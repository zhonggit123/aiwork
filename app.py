# -*- coding: utf-8 -*-
"""
可视化 Web 服务：上传 Word → AI 解析 → 表格展示/编辑 → 一键提交。
启动：uvicorn app:app --reload --host 0.0.0.0 --port 8765
"""
import os
import sys
import tempfile
import uuid
from pathlib import Path

import yaml

# 打包成 exe 时：配置从 exe 同目录读取，静态资源从打包目录读取
if getattr(sys, "frozen", False):
    _BASE_DIR = Path(sys.executable).parent
    _RESOURCE_DIR = Path(sys._MEIPASS)
else:
    _BASE_DIR = _RESOURCE_DIR = Path(__file__).parent
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from word_reader import read_word_text, chunk_text, extract_word_images
from llm_extract import (
    extract_questions_from_word_chunks_async,
    classify_file_fast,
    parse_exam_with_answers,
    parse_single_file,
)
from submit_api import submit_all as submit_all_api

app = FastAPI(title="Word 题库录入")

# 允许浏览器插件、本地页面调用接口
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 全局 config，首次请求时加载
_CONFIG = None

# ── 图片会话管理 ──────────────────────────────────────────────────────────────
# session_id → 图片临时目录 Path
_IMAGE_SESSIONS: dict = {}
# 最多保留最近 N 个会话的图片（防止磁盘占用无限增长）
_MAX_IMAGE_SESSIONS = 8


def _get_image_session_dir(session_id: str) -> Path:
    """返回（并创建）该会话的图片临时目录。"""
    img_base = Path(tempfile.gettempdir()) / "ai_luti_images"
    session_dir = img_base / session_id
    session_dir.mkdir(parents=True, exist_ok=True)
    return session_dir


def _register_image_session(session_id: str) -> None:
    """注册新会话，超过上限时删除最早的会话目录。"""
    import shutil
    _IMAGE_SESSIONS[session_id] = _get_image_session_dir(session_id)
    if len(_IMAGE_SESSIONS) > _MAX_IMAGE_SESSIONS:
        oldest_key = next(iter(_IMAGE_SESSIONS))
        old_dir = _IMAGE_SESSIONS.pop(oldest_key)
        try:
            shutil.rmtree(old_dir, ignore_errors=True)
        except Exception:
            pass


def _save_option_images(extracted_images: list, session_id: str) -> list:
    """将提取到的图片字节写入会话临时目录，返回带 filename 字段的列表（去除 image_bytes）。"""
    session_dir = _get_image_session_dir(session_id)
    saved = []
    for item in extracted_images:
        ext = item.get("image_ext", "jpg")
        filename = f"img_{item['image_index']:04d}.{ext}"
        dest = session_dir / filename
        dest.write_bytes(item["image_bytes"])
        saved.append({
            "para_index": item["para_index"],
            "para_text": item["para_text"],
            "prev_para_text": item["prev_para_text"],
            "option_label": item["option_label"],
            "image_ext": ext,
            "image_index": item["image_index"],
            "filename": filename,
        })
    return saved


def _enrich_questions_with_option_images(
    questions: list,
    saved_images: list,
    session_id: str,
    base_url: str,
    page_structure: list | None = None,
) -> list:
    """将 saved_images 中的图片 URL 注入到题目选项里，替换 <<IMG>> 占位符。

    匹配策略：按文档顺序（sequential）将图片依次分配给含有 <<IMG>> 选项的题目/小题。
    支持两种结构：
      - 单题顶层选项：q.options
      - 多小题：q.blanks[j].options（AI 规定小题选项不放顶层）
    """
    if not saved_images:
        return questions

    # 构建 slot optionKind 映射
    slot_option_kinds: dict = {}
    if page_structure:
        for slot in page_structure:
            idx = slot.get("index", 0)
            slot_option_kinds[idx] = slot.get("optionKind", "text")

    img_ptr = 0

    def _is_img_placeholder(v) -> bool:
        return isinstance(v, str) and v.strip() == "<<IMG>>"

    def _replace_options(opts: list) -> list:
        nonlocal img_ptr
        new_opts = []
        for opt in opts:
            if _is_img_placeholder(opt) and img_ptr < len(saved_images):
                img_info = saved_images[img_ptr]
                url = f"{base_url}/api/images/{session_id}/{img_info['filename']}"
                new_opts.append(url)
                img_ptr += 1
            else:
                new_opts.append(opt)
        return new_opts

    for q_idx, q in enumerate(questions):
        slot_idx = q_idx + 1  # 1-indexed
        opt_kind = slot_option_kinds.get(slot_idx, "text")

        blanks = q.get("blanks") or []
        if blanks:
            # 多小题：遍历每个 blank，各自替换其 options
            for blank in blanks:
                b_opts = blank.get("options") or []
                if not b_opts:
                    continue
                all_img = all(_is_img_placeholder(o) for o in b_opts if o is not None)
                if (opt_kind == "image") or all_img:
                    blank["options"] = _replace_options(b_opts)
        else:
            # 单题：替换顶层 options
            options = q.get("options") or []
            if not options:
                continue
            all_img = all(_is_img_placeholder(o) for o in options if o is not None)
            is_image_option = (opt_kind == "image") or all_img
            if is_image_option:
                q["options"] = _replace_options(options)

    return questions


def get_config() -> dict:
    global _CONFIG
    if _CONFIG is None:
        path = _BASE_DIR / "config.yaml"
        if not path.exists():
            raise FileNotFoundError("请复制 config.example.yaml 为 config.yaml 并填写（与 exe 同目录）")
        with open(path, "r", encoding="utf-8") as f:
            _CONFIG = yaml.safe_load(f)
    return _CONFIG


@app.get("/api/images/{session_id}/{filename}")
async def serve_option_image(session_id: str, filename: str):
    """提供 Word 图片提取会话中的图片文件（供插件填充图片选项使用）。"""
    # 安全检查：不允许路径穿越
    if "/" in filename or "\\" in filename or ".." in filename:
        raise HTTPException(400, "非法文件名")
    session_dir = Path(tempfile.gettempdir()) / "ai_luti_images" / session_id
    img_path = session_dir / filename
    if not img_path.exists():
        raise HTTPException(404, f"图片不存在: {filename}")
    ext = img_path.suffix.lstrip(".").lower()
    mime_map = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png",
                "gif": "image/gif", "bmp": "image/bmp", "webp": "image/webp"}
    media_type = mime_map.get(ext, "image/jpeg")
    return FileResponse(str(img_path), media_type=media_type)


class SubmitBody(BaseModel):
    questions: list  # list of {type, question, options, answer, explanation}


class DebugPageHtmlBody(BaseModel):
    """插件一键保存当前题 HTML 到项目，供 AI 读取真实 DOM 优化识别。"""
    html: str
    filename: str = "last-question.html"


async def _read_word_content(content: bytes) -> str:
    """读取 Word 文件内容为文本（异步，不阻塞事件循环）。"""
    import asyncio
    
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as fp:
        fp.write(content)
        tmp = fp.name
    try:
        text = await asyncio.to_thread(read_word_text, tmp)
    finally:
        Path(tmp).unlink(missing_ok=True)
    return text


async def _read_word_content_with_images(content: bytes) -> tuple:
    """读取 Word 文件内容（文字 + 图片），异步，返回 (text, extracted_images)。"""
    import asyncio

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as fp:
        fp.write(content)
        tmp = fp.name
    try:
        text, images = await asyncio.gather(
            asyncio.to_thread(read_word_text, tmp),
            asyncio.to_thread(extract_word_images, tmp),
        )
    finally:
        Path(tmp).unlink(missing_ok=True)
    return text, images


async def _parse_one_file(
    content: bytes,
    field_structure: list | None = None,
    paper_metadata: dict | None = None,
) -> list:
    """解析单个 Word 文件内容，返回题目列表（异步，不阻塞事件循环）。
    
    field_structure: 页面检测到的字段结构，用于生成更精准的 LLM prompt
    paper_metadata: 试卷元信息（总题数、题型分布），用于指导精确解析
    """
    text = await _read_word_content(content)
    
    config = get_config()
    chunks = chunk_text(
        text,
        questions_per_batch=config.get("word", {}).get("questions_per_batch", 5),
    )
    if not chunks:
        return []

    llm = config.get("llm", {})
    # 并发调用 LLM，最多 5 个请求同时进行
    return await extract_questions_from_word_chunks_async(
        chunks,
        base_url=llm.get("base_url", "https://api.openai.com/v1"),
        api_key=llm.get("api_key", ""),
        model=llm.get("model", "gpt-4o-mini"),
        max_concurrency=5,
        field_structure=field_structure,
        paper_metadata=paper_metadata,
    )


@app.post("/api/parse")
async def parse_word(file: UploadFile = File(...)):
    """上传单个 Word，解析为题目 JSON。"""
    if not file.filename or not file.filename.lower().endswith((".docx", ".doc")):
        raise HTTPException(400, "请上传 .docx 文件")
    try:
        content = await file.read()
        questions = await _parse_one_file(content)
        return {"questions": questions}
    except FileNotFoundError as e:
        raise HTTPException(500, detail={"error": str(e), "stage": "config"})
    except Exception as e:
        raise HTTPException(500, detail={"error": str(e), "stage": "parse"})


@app.post("/api/debug-page-html")
async def save_debug_page_html(body: DebugPageHtmlBody):
    """插件一键保存：把当前题目区域 HTML 写入项目 samples/ 目录，无需复制粘贴，AI 可直接读取。"""
    samples_dir = _RESOURCE_DIR / "samples"
    samples_dir.mkdir(exist_ok=True)
    # 只允许写入 .html 文件名，避免路径穿越
    name = (body.filename or "last-question.html").strip()
    if not name.endswith(".html"):
        name += ".html"
    if "/" in name or "\\" in name:
        name = "last-question.html"
    path = samples_dir / name
    path.write_text(body.html, encoding="utf-8")
    return {"ok": True, "path": str(path), "message": f"已保存到 {path.name}，可直接让 AI 读取"}


@app.post("/api/parse-multiple")
async def parse_word_multiple(
    files: list[UploadFile] = File(..., description="多个 .docx 文件"),
    field_structure: str | None = Form(None, description="页面字段结构 JSON"),
    expected_total: str | None = Form(None, description="录题页题目数量，用于约束大模型只提取对应题数"),
    page_structure: str | None = Form(None, description="录题页每题的题型/小题数等，JSON 数组，便于大模型精确对应"),
    model_override: str | None = Form(None, description="覆盖配置中的模型名，如 doubao-seed-2-0-pro-260215"),
    reasoning_effort: str | None = Form(None, description="思考程度：minimal/low/medium/high，默认 medium"),
    debug: str | None = Form(None, description="传 1 或 true 时返回 debug_info（原文+prompt），便于调试"),
):
    """上传多个 Word，智能识别文件类型并合并解析。
    
    流程：
    1. 快速识别文件类型（根据文件名+内容特征，不调 LLM）
    2. 如果有试题+答案材料，合并发给 LLM 按题号对齐解析
    3. 如果只有单文件，直接解析
    
    field_structure: JSON 字符串，包含页面检测到的字段结构
    expected_total: 录题页题数（扩展在「开始检测」时得到），传给大模型约束题数一一对应
    """
    import asyncio
    import json as json_module
    
    if not files:
        raise HTTPException(400, "请至少上传一个 .docx 文件")

    valid = [f for f in files if f.filename and f.filename.lower().endswith((".docx", ".doc"))]
    if not valid:
        raise HTTPException(400, "未找到有效的 .docx 文件")

    # 解析字段结构 JSON
    parsed_field_structure = None
    if field_structure:
        try:
            parsed_field_structure = json_module.loads(field_structure)
            print(f"[parse-multiple] 收到页面字段结构: {parsed_field_structure}")
        except json_module.JSONDecodeError:
            print(f"[parse-multiple] field_structure 解析失败")

    # 试卷题数与录题页每题结构（题型、小题数等），传给大模型以精确对应
    paper_metadata = None
    if expected_total and expected_total.strip():
        try:
            n = int(expected_total.strip())
            if n > 0:
                paper_metadata = {"total_questions": n}
                print(f"[parse-multiple] 试卷结构题数: {n}，将约束大模型只提取 {n} 题")
        except ValueError:
            pass
    if page_structure and page_structure.strip():
        try:
            slots = json_module.loads(page_structure)
            if isinstance(slots, list) and len(slots) > 0:
                if paper_metadata is None:
                    paper_metadata = {}
                paper_metadata["question_slots"] = slots
                print(f"[parse-multiple] 收到录题页结构 {len(slots)} 题，将传给大模型")
        except json_module.JSONDecodeError:
            print("[parse-multiple] page_structure 解析失败")

    config = get_config()
    llm = config.get("llm", {})
    effort = (reasoning_effort or "").strip().lower() if reasoning_effort else "medium"
    if effort not in ("minimal", "low", "medium", "high"):
        effort = "medium"
    llm_params = {
        "base_url": llm.get("base_url", "https://api.openai.com/v1"),
        "api_key": llm.get("api_key", ""),
        "model": (model_override.strip() if model_override and model_override.strip() else None)
                 or llm.get("model", "gpt-4o-mini"),
        "reasoning_effort": effort,
    }
    print(f"[parse-multiple] 使用模型: {llm_params['model']}, 思考程度: {effort}")

    # 1. 读取所有文件内容（文字 + 图片并行提取）
    filenames = [f.filename for f in valid]
    contents = [await f.read() for f in valid]
    text_and_images = await asyncio.gather(*[_read_word_content_with_images(c) for c in contents])

    # 生成本次解析的图片会话 ID，保存提取到的图片到临时目录
    session_id = str(uuid.uuid4())
    all_extracted_images: list = []
    for _text, _imgs in text_and_images:
        all_extracted_images.extend(_imgs)
    if all_extracted_images:
        saved_images = _save_option_images(all_extracted_images, session_id)
        _register_image_session(session_id)
        print(f"[parse-multiple] 从 Word 提取图片 {len(all_extracted_images)} 张，会话 {session_id}")
    else:
        saved_images = []

    # 推断本地服务器 base_url（供图片 URL 拼接）
    _port = 8766
    _image_base_url = f"http://localhost:{_port}"

    files_data = []
    for i in range(len(valid)):
        text_i, _ = text_and_images[i]
        # 快速识别文件类型（不调 LLM）
        file_info = classify_file_fast(filenames[i], text_i)
        files_data.append({
            "filename": filenames[i],
            "text": text_i,
            "content": contents[i],
            **file_info,
        })
        print(f"[parse-multiple] 文件 '{filenames[i]}' -> 类型: {file_info['file_type']}")

    # 2. 分类文件
    exam_files = [f for f in files_data if f["file_type"] == "exam"]
    answer_files = [f for f in files_data if f["file_type"] == "answer_material"]
    combined_files = [f for f in files_data if f["file_type"] == "combined"]
    
    print(f"[parse-multiple] 试题文件: {len(exam_files)} 个, 答案材料: {len(answer_files)} 个, 题答合并: {len(combined_files)} 个")

    want_debug = (debug or "").strip().lower() in ("1", "true", "yes")
    debug_info = None

    try:
        # 3. 根据文件组合选择解析策略
        if len(combined_files) >= 1 and len(exam_files) == 0 and len(answer_files) == 0:
            # 情况 AA：题目+答案+听力材料全在同一文件（内嵌答案格式如 (C)1.A. B. C.）
            print("[parse-multiple] 策略: 题答合并文件解析（内嵌答案）")
            # 多个 combined 文件时合并文本
            combined_text = "\n\n---\n\n".join(f["text"] for f in combined_files)
            out = await parse_single_file(
                combined_text,
                field_structure=parsed_field_structure,
                paper_metadata=paper_metadata,
                return_debug=want_debug,
                **llm_params,
            )
            if want_debug:
                questions, debug_info = out
            else:
                questions = out

        elif len(exam_files) == 1 and len(answer_files) == 1:
            # 情况 A：一个试题 + 一个答案材料 → 合并解析
            print("[parse-multiple] 策略: 试题+答案材料合并解析")
            out = await parse_exam_with_answers(
                exam_files[0]["text"],
                answer_files[0]["text"],
                field_structure=parsed_field_structure,
                paper_metadata=paper_metadata,
                return_debug=want_debug,
                **llm_params,
            )
            if want_debug:
                questions, debug_info = out
            else:
                questions = out

        elif len(exam_files) == 1 and len(answer_files) == 0:
            # 情况 B：只有一个试题文件 → 单独解析
            print("[parse-multiple] 策略: 单试题文件解析")
            out = await parse_single_file(
                exam_files[0]["text"],
                field_structure=parsed_field_structure,
                paper_metadata=paper_metadata,
                return_debug=want_debug,
                **llm_params,
            )
            if want_debug:
                questions, debug_info = out
            else:
                questions = out

        elif len(exam_files) == 0 and len(answer_files) == 1:
            # 情况 C：只有答案材料（可能包含完整信息）→ 单独解析
            print("[parse-multiple] 策略: 答案材料文件解析（可能包含完整题目）")
            out = await parse_single_file(
                answer_files[0]["text"],
                field_structure=parsed_field_structure,
                paper_metadata=paper_metadata,
                return_debug=want_debug,
                **llm_params,
            )
            if want_debug:
                questions, debug_info = out
            else:
                questions = out

        elif len(files_data) == 1:
            # 情况 D：只有一个文件 → 直接解析
            print("[parse-multiple] 策略: 单文件解析")
            out = await parse_single_file(
                files_data[0]["text"],
                field_structure=parsed_field_structure,
                paper_metadata=paper_metadata,
                return_debug=want_debug,
                **llm_params,
            )
            if want_debug:
                questions, debug_info = out
            else:
                questions = out

        else:
            # 情况 E：多个文件，复杂情况 → 合并所有文本一起解析
            print("[parse-multiple] 策略: 多文件合并解析")
            # 优先用试题文件，再加答案材料，combined 文件也包含进来
            all_texts = [f["text"] for f in exam_files] + [f["text"] for f in answer_files] + [f["text"] for f in combined_files]
            if not all_texts:
                all_texts = [f["text"] for f in files_data]

            combined_text = "\n\n---\n\n".join(all_texts)
            out = await parse_single_file(
                combined_text,
                field_structure=parsed_field_structure,
                paper_metadata=paper_metadata,
                return_debug=want_debug,
                **llm_params,
            )
            if want_debug:
                questions, debug_info = out
            else:
                questions = out

        print(f"[parse-multiple] 解析完成，共 {len(questions)} 题")

        # ── 图片选项注入 ──────────────────────────────────────────────────────
        if saved_images:
            parsed_slots = None
            if page_structure and page_structure.strip():
                try:
                    parsed_slots = json_module.loads(page_structure)
                except Exception:
                    parsed_slots = None
            questions = _enrich_questions_with_option_images(
                questions,
                saved_images,
                session_id,
                _image_base_url,
                parsed_slots,
            )
            injected = sum(
                1 for q in questions
                for opt in q.get("options", [])
                if isinstance(opt, str) and opt.startswith(_image_base_url)
            )
            print(f"[parse-multiple] 图片选项注入完成，共替换 {injected} 个选项")

        result = {"questions": questions}
        if debug_info is not None:
            result["debug_info"] = debug_info
        return result
    
    except Exception as e:
        print(f"[parse-multiple] 解析失败: {e}")
        raise HTTPException(500, detail={"error": str(e), "stage": "parse"})


@app.post("/api/submit")
async def submit_questions(body: SubmitBody):
    """将题目列表通过配置的 API 提交。"""
    config = get_config()
    submit_cfg = config.get("submit", {})
    if submit_cfg.get("mode") != "api":
        raise HTTPException(400, "当前配置为 playwright 模式，请在 config 中改为 api 并配置 submit.api")
    api_cfg = submit_cfg.get("api", {})
    url = api_cfg.get("url")
    if not url:
        raise HTTPException(400, "未配置 submit.api.url，请先抓包填写")
    headers = api_cfg.get("headers", {})
    body_mapping = api_cfg.get("body_mapping", {})
    method = api_cfg.get("method", "POST")
    try:
        results = submit_all_api(
            body.questions,
            url=url,
            headers=headers,
            body_mapping=body_mapping,
            method=method,
        )
        return {"success": sum(results), "total": len(results), "results": results}
    except Exception as e:
        raise HTTPException(500, str(e))


@app.get("/api/config")
async def api_config():
    """返回提交方式、默认模型名等前端需要的配置（不暴露 key）。"""
    try:
        config = get_config()
        submit = config.get("submit", {})
        llm = config.get("llm", {})
        return {
            "submit_mode": submit.get("mode", "api"),
            "api_configured": bool(submit.get("api", {}).get("url")),
            "llm": {"model": llm.get("model", "")},
        }
    except FileNotFoundError:
        return {"submit_mode": None, "api_configured": False, "llm": {"model": ""}}


# 前端单页
STATIC_DIR = _RESOURCE_DIR / "static"


@app.get("/", response_class=HTMLResponse)
async def index():
    return FileResponse(STATIC_DIR / "index.html")


if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
