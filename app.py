# -*- coding: utf-8 -*-
"""
可视化 Web 服务：上传 Word → AI 解析 → 表格展示/编辑 → 一键提交。
启动：uvicorn app:app --reload --host 0.0.0.0 --port 8765
"""
import asyncio
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

from word_reader import (
    read_word_text, chunk_text, extract_word_images, extract_word_tables_as_images,
    read_pdf_text, extract_pdf_images, render_pdf_pages_as_images,
)
from llm_extract import (
    extract_questions_from_word_chunks_async,
    classify_file_fast,
    parse_exam_with_answers,
    parse_single_file,
    parse_pdf_with_images,
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
_CONFIG_MTIME = 0.0  # 配置文件修改时间，用于热加载

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
    global _CONFIG, _CONFIG_MTIME
    path = _BASE_DIR / "config.yaml"
    if not path.exists():
        raise FileNotFoundError("请复制 config.example.yaml 为 config.yaml 并填写（与 exe 同目录）")
    # 检测文件修改时间，支持热加载
    mtime = path.stat().st_mtime
    if _CONFIG is None or mtime > _CONFIG_MTIME:
        with open(path, "r", encoding="utf-8") as f:
            _CONFIG = yaml.safe_load(f)
        _CONFIG_MTIME = mtime
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


class TtsRequest(BaseModel):
    """TTS 合成请求体。"""
    text: str
    speaker: str | None = None
    format: str = "mp3"
    sample_rate: int = 24000
    speed_ratio: float = 1.0
    volume_ratio: float = 1.0  # 音量倍率
    dialogue: bool = False  # True 时自动解析 W:/M:/Q:/A: 标记，分配男女声
    # 对话模式下的音色、语速、音量设置
    female_speaker: str | None = None
    male_speaker: str | None = None
    female_speed: float | None = None
    male_speed: float | None = None
    female_volume: float | None = None
    male_volume: float | None = None
    # TTS 2.0 提示词（用于调节语速、情绪等，仅豆包声音复刻音色支持）
    context_texts: str | None = None
    female_context_texts: str | None = None
    male_context_texts: str | None = None
    # TTS 服务商：doubao（豆包/火山引擎）、youdao（有道智云）或 edge（微软 Edge TTS）
    provider: str = "doubao"


# ── TTS 合成 ──────────────────────────────────────────────────────────────────
import re
import base64
import httpx


def _parse_dialogue_lines(text: str) -> list[dict]:
    """解析对话文本，返回 [{ "speaker": "male"|"female", "text": "..." }, ...]。
    
    规则：
    - W: / W： / Q: / Q： 开头 -> female
    - M: / M： / A: / A： 开头 -> male
    - 无标记行 -> male (默认)
    - 连续同性别行合并为一段
    """
    lines = text.strip().split("\n")
    segments: list[dict] = []
    
    female_pattern = re.compile(r"^[WwQq][：:]\s*")
    male_pattern = re.compile(r"^[MmAa][：:]\s*")
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        if female_pattern.match(line):
            speaker = "female"
            content = female_pattern.sub("", line).strip()
        elif male_pattern.match(line):
            speaker = "male"
            content = male_pattern.sub("", line).strip()
        else:
            speaker = "male"
            content = line
        
        if not content:
            continue
        
        # 合并连续同性别段落
        if segments and segments[-1]["speaker"] == speaker:
            segments[-1]["text"] += " " + content
        else:
            segments.append({"speaker": speaker, "text": content})
    
    return segments


def _is_tts20_voice(speaker: str) -> bool:
    """判断是否为声音复刻音色（S_ 开头的音色 ID）"""
    return speaker and speaker.startswith("S_")


async def _synthesize_single(
    text: str,
    speaker: str,
    format: str,
    sample_rate: int,
    speed_ratio: float,
    volume_ratio: float,
    access_key: str,
    app_id: str | None,
    context_texts: str | None = None,
) -> bytes:
    """调用火山引擎 TTS v3 合成单段音频，返回原始音频字节。"""
    # 根据 speaker 名称选择 resource_id
    # S_ 开头的是声音复刻音色，使用 volc.megatts.default
    # uranus 系列用 seed-tts-2.0，其他（moon/mars）用 seed-tts-1.0
    if _is_tts20_voice(speaker):
        resource_id = "volc.megatts.default"  # 声音复刻音色
    elif "_uranus_" in speaker:
        resource_id = "seed-tts-2.0"
    else:
        resource_id = "seed-tts-1.0"
    
    url = "https://openspeech.bytedance.com/api/v3/tts/unidirectional"
    
    # TTS 2.0 音色：使用 context_texts 控制语速等，将提示词用中括号放到文本前面
    final_text = text.strip()
    if _is_tts20_voice(speaker) and context_texts:
        final_text = f"[{context_texts}]{final_text}"
        print(f"[TTS] TTS 2.0 音色，添加 context_texts 前缀: [{context_texts[:50]}...]")
    
    request_body = {
        "user": {"uid": "ai_luti_tts"},
        "req_params": {
            "text": final_text,
            "speaker": speaker,
            "speed_ratio": speed_ratio,
            "loudness_ratio": volume_ratio,  # 豆包 TTS 使用 loudness_ratio 而非 volume_ratio
            "additions": '{"disable_markdown_filter":true,"enable_language_detector":true}',
            "audio_params": {
                "format": format,
                "sample_rate": sample_rate,
            },
        },
    }
    
    # TTS 2.0 音色：添加 context_texts 到 additions
    if _is_tts20_voice(speaker) and context_texts:
        import json as json_module
        additions = {"disable_markdown_filter": True, "enable_language_detector": True, "context_texts": [context_texts]}
        request_body["req_params"]["additions"] = json_module.dumps(additions)
    
    headers = {
        "Content-Type": "application/json",
        "X-Api-Key": access_key,
        "X-Api-Resource-Id": resource_id,
    }
    if app_id:
        headers["X-Api-App-Id"] = app_id
    
    print(f"[TTS] 请求: speaker={speaker}, resource_id={resource_id}, text={final_text[:80]}...")
    print(f"[TTS] Headers: X-Api-Resource-Id={resource_id}, X-Api-App-Id={app_id}")
    
    async with httpx.AsyncClient(timeout=60.0) as client:
        resp = await client.post(url, json=request_body, headers=headers)
        print(f"[TTS] 响应状态: {resp.status_code}")
        
        # 火山引擎返回流式 JSON，每行一个 JSON 对象
        raw_text = resp.text
        print(f"[TTS] 响应长度: {len(raw_text)} 字符")
        
        # 打印前 500 字符用于调试
        if len(raw_text) < 1000:
            print(f"[TTS] 响应内容: {raw_text}")
        else:
            print(f"[TTS] 响应前500字符: {raw_text[:500]}")
        
        resp.raise_for_status()
        
        audio_parts: list[bytes] = []
        
        for line in raw_text.split("\n"):
            line = line.strip()
            if not line:
                continue
            try:
                import json as json_module
                obj = json_module.loads(line)
                # 音频数据可能在 data / data.audio / audio 字段
                audio_b64 = None
                if isinstance(obj.get("data"), dict):
                    audio_b64 = obj["data"].get("audio")
                elif isinstance(obj.get("data"), str):
                    audio_b64 = obj["data"]
                elif obj.get("audio"):
                    audio_b64 = obj["audio"]
                
                if audio_b64:
                    audio_parts.append(base64.b64decode(audio_b64))
            except Exception as e:
                print(f"[TTS] 解析行失败: {e}, line={line[:100]}")
                continue
        
        print(f"[TTS] 解析到 {len(audio_parts)} 个音频片段")
        return b"".join(audio_parts)


async def _synthesize_edge_tts(
    text: str,
    speaker: str,
    speed_ratio: float,
    volume_ratio: float,
) -> bytes:
    """调用微软 Edge TTS API 合成音频，返回原始音频字节。"""
    import edge_tts
    import io
    
    # 语速转换：edge-tts 使用百分比格式，如 "+50%" 或 "-25%"
    # speed_ratio 1.0 = 正常，0.5 = -50%，2.0 = +100%
    speed_percent = int((speed_ratio - 1.0) * 100)
    rate_str = f"+{speed_percent}%" if speed_percent >= 0 else f"{speed_percent}%"
    
    # 音量转换：edge-tts 使用百分比格式
    volume_percent = int((volume_ratio - 1.0) * 100)
    volume_str = f"+{volume_percent}%" if volume_percent >= 0 else f"{volume_percent}%"
    
    print(f"[TTS-Edge] 请求: speaker={speaker}, rate={rate_str}, volume={volume_str}, text={text[:80]}...")
    
    try:
        communicate = edge_tts.Communicate(text, speaker, rate=rate_str, volume=volume_str)
        audio_data = io.BytesIO()
        
        async for chunk in communicate.stream():
            if chunk["type"] == "audio":
                audio_data.write(chunk["data"])
        
        result = audio_data.getvalue()
        print(f"[TTS-Edge] 合成成功，音频大小: {len(result)} 字节")
        return result
    except Exception as e:
        print(f"[TTS-Edge] 合成失败: {e}")
        raise Exception(f"Edge TTS 错误: {str(e)}")


async def _synthesize_youdao(
    text: str,
    speaker: str,
    speed_ratio: float,
    volume_ratio: float,
    app_key: str,
    app_secret: str,
) -> bytes:
    """调用有道智云 TTS API 合成音频，返回原始音频字节。"""
    import hashlib
    import uuid
    import time
    import urllib.parse
    
    url = "https://openapi.youdao.com/ttsapi"
    
    # 生成签名
    salt = str(uuid.uuid4())
    curtime = str(int(time.time()))
    
    # input 计算：q前10个字符 + q长度 + q后10个字符（当q长度大于20）
    q = text.strip()
    if len(q) > 20:
        input_str = q[:10] + str(len(q)) + q[-10:]
    else:
        input_str = q
    
    # sign = sha256(appKey + input + salt + curtime + appSecret)
    sign_str = app_key + input_str + salt + curtime + app_secret
    sign = hashlib.sha256(sign_str.encode('utf-8')).hexdigest()
    
    # 有道 TTS 语速范围：0.5 ~ 2.0，默认 1.0
    speed = str(max(0.5, min(2.0, speed_ratio)))
    # 有道 TTS 音量范围：0.5 ~ 5.0，默认 1.0
    volume = str(max(0.5, min(5.0, volume_ratio)))
    
    data = {
        "q": q,
        "appKey": app_key,
        "salt": salt,
        "sign": sign,
        "signType": "v3",
        "curtime": curtime,
        "format": "mp3",
        "speed": speed,
        "volume": volume,
        "voiceName": speaker,
    }
    
    print(f"[TTS-Youdao] 请求: speaker={speaker}, speed={speed}, volume={volume}, text={q[:80]}...")
    
    async with httpx.AsyncClient(timeout=60.0) as client:
        resp = await client.post(url, data=data)
        print(f"[TTS-Youdao] 响应状态: {resp.status_code}, Content-Type: {resp.headers.get('content-type', '')}")
        
        # 有道 TTS：成功时返回 audio/mp3，失败时返回 application/json
        content_type = resp.headers.get("content-type", "")
        if "audio" in content_type:
            print(f"[TTS-Youdao] 合成成功，音频大小: {len(resp.content)} 字节")
            return resp.content
        else:
            # 返回 JSON 错误
            try:
                err = resp.json()
                error_code = err.get("errorCode", "unknown")
                print(f"[TTS-Youdao] 合成失败: errorCode={error_code}")
                raise Exception(f"有道 TTS 错误: {error_code}")
            except Exception as e:
                print(f"[TTS-Youdao] 解析错误响应失败: {e}, body={resp.text[:200]}")
                raise Exception(f"有道 TTS 错误: {resp.text[:200]}")


@app.post("/api/tts")
async def synthesize_tts(req: TtsRequest):
    """TTS 语音合成接口。
    
    支持两种服务商：
    - doubao（豆包/火山引擎）：默认
    - youdao（有道智云）
    
    支持两种模式：
    1. 普通模式：直接合成 text，使用指定 speaker
    2. 对话模式 (dialogue=True)：解析 W:/M:/Q:/A: 标记，自动分配男女声
    """
    config = get_config()
    provider = req.provider or "doubao"
    
    text = (req.text or "").strip()
    if not text:
        raise HTTPException(400, "text 不能为空")
    
    # 语速/音量设置
    male_speed = req.male_speed if req.male_speed is not None else req.speed_ratio
    female_speed = req.female_speed if req.female_speed is not None else req.speed_ratio
    male_volume = req.male_volume if req.male_volume is not None else req.volume_ratio
    female_volume = req.female_volume if req.female_volume is not None else req.volume_ratio
    
    try:
        if provider == "youdao":
            # ── 有道智云 TTS ──
            youdao_cfg = config.get("youdao_tts", {})
            app_key = youdao_cfg.get("app_key") or os.environ.get("YOUDAO_TTS_APP_KEY", "")
            app_secret = youdao_cfg.get("app_secret") or os.environ.get("YOUDAO_TTS_APP_SECRET", "")
            
            if not app_key or not app_secret:
                raise HTTPException(500, "有道 TTS 未配置，请在 config.yaml 的 youdao_tts 中填写 app_key 和 app_secret")
            
            # 有道默认音色
            male_speaker = req.male_speaker or youdao_cfg.get("male_speaker", "youxiaozhi")
            female_speaker = req.female_speaker or youdao_cfg.get("female_speaker", "youxiaoqin")
            
            if req.dialogue:
                segments = _parse_dialogue_lines(text)
                if not segments:
                    raise HTTPException(400, "对话文本解析后为空")
                
                # 并行合成所有对话段落
                async def _synth_youdao_seg(seg):
                    spk = female_speaker if seg["speaker"] == "female" else male_speaker
                    spd = female_speed if seg["speaker"] == "female" else male_speed
                    vol = female_volume if seg["speaker"] == "female" else male_volume
                    return await _synthesize_youdao(seg["text"], spk, spd, vol, app_key, app_secret)
                
                audio_parts = await asyncio.gather(*[_synth_youdao_seg(seg) for seg in segments])
                combined = b"".join(audio_parts)
            else:
                speaker = req.speaker or female_speaker
                combined = await _synthesize_youdao(text, speaker, female_speed, female_volume, app_key, app_secret)
        
        elif provider == "edge":
            # ── 微软 Edge TTS ──
            edge_cfg = config.get("edge_tts", {})
            male_speaker = req.male_speaker or edge_cfg.get("male_speaker", "zh-CN-YunxiNeural")
            female_speaker = req.female_speaker or edge_cfg.get("female_speaker", "zh-CN-XiaoxiaoNeural")

            if req.dialogue:
                segments = _parse_dialogue_lines(text)
                if not segments:
                    raise HTTPException(400, "对话文本解析后为空")

                # 并行合成所有对话段落
                async def _synth_edge_seg(seg):
                    spk = female_speaker if seg["speaker"] == "female" else male_speaker
                    spd = female_speed if seg["speaker"] == "female" else male_speed
                    vol = female_volume if seg["speaker"] == "female" else male_volume
                    return await _synthesize_edge_tts(seg["text"], spk, spd, vol)
                
                audio_parts = await asyncio.gather(*[_synth_edge_seg(seg) for seg in segments])
                combined = b"".join(audio_parts)
            else:
                speaker = req.speaker or female_speaker
                combined = await _synthesize_edge_tts(text, speaker, female_speed, female_volume)

        else:
            # ── 豆包/火山引擎 TTS ──
            tts_cfg = config.get("tts", {})
            access_key = tts_cfg.get("access_key") or os.environ.get("DOUBAO_TTS_ACCESS_KEY", "")
            app_id = tts_cfg.get("app_id") or os.environ.get("DOUBAO_TTS_APP_ID", "")
            
            if not access_key:
                raise HTTPException(500, "豆包 TTS access_key 未配置，请在 config.yaml 的 tts.access_key 中填写")
            
            male_speaker = req.male_speaker or tts_cfg.get("male_speaker", "zh_male_wennuanahu_moon_bigtts")
            female_speaker = req.female_speaker or tts_cfg.get("female_speaker", "zh_female_wanwanxiaohe_moon_bigtts")
            
            # context_texts 设置（声音复刻音色使用）
            male_context_texts = req.male_context_texts or req.context_texts
            female_context_texts = req.female_context_texts or req.context_texts
            
            if req.dialogue:
                segments = _parse_dialogue_lines(text)
                if not segments:
                    raise HTTPException(400, "对话文本解析后为空")
                
                # 并行合成所有对话段落
                async def _synth_doubao_seg(seg):
                    spk = female_speaker if seg["speaker"] == "female" else male_speaker
                    spd = female_speed if seg["speaker"] == "female" else male_speed
                    vol = female_volume if seg["speaker"] == "female" else male_volume
                    ctx = female_context_texts if seg["speaker"] == "female" else male_context_texts
                    return await _synthesize_single(
                        seg["text"], spk, req.format, req.sample_rate, spd, vol, access_key, app_id, ctx
                    )
                
                audio_parts = await asyncio.gather(*[_synth_doubao_seg(seg) for seg in segments])
                combined = b"".join(audio_parts)
            else:
                speaker = req.speaker or female_speaker
                ctx = req.context_texts
                if speaker == male_speaker:
                    ctx = male_context_texts
                elif speaker == female_speaker:
                    ctx = female_context_texts
                combined = await _synthesize_single(
                    text, speaker, req.format, req.sample_rate, female_speed, female_volume, access_key, app_id, ctx
                )
        
        if not combined:
            raise HTTPException(500, "TTS 合成返回空音频")
        
        return {
            "audioBase64": base64.b64encode(combined).decode("utf-8"),
            "format": req.format,
        }
    
    except httpx.HTTPStatusError as e:
        raise HTTPException(e.response.status_code, f"TTS API 错误: {e.response.text[:200]}")
    except Exception as e:
        raise HTTPException(500, f"TTS 合成失败: {str(e)}")


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
    """读取 Word 文件内容（文字 + 图片 + 表格图片），异步，返回 (text, extracted_images, table_images)。"""
    import asyncio

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as fp:
        fp.write(content)
        tmp = fp.name
    try:
        text, images, tables = await asyncio.gather(
            asyncio.to_thread(read_word_text, tmp),
            asyncio.to_thread(extract_word_images, tmp),
            asyncio.to_thread(extract_word_tables_as_images, tmp),
        )
    finally:
        Path(tmp).unlink(missing_ok=True)
    return text, images, tables


async def _read_pdf_content_with_images(content: bytes) -> tuple:
    """读取 PDF 文件内容（文字 + 图片 + 页面渲染图），异步，返回 (text, extracted_images, page_images)。"""
    import asyncio

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as fp:
        fp.write(content)
        tmp = fp.name
    try:
        text, images, pages = await asyncio.gather(
            asyncio.to_thread(read_pdf_text, tmp),
            asyncio.to_thread(extract_pdf_images, tmp),
            asyncio.to_thread(render_pdf_pages_as_images, tmp),
        )
        # 将页面渲染图转换为与 Word 表格图片兼容的格式
        table_images = []
        for pg in pages:
            table_images.append({
                "table_index": pg["page_index"],
                "image_bytes": pg["image_bytes"],
                "image_ext": pg["image_ext"],
                "first_cell": f"PDF 第 {pg['page_index'] + 1} 页",
            })
    finally:
        Path(tmp).unlink(missing_ok=True)
    return text, images, table_images


async def _read_file_content_with_images(content: bytes, filename: str) -> tuple:
    """根据文件类型读取内容，返回 (text, extracted_images, table_images)。"""
    ext = (filename or "").lower().split(".")[-1] if filename else ""
    if ext == "pdf":
        return await _read_pdf_content_with_images(content)
    else:
        return await _read_word_content_with_images(content)


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
    files: list[UploadFile] = File(..., description="多个 Word (.docx/.doc) 或 PDF 文件"),
    field_structure: str | None = Form(None, description="页面字段结构 JSON"),
    expected_total: str | None = Form(None, description="录题页题目数量，用于约束大模型只提取对应题数"),
    page_structure: str | None = Form(None, description="录题页每题的题型/小题数等，JSON 数组，便于大模型精确对应"),
    model_override: str | None = Form(None, description="覆盖配置中的模型名，如 doubao-seed-2-0-pro-260215"),
    reasoning_effort: str | None = Form(None, description="思考程度：minimal/low/medium/high，默认 medium"),
    debug: str | None = Form(None, description="传 1 或 true 时返回 debug_info（原文+prompt），便于调试"),
):
    """上传多个 Word 或 PDF 文件，智能识别文件类型并合并解析。
    
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
        raise HTTPException(400, "请至少上传一个文件")

    valid = [f for f in files if f.filename and f.filename.lower().endswith((".docx", ".doc", ".pdf"))]
    if not valid:
        raise HTTPException(400, "未找到有效的 Word 或 PDF 文件")

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

    # 1. 读取所有文件内容（文字 + 图片并行提取，支持 Word 和 PDF）
    filenames = [f.filename for f in valid]
    contents = [await f.read() for f in valid]
    text_and_images = await asyncio.gather(*[
        _read_file_content_with_images(c, fn) for c, fn in zip(contents, filenames)
    ])

    # 生成本次解析的图片会话 ID，保存提取到的图片到临时目录
    session_id = str(uuid.uuid4())
    all_extracted_images: list = []
    all_table_images: list = []
    for _text, _imgs, _tables in text_and_images:
        all_extracted_images.extend(_imgs)
        all_table_images.extend(_tables)
    if all_extracted_images:
        saved_images = _save_option_images(all_extracted_images, session_id)
        _register_image_session(session_id)
        print(f"[parse-multiple] 从 Word 提取图片 {len(all_extracted_images)} 张，会话 {session_id}")
    else:
        saved_images = []

    # 保存表格图片
    saved_table_images: list = []
    if all_table_images:
        session_dir = _get_image_session_dir(session_id)
        _register_image_session(session_id)
        for tbl in all_table_images:
            filename = f"table_{tbl['table_index']:04d}.{tbl['image_ext']}"
            dest = session_dir / filename
            dest.write_bytes(tbl["image_bytes"])
            saved_table_images.append({
                "table_index": tbl["table_index"],
                "filename": filename,
                "first_cell": tbl.get("first_cell", ""),
            })
        print(f"[parse-multiple] 从 Word 提取表格图片 {len(all_table_images)} 张")

    # 推断本地服务器 base_url（供图片 URL 拼接）
    _port = 8766
    _image_base_url = f"http://localhost:{_port}"

    files_data = []
    for i in range(len(valid)):
        text_i, imgs_i, tables_i = text_and_images[i]
        ext_i = (filenames[i] or "").lower().split(".")[-1]
        is_pdf = ext_i == "pdf"
        # 对于 PDF，tables_i 实际上是渲染的页面图片（已按 page_index 排序）
        page_images_i = tables_i if is_pdf else []
        
        # 判断 PDF 是否需要视觉识别：
        # - 扫描件：文字极少（< 100 字符）→ 必须用视觉
        # - 可提取文字的 PDF：直接用文字解析，省 AI 视觉 token
        text_len = len((text_i or "").strip())
        is_scanned_pdf = is_pdf and text_len < 100
        
        # 快速识别文件类型（不调 LLM）
        file_info = classify_file_fast(filenames[i], text_i)
        files_data.append({
            "filename": filenames[i],
            "text": text_i,
            "content": contents[i],
            "is_pdf": is_pdf,
            "is_scanned_pdf": is_scanned_pdf,
            "page_images": page_images_i,  # 保留页面图片，扫描件时使用
            **file_info,
        })
        if is_pdf:
            if is_scanned_pdf:
                print(f"[parse-multiple] 文件 '{filenames[i]}' -> PDF 扫描件（文字 {text_len} 字符），将使用视觉识别（{len(page_images_i)} 页）")
            else:
                print(f"[parse-multiple] 文件 '{filenames[i]}' -> PDF（文字 {text_len} 字符），直接解析文字")
        else:
            print(f"[parse-multiple] 文件 '{filenames[i]}' -> 类型: {file_info['file_type']}")

    # 2. 分类文件
    exam_files = [f for f in files_data if f["file_type"] == "exam"]
    answer_files = [f for f in files_data if f["file_type"] == "answer_material"]
    combined_files = [f for f in files_data if f["file_type"] == "combined"]
    
    print(f"[parse-multiple] 试题文件: {len(exam_files)} 个, 答案材料: {len(answer_files)} 个, 题答合并: {len(combined_files)} 个")

    want_debug = (debug or "").strip().lower() in ("1", "true", "yes")
    debug_info = None
    
    # 检查是否有扫描件 PDF 需要视觉识别
    scanned_pdfs = [f for f in files_data if f.get("is_scanned_pdf")]

    try:
        # 3. 根据文件组合选择解析策略
        
        # 特殊情况：扫描件 PDF → 使用视觉模型识别
        if len(scanned_pdfs) > 0 and len(files_data) == len(scanned_pdfs):
            print("[parse-multiple] 策略: PDF 扫描件视觉识别")
            # 合并所有 PDF 的页面图片，保持文件顺序 + 页码顺序
            all_page_images = []
            all_text = []
            page_offset = 0  # 用于多文件时的页码偏移，确保全局顺序
            for f in scanned_pdfs:
                all_text.append(f.get("text", ""))
                file_pages = f.get("page_images", [])
                # 按页码排序后添加，更新全局 page_index
                sorted_pages = sorted(file_pages, key=lambda x: x.get("page_index", 0))
                for pg in sorted_pages:
                    pg_copy = dict(pg)
                    pg_copy["page_index"] = page_offset + pg.get("page_index", 0)
                    all_page_images.append(pg_copy)
                page_offset += len(sorted_pages)
            
            combined_text = "\n\n".join(t for t in all_text if t.strip())
            out = await parse_pdf_with_images(
                combined_text,
                all_page_images,
                field_structure=parsed_field_structure,
                paper_metadata=paper_metadata,
                return_debug=want_debug,
                **llm_params,
            )
            if want_debug:
                questions, debug_info = out
            else:
                questions = out
        
        elif len(combined_files) >= 1 and len(exam_files) == 0 and len(answer_files) == 0:
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

        # 解析页面结构（用于图片注入）
        parsed_slots = None
        if page_structure and page_structure.strip():
            try:
                parsed_slots = json_module.loads(page_structure)
            except Exception:
                parsed_slots = None

        # ── 图片选项注入 ──────────────────────────────────────────────────────
        if saved_images and parsed_slots:
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

        # ── 表格图片注入（题干图片）──────────────────────────────────────────────
        if saved_table_images:
            # 策略：将表格图片分配给录题页结构中有 image_url 字段的题目
            tbl_ptr = 0
            for q_idx, q in enumerate(questions):
                if tbl_ptr >= len(saved_table_images):
                    break
                # 检查该题对应的 slot 是否有 image_url 字段
                slot_idx = q_idx
                slot_has_image = False
                if parsed_slots and slot_idx < len(parsed_slots):
                    slot = parsed_slots[slot_idx]
                    slot_fields = slot.get("currentSlotFields", [])
                    slot_has_image = any(
                        f.get("role") == "image_url" or f == "image_url"
                        for f in slot_fields
                    ) if isinstance(slot_fields, list) else "image_url" in str(slot_fields)
                
                # 如果该题的录题页有图片字段，且题目没有 image_url，则注入
                existing_img = (q.get("image_url") or "").strip()
                if slot_has_image and not existing_img:
                    tbl_info = saved_table_images[tbl_ptr]
                    q["image_url"] = f"{_image_base_url}/api/images/{session_id}/{tbl_info['filename']}"
                    print(f"[parse-multiple] 题目 {q_idx+1} 注入表格图片: {tbl_info['filename']} (slot有image_url字段)")
                    tbl_ptr += 1

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
