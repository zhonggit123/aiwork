# -*- coding: utf-8 -*-
"""
调用大模型 API，将题目文本块解析为结构化 JSON。
支持 OpenAI 兼容接口（火山引擎豆包/通义/智谱/DeepSeek 等）。
"""
import asyncio
import json
import os
import re
from typing import Any, Dict, List

from openai import AsyncOpenAI, OpenAI


# ─── ★ 占位符：避免 AI 错误删除题干中的 ★ ─────────────────────────────────────
# 发送给 AI 前将 ★ 替换为此占位符；AI 收到后不会当作格式符号处理。
# 收到 AI 返回后，非 candidates 字段替换回 ★，candidates 字段直接去掉。
_STAR_PH = "【STAR】"


def _encode_star(text: str) -> str:
    """将文本中的 ★ 替换为占位符，防止 AI 误删。"""
    return text.replace("★", _STAR_PH) if text else text


def _decode_star_questions(questions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """AI 返回后：除 candidates 外所有字段恢复 ★；candidates 中去掉占位符（原意是删掉 ★）。"""
    def restore(s):
        return s.replace(_STAR_PH, "★") if isinstance(s, str) else s

    def strip_star(s):
        return s.replace(_STAR_PH, "").lstrip() if isinstance(s, str) else s

    for q in questions:
        for field in ("question", "keyword", "listening_script", "explanation"):
            if field in q:
                q[field] = restore(q[field])
        # candidates：去掉占位符（原来的 ★ 标记正确答案，不需要显示）
        if "candidates" in q and isinstance(q["candidates"], list):
            q["candidates"] = [strip_star(c) for c in q["candidates"]]
        # blanks
        if "blanks" in q and isinstance(q["blanks"], list):
            for blank in q["blanks"]:
                for field in ("question", "keyword", "answer", "listening_script"):
                    if field in blank:
                        blank[field] = restore(blank[field])
    return questions


def _strip_leading_question_no(text: Any) -> Any:
    """去掉题干开头的大题号，如「第12题」「12.」「（12）」；保留「第1小题」这类小题号。"""
    if not isinstance(text, str):
        return text
    s = text.strip()
    if not s:
        return s
    patterns = [
        r"^第\s*\d+\s*题[\s：:、.．)]*",
        r"^[(（]\s*\d+\s*[)）][\s：:、.．]*",
        r"^\d+\s*[.．、:：)]\s*",
    ]
    for p in patterns:
        s = re.sub(p, "", s).strip()
    return s


def _split_numbered_dialogues(text: str, n: int) -> List[str]:
    """尝试把 '1. W:... 2. W:...' 格式的多段对话按序号拆成 n 份。
    拆分成功（得到恰好 n 段）则返回列表，否则返回空列表。"""
    parts = re.split(r'(?m)(?:^|\n)\s*\d+[.、．]\s*(?=[A-Za-zW])', text.strip())
    parts = [p.strip() for p in parts if p.strip()]
    if len(parts) == n:
        return parts
    # 尝试宽松匹配：序号前可能有换行或空格
    parts2 = re.split(r'\s*(?<!\w)\d+[.、．]\s+(?=W:|M:|Boy:|Girl:|Man:|Woman:)', text.strip())
    parts2 = [p.strip() for p in parts2 if p.strip()]
    if len(parts2) == n:
        return parts2
    return []


def _format_dialogue_script(text: str) -> str:
    """将 W:/M: 对话型听力原文按说话人分行，方便在页面多行文本框中显示。"""
    if not text or not isinstance(text, str):
        return text or ""
    # 在 W:/M:/Boy:/Girl:/Man:/Woman:/Q:/A: 等说话人标记前插入换行（首次出现不加）
    import re as _re
    formatted = _re.sub(
        r'(?<=[^\n])\s+(W:|M:|Boy:|Girl:|Man:|Woman:|Narrator:|Q:|Qs?:|A:)(?=\s)',
        r'\n\1', text
    )
    return formatted.strip()


def _sanitize_english_punctuation(text: Any) -> Any:
    """仅对纯英文/非 CJK 文本做标点半角化，避免平台校验因中英文标点不一致而报错。"""
    if not isinstance(text, str):
        return text
    if not text:
        return text
    if re.search(r"[\u4e00-\u9fff\u3400-\u4dbf\U00020000-\U0002A6DF]", text):
        return text
    return (text
            .replace("\uff0c", ",")
            .replace("\u3002", ".")
            .replace("\uff1f", "?")
            .replace("\uff01", "!")
            .replace("\uff1a", ":")
            .replace("\uff1b", ";")
            .replace("\u2018", "'")
            .replace("\u2019", "'")
            .replace("\u201c", '"')
            .replace("\u201d", '"')
            .replace("\uff08", "(")
            .replace("\uff09", ")")
            .replace("\u3010", "[")
            .replace("\u3011", "]")
            .replace("\u2026", "...")
            .replace("\u2014", "-")
            .replace("\uff5e", "~"))


def _sanitize_question_texts(q: Dict[str, Any]) -> None:
    """规范题目/小题中的纯英文文本标点，覆盖题干、选项、答案、原文、解析等字段。"""
    for field in ("question", "keyword", "answer", "listening_script", "explanation"):
        if field in q:
            q[field] = _sanitize_english_punctuation(q.get(field))
    if isinstance(q.get("options"), list):
        q["options"] = [_sanitize_english_punctuation(x) for x in q["options"]]
    if isinstance(q.get("candidates"), list):
        q["candidates"] = [_sanitize_english_punctuation(x) for x in q["candidates"]]
    if isinstance(q.get("blanks"), list):
        for blank in q["blanks"]:
            if not isinstance(blank, dict):
                continue
            for field in ("question", "keyword", "answer", "listening_script"):
                if field in blank:
                    blank[field] = _sanitize_english_punctuation(blank.get(field))
            if isinstance(blank.get("options"), list):
                blank["options"] = [_sanitize_english_punctuation(x) for x in blank["options"]]
            if isinstance(blank.get("candidates"), list):
                blank["candidates"] = [_sanitize_english_punctuation(x) for x in blank["candidates"]]


def _normalize_questions(questions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """对 LLM 输出做轻量纠偏，避免明显不符合录题页约束的数据进入填充阶段。"""
    for q in questions:
        if not isinstance(q, dict):
            continue
        _sanitize_question_texts(q)
        q["question"] = _strip_leading_question_no(q.get("question", ""))
        if "keyword" in q and q.get("keyword") is None:
            q["keyword"] = ""
        if isinstance(q.get("blanks"), list):
            for blank in q["blanks"]:
                if not isinstance(blank, dict):
                    continue
                blank["question"] = _strip_leading_question_no(blank.get("question", ""))
                if "keyword" in blank and blank.get("keyword") is None:
                    blank["keyword"] = ""
                if "listening_script" in blank and blank.get("listening_script") is None:
                    blank["listening_script"] = ""

            # 对话格式化：将各 blank 的 listening_script 中的 W:/M: 对话分行
            for blank in q["blanks"]:
                if isinstance(blank.get("listening_script"), str) and blank["listening_script"]:
                    blank["listening_script"] = _format_dialogue_script(blank["listening_script"])

            # 情形 A 自动拆分：若 blank[0] 含所有内容（其余 blank 为空）且能按序号拆开，
            # 则拆分到各 blank —— 对应「多段独立对话分别对应各小题音频」的情形。
            if (len(q["blanks"]) > 1
                    and q["blanks"][0].get("listening_script")
                    and all(not b.get("listening_script") for b in q["blanks"][1:])):
                split = _split_numbered_dialogues(
                    q["blanks"][0]["listening_script"], len(q["blanks"])
                )
                if split:
                    for i, blank in enumerate(q["blanks"]):
                        if isinstance(blank, dict):
                            blank["listening_script"] = _format_dialogue_script(split[i])
            # 同理：若顶层 listening_script 含所有内容而各 blank 均无，也尝试拆分
            elif (len(q.get("blanks") or []) > 1
                    and (q.get("listening_script") or "")
                    and all(not b.get("listening_script") for b in q["blanks"])):
                split = _split_numbered_dialogues(
                    q["listening_script"], len(q["blanks"])
                )
                if split:
                    for i, blank in enumerate(q["blanks"]):
                        if isinstance(blank, dict):
                            blank["listening_script"] = _format_dialogue_script(split[i])
                    q["listening_script"] = ""

            # 共享对话提升：若所有 blank 的 listening_script 相同且文字较长（>30字），
            # 说明是多小题共用同一段对话，应放在顶层，各 blank 只保留小题专属内容。
            # 规则：共享脚本→顶层 q.listening_script；blank.listening_script 改为
            #       该小题的 blank.question（题目读出的问题句），否则置空。
            blank_scripts = [b.get("listening_script", "") for b in q["blanks"] if isinstance(b, dict)]
            if (blank_scripts
                    and all(s == blank_scripts[0] for s in blank_scripts)
                    and len(blank_scripts[0]) > 30):
                shared = blank_scripts[0]
                # 仅在顶层确实为空时才提升，避免覆盖已正确填入的顶层原文
                if not (q.get("listening_script") or "").strip():
                    q["listening_script"] = shared
                for blank in q["blanks"]:
                    if isinstance(blank, dict):
                        # 用 blank.question 作为小题专属脚本（如果它比共享脚本短很多）
                        bq = (blank.get("question") or "").strip()
                        blank["listening_script"] = bq if (bq and len(bq) < len(shared) * 0.5) else ""


        # 兜底修正：该题型不应把长段转述/表达内容放到顶层 answer，
        # 否则会被填进页面的「设置答案」框，触发“中文或非法字符”校验。
        if q.get("type") == "listening_fill_and_retell":
            ans = (q.get("answer") or "").strip() if isinstance(q.get("answer"), str) else q.get("answer")
            has_blanks = isinstance(q.get("blanks"), list) and len(q.get("blanks") or []) > 0
            if ans and not has_blanks:
                q["blanks"] = [{
                    "question": _strip_leading_question_no(q.get("question", "")),
                    "keyword": (q.get("keyword") or "") if isinstance(q.get("keyword"), str) else "",
                    "answer": ans,
                    "options": [],
                    "listening_script": (q.get("listening_script") or "") if isinstance(q.get("listening_script"), str) else "",
                }]
                q["answer"] = ""
        if q.get("answer") is None:
            q["answer"] = ""
        # 对话格式化：顶层 listening_script 中的 W:/M: 分行
        if isinstance(q.get("listening_script"), str) and q["listening_script"]:
            q["listening_script"] = _format_dialogue_script(q["listening_script"])

        # 解析编号修正：多小题题目的解析里可能用了试卷绝对题号（第14题、第15题），
        # 将其按出现顺序重映射为「第1小题」「第2小题」，保证录入后语义正确。
        if isinstance(q.get("blanks"), list) and len(q["blanks"]) > 1:
            expl = q.get("explanation") or ""
            if expl and re.search(r'第\d+题', expl):
                # 收集所有出现过的绝对题号，按数值排序后逐一替换
                nums = sorted(set(int(m.group(1)) for m in re.finditer(r'第(\d+)题', expl)))
                for rank, num in enumerate(nums, 1):
                    expl = re.sub(rf'第{num}题', f'第{rank}小题', expl)
                q["explanation"] = expl
    return questions


DEFAULT_SYSTEM_PROMPT = """你是一个题目解析助手。用户会给你一段从 Word 里复制出来的题目文本，可能包含多道题。
请严格按下面要求输出 JSON，不要输出任何其他内容。

每道题的结构为：
- type: 题型，如 "single"（单选）、"multiple"（多选）、"judge"（判断）、"blank"（填空）等
- question: 题干文字
- options: 选项列表，如 ["A. xxx", "B. xxx"]；若无选项则为 []
- answer: 答案，如 "A" 或 "A,B" 或 "对/错" 或填空答案
- explanation: 解析/详解，没有则填空字符串 ""

输出格式：一个 JSON 数组，每道题一个对象。例如：
[
  {"type": "single", "question": "题目内容", "options": ["A. 选项1", "B. 选项2"], "answer": "A", "explanation": "解析"},
  ...
]
只输出这一组 JSON 数组，不要 markdown 代码块包裹，不要其他说明。"""


def build_system_prompt(
    field_structure: List[Dict[str, Any]] | None = None,
    paper_metadata: Dict[str, Any] | None = None,
) -> str:
    """根据页面检测到的字段结构和试卷元信息，生成动态的 system prompt。
    
    field_structure 示例：
    [
      {"role": "question", "label": "题干"},
      {"role": "options", "label": "选项"},
      {"role": "answer", "label": "正确答案"},
      {"role": "explanation", "label": "解析"},
      {"role": "type", "label": "题型"},
    ]
    
    paper_metadata 示例：
    {
      "total_questions": 50,
      "question_types": {"听力选择": 15, "阅读选择": 35},
      "structure_summary": "Part A 听力，Part B 阅读"
    }
    """
    # 构建元信息部分（含录题页每题结构：题型、小题数等，便于大模型精确对应）
    metadata_hint = ""
    if paper_metadata:
        total = paper_metadata.get("total_questions")
        types = paper_metadata.get("question_types", {})
        summary = paper_metadata.get("structure_summary", "")
        slots = paper_metadata.get("question_slots")  # 录题页每题的 partName/typeHint/subCount
        if total:
            metadata_hint = f"\n【试卷信息】这份试卷共 {total} 道题。"
            if types:
                type_str = "、".join([f"{k} {v} 题" for k, v in types.items()])
                metadata_hint += f"题型分布：{type_str}。"
            if summary:
                metadata_hint += f"\n结构：{summary}"
        if slots and isinstance(slots, list):
            parts = []
            for s in slots[: (total or len(slots))]:
                idx = s.get("index", len(parts) + 1)
                part = s.get("partName") or ""
                hint = s.get("typeHint") or ""
                sub = s.get("subCount")
                desc = f"序号{idx}"
                if part or hint:
                    desc += f"（{part or hint}"
                    if sub and sub > 1:
                        desc += f"，{sub} 个空/小题"
                    desc += "）"
                parts.append(desc)
            if parts:
                if not metadata_hint:
                    metadata_hint = "\n"
                metadata_hint += "【录题页结构】" + "；".join(parts) + "。\n"
        if metadata_hint:
            metadata_hint += "【重要】请严格按题号提取，确保不重复、不遗漏。每道题只提取一次。\n"
    
    if not field_structure and not paper_metadata:
        return DEFAULT_SYSTEM_PROMPT
    
    if not field_structure:
        # 只有元信息，没有字段结构
        return DEFAULT_SYSTEM_PROMPT.replace(
            "请严格按下面要求输出 JSON，不要输出任何其他内容。",
            f"请严格按下面要求输出 JSON，不要输出任何其他内容。{metadata_hint}"
        )
    
    # 从字段结构中提取需要解析的字段
    detected_roles = {f.get("role") for f in field_structure if f.get("role")}
    detected_labels = {f.get("role"): f.get("label", f.get("role")) for f in field_structure if f.get("role")}
    # 同时收集每个字段的 placeholder（来自页面输入框的灰色提示文字）
    detected_placeholders = {
        f.get("role"): f.get("placeholder", "")
        for f in field_structure if f.get("role") and f.get("placeholder")
    }

    # 构建动态字段说明
    field_descriptions = []
    
    # 根据检测到的字段生成说明
    role_to_desc = {
        "type": ('type', '题型，如 "single"（单选）、"multiple"（多选）、"judge"（判断）、"blank"（填空）等'),
        "question": ('question', '题干文字；若题干开头有题号（如 1. 2. 一、(1) 等）请去掉题号只保留正文'),
        "keyword": ('keyword', '参考单词/送评单词；仅页面存在该输入框时填写，没有则填空字符串 ""'),
        "options": ('options', '选项列表，如 ["A. xxx", "B. xxx"]；若无选项则为 []'),
        "answer": ('answer', '答案，如 "A" 或 "A,B" 或 "对/错" 或填空答案'),
        "explanation": ('explanation', '解析/详解，没有则填空字符串 ""'),
        "score": ('score', '分值，数字类型'),
        "difficulty": ('difficulty', '难度等级'),
        "category": ('category', '题目分类/知识点'),
        "source": ('source', '题目来源'),
    }
    
    # 优先添加检测到的字段
    for role in detected_roles:
        ph = detected_placeholders.get(role, "")
        ph_hint = f"，输入框提示：「{ph}」" if ph else ""
        if role in role_to_desc:
            key, desc = role_to_desc[role]
            label = detected_labels.get(role, "")
            if label and label != key:
                field_descriptions.append(f"- {key}: {desc}（页面字段：{label}{ph_hint}）")
            else:
                field_descriptions.append(f"- {key}: {desc}{('（' + ph_hint.strip('，') + '）') if ph_hint else ''}")
        else:
            # 未知字段也尝试提取
            label = detected_labels.get(role, role)
            field_descriptions.append(f"- {role}: 对应页面字段「{label}」的内容{ph_hint}")
    
    # 补充默认必需字段（如果没检测到）
    default_fields = ["type", "question", "options", "answer", "explanation"]
    for role in default_fields:
        if role not in detected_roles and role in role_to_desc:
            key, desc = role_to_desc[role]
            field_descriptions.append(f"- {key}: {desc}")
    
    fields_text = "\n".join(field_descriptions)
    
    return f"""你是一个题目解析助手。用户会给你一段从 Word 里复制出来的题目文本，可能包含多道题。
请严格按下面要求输出 JSON，不要输出任何其他内容。
{metadata_hint}
【重要】页面表单需要以下字段，请务必提取：
{fields_text}

输出格式：一个 JSON 数组，每道题一个对象。例如：
[
  {{"type": "single", "question": "题目内容", "options": ["A. 选项1", "B. 选项2"], "answer": "A", "explanation": "解析"}},
  ...
]
只输出这一组 JSON 数组，不要 markdown 代码块包裹，不要其他说明。"""


def extract_json_from_response(content: str) -> List[Dict[str, Any]]:
    """从模型回复中剥离并解析 JSON 数组。"""
    content = content.strip()
    # 去掉可能的 ```json ... ``` 包裹
    if "```" in content:
        match = re.search(r"```(?:json)?\s*([\s\S]*?)```", content)
        if match:
            content = match.group(1).strip()
    try:
        data = json.loads(content)
        if isinstance(data, list):
            return data
        if isinstance(data, dict) and "questions" in data:
            return data["questions"]
        return [data]
    except json.JSONDecodeError:
        # 尝试找到第一个 [ 到最后一个 ] 之间的内容
        start = content.find("[")
        end = content.rfind("]") + 1
        if start >= 0 and end > start:
            return json.loads(content[start:end])
        raise


def _is_volc(base_url: str) -> bool:
    """是否火山引擎豆包，需传 extra_body 控制深度思考时用。"""
    return "volces.com" in (base_url or "")


def _volc_extra_body(reasoning_effort: str = "medium") -> dict:
    """豆包：通过 reasoning_effort 调节思考长度。
    minimal=关闭思考，low/medium/high=开启思考并设置对应 level，平衡效果、时延与成本。
    """
    if reasoning_effort == "minimal":
        return {"thinking": {"type": "disabled"}}
    level = reasoning_effort if reasoning_effort in ("low", "medium", "high") else "medium"
    return {"thinking": {"type": "enabled", "level": level}}


def _resolve_api_key(api_key: str, base_url: str) -> str:
    """优先使用传入的 api_key，为空时火山引擎可用环境变量 ARK_API_KEY。"""
    if (api_key or "").strip():
        return (api_key or "").strip()
    if "volces.com" in (base_url or ""):
        return (os.getenv("ARK_API_KEY") or "").strip()
    return ""


def call_llm_extract(
    text_chunk: str,
    *,
    base_url: str,
    api_key: str,
    model: str,
    field_structure: List[Dict[str, Any]] | None = None,
    reasoning_effort: str = "medium",
) -> List[Dict[str, Any]]:
    """将一块题目文本发给 LLM，返回解析后的题目列表。"""
    key = _resolve_api_key(api_key, base_url)
    if not key:
        raise ValueError("未配置 api_key，请在 config.yaml 的 llm.api_key 填写或设置环境变量 ARK_API_KEY")
    system_prompt = build_system_prompt(field_structure)
    client = OpenAI(base_url=base_url, api_key=key)
    kwargs = {"model": model, "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": text_chunk}], "temperature": 0.1}
    if _is_volc(base_url):
        kwargs["extra_body"] = _volc_extra_body(reasoning_effort)
    resp = client.chat.completions.create(**kwargs)
    content = resp.choices[0].message.content or ""
    return _normalize_questions(extract_json_from_response(content))


def extract_questions_from_word_chunks(
    chunks: List[str],
    *,
    base_url: str,
    api_key: str,
    model: str,
) -> List[Dict[str, Any]]:
    """对多个文本块依次调用 LLM，合并返回所有题目（同步版，兼容旧调用）。"""
    all_questions = []
    for chunk in chunks:
        questions = call_llm_extract(
            chunk,
            base_url=base_url,
            api_key=api_key,
            model=model,
        )
        all_questions.extend(questions)
    return _normalize_questions(all_questions)


async def _call_llm_extract_async(
    text_chunk: str,
    *,
    base_url: str,
    api_key: str,
    model: str,
    semaphore: asyncio.Semaphore,
    system_prompt: str,
    reasoning_effort: str = "medium",
) -> List[Dict[str, Any]]:
    """异步调用 LLM，受 semaphore 控制并发数。"""
    key = _resolve_api_key(api_key, base_url)
    async with semaphore:
        client = AsyncOpenAI(base_url=base_url, api_key=key)
        kwargs = {"model": model, "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": _encode_star(text_chunk)}], "temperature": 0.1}
        if _is_volc(base_url):
            kwargs["extra_body"] = _volc_extra_body(reasoning_effort)
        resp = await client.chat.completions.create(**kwargs)
    content = resp.choices[0].message.content or ""
    return _normalize_questions(_decode_star_questions(extract_json_from_response(content)))


async def extract_questions_from_word_chunks_async(
    chunks: List[str],
    *,
    base_url: str,
    api_key: str,
    model: str,
    max_concurrency: int = 5,
    field_structure: List[Dict[str, Any]] | None = None,
    paper_metadata: Dict[str, Any] | None = None,
    reasoning_effort: str = "medium",
) -> List[Dict[str, Any]]:
    """并发调用 LLM 解析所有文本块，最多同时 max_concurrency 个请求。
    结果按原始顺序合并，失败的块跳过并打印警告。
    
    field_structure: 页面检测到的字段结构，用于生成动态 prompt
    paper_metadata: 试卷元信息（总题数、题型分布等），用于指导精确解析
    reasoning_effort: 思考程度 minimal/low/medium/high
    """
    if not chunks:
        return []
    
    # 根据字段结构生成动态 prompt
    system_prompt = build_system_prompt(field_structure, paper_metadata)
    
    sem = asyncio.Semaphore(max_concurrency)
    tasks = [
        _call_llm_extract_async(
            chunk,
            base_url=base_url,
            api_key=api_key,
            model=model,
            semaphore=sem,
            system_prompt=system_prompt,
            reasoning_effort=reasoning_effort,
        )
        for chunk in chunks
    ]
    results = await asyncio.gather(*tasks, return_exceptions=True)
    all_questions: List[Dict[str, Any]] = []
    for i, r in enumerate(results):
        if isinstance(r, Exception):
            print(f"[llm_extract] chunk {i} 解析失败，已跳过: {r}")
        else:
            all_questions.extend(r)
    return _normalize_questions(all_questions)


# ─── 快速文件类型识别（不调 LLM，根据文件名和内容特征）─────────────────────

def classify_file_fast(filename: str, text: str) -> Dict[str, Any]:
    """快速识别文件类型，不调用 LLM，根据文件名关键词和内容特征判断。
    
    返回：{"file_type": "exam|answer_material|combined|other", "has_questions": bool, "has_answers": bool, "has_script": bool}
    combined：题目、答案、听力材料全在同一文件（如 (C)1.A. B. C. 格式，答案内嵌题目行首）
    """
    fn = filename.lower()
    
    # 内容特征检测
    has_options = bool(re.search(r'[A-D]\s*[.、．]\s*\w', text[:3000]))  # 有选项 A. B. C.
    has_answers = '参考答案' in text or bool(re.search(r'\d+[-~]\d+\s*[A-D]', text))  # 1-4 CAAC 格式
    has_script = '听力材料' in text or '听力原文' in text or bool(re.search(r'[MW][：:]\s*\w', text[:2000]))
    has_question_nums = bool(re.search(r'[(（]\s*[)）]\s*\d+\.', text[:1500]))  # (  )1. 格式
    # 内嵌答案格式检测：(C)1. / (A)2. 等，答案字母直接写在题号前括号内
    has_inline_answers = bool(re.search(r'^[（(][A-Da-d][)）]\s*\d+\s*[.．]', text, re.MULTILINE))
    
    # 根据文件名判断（优先级最高）
    # 注意：文件名同时包含"试题"和"答案"时，优先看更具体的关键词
    fn_has_exam = '试题' in fn or '试卷' in fn
    fn_has_answer = '答案' in fn or '材料' in fn
    
    # 内嵌答案格式（combined）：答案随题目，且有题目内容 ——优先于文件名判断
    # 特征：(C)1.A. B. C. 格式的行 + 有选项/有题目
    if has_inline_answers and (has_options or has_question_nums):
        return {"file_type": "combined", "has_questions": True, "has_answers": True, "has_script": has_script}
    
    if fn_has_answer and not fn_has_exam:
        # 文件名明确是答案/材料
        return {"file_type": "answer_material", "has_questions": False, "has_answers": has_answers, "has_script": has_script}
    
    if fn_has_exam and not fn_has_answer:
        # 文件名明确是试题/试卷
        return {"file_type": "exam", "has_questions": True, "has_answers": False, "has_script": False}
    
    if fn_has_exam and fn_has_answer:
        # 文件名同时有试题和答案，根据内容判断
        # 如果有听力材料或参考答案段落，当作答案材料
        if has_script or has_answers:
            return {"file_type": "answer_material", "has_questions": False, "has_answers": has_answers, "has_script": has_script}
        else:
            return {"file_type": "exam", "has_questions": True, "has_answers": False, "has_script": False}
    
    # 文件名没有明确标识，根据内容特征判断
    if has_script or has_answers:
        return {"file_type": "answer_material", "has_questions": False, "has_answers": has_answers, "has_script": has_script}
    
    if has_question_nums and has_options:
        return {"file_type": "exam", "has_questions": True, "has_answers": False, "has_script": False}
    
    # 默认当作试题
    return {"file_type": "exam", "has_questions": True, "has_answers": False, "has_script": False}


# ─── 通用解析 Prompt（不写死具体试卷；目标结构由「页面检测」结果动态注入）────────────────

# 合并解析（试题 + 答案材料）：通用规则，具体题数、一题多空等由 paper_metadata 注入
MERGED_PARSE_PROMPT = """你是一个题目解析助手。用户会给你【试题】和【答案材料】（或其一），请按**录题页结构**解析并输出 JSON。

【通用规则】
- 若下方有【录题页结构】/【目标题数】：输出条数、一题多空以之为准；某题多空时只输出一条，用 blanks: [{question, answer}, ...]，顺序与页面一致。
- 若未提供录题页结构：按文档题号与题型逐条提取；同一表格或同一材料下的多空可合并为一条并用 blanks。
- 试题部分：题干、选项等；答案材料部分：**听力原文**（仅听力材料正文）、参考答案。听后应答类题干可能在答案材料里，选项在试题上。按题号对齐合并。

【内嵌答案规则】答案可能以 (A)/(B)/(C) 形式写在题目行首（如 "(C)1.A.  B．  C．"），请提取括号内字母作为 answer，去掉行首的 (X) 前缀，只保留题干正文和选项内容。

【图片选项规则】若某题的选项在原卷中为图片（选项字母后内容为空白，如 "A.    B．    C．"），options 数组中每项填 "<<IMG>>"（系统在填充时会自动替换为默认图片文件名）。

【★ 符号处理规则（重要）】
- **candidates 字段**（听后应答候选项）：去掉每条候选项文字**最前面**的 ★ 等标注符号（如 "★In the zoo." → "In the zoo."）。
- **其他所有字段**（question、listening_script、blanks[].question 等）：**严禁删除** ★ 符号，原文有 ★ 就保留 ★。

【标点符号规则】英文内容（听力原文、题干、答案等）中请统一使用半角 ASCII 标点（, . ? ! : ;），不要使用中文全角标点（如，。？！：；）；中文内容保持原格式不变。

每道题（每条）的 JSON 结构：
- type: 题型，如 "listening_choice" | "listening_response" | "reading_aloud" | "listening_fill" | "listening_fill_and_retell" 等；其中 `reading_aloud` 对应录题页里的「交际朗读」
- listening_script: **仅听力材料正文**，填到页面的「听力原文」输入框；无听力则 ""。若题目材料里提供了对应原文，即使页面该字段标注"非必填"也要输出，不要省略。对话型原文（W:/M: 格式）各说话人之间用 \n 换行分隔，不要写成一行。多小题（blanks）时按以下规则处理：① 若多小题**共用同一段**对话/段落（最常见情况）：把完整对话放在**顶层 listening_script**，各 blank 的 listening_script 填该小题**专属**的短内容（如该小题播放的提问句，例如 "What's wrong with Lily?"），若无专属内容则填 ""——**严禁把共享对话复制到每个 blank**；② 若每小题有**完全独立**的对话/段落：顶层 listening_script 留 ""，各 blank 填各自对应的完整原文，不截断。
- question: 题干，必须是**试卷上印刷的原始内容**（学生看到的题目文字），不要自己编造任务描述；若题干需要上传图片则填 ""。**若题干内容开头带有题号（如 1. 2. 一、(1) （1）等），请去掉题号，只保留题干正文。****题干应包含所有给学生看的文字，包括题目要求、提示内容（如"提示：1. 2. 3."）、参考起始句等，完整保留，不要截断或省略。**
- keyword: 仅在录题页存在「参考单词/送评单词」输入框时使用，填该题对应的送评单词/参考单词；无则 ""。
- options: 选项数组，无则 []
- candidates: **仅 listening_response（听后应答）题型使用**，所有候选答案的数组（去掉每条候选项最前面的★等标注符号），如 ["In the zoo.", "Mr Smith."]
- answer: 单题时的答案，**只填选项字母**（如 "A"、"B"、"C"），不要带选项文字；若该题有多空则用 blanks，不填 answer
- blanks: 仅当该题在录题页有多个答案框时使用，数组每项 {question, keyword, answer, options, listening_script}，options 为该小题的选项数组（无选项则 []），listening_script 为该小题对应的音频原文片段（详见下方听力原文规则），与页面顺序一致；**不要**把所有小题选项放到顶层 options 字段
- image_url: 若【录题页结构】标注本题「录入类型:image」或「包含：上传图片」，说明本题需要上传题干图片；若试卷中有对应的图片URL则填入，否则填 ""
- 填空类题型（如 listening_fill、listening_fill_and_retell）特别说明：
  - 若某个答案框的 placeholder 提示「各个答案之间用'X'分隔」，说明该框一次填多个答案，answer 用**数组**输出（如 ["red","potatoes","count"]），填充时会自动按 placeholder 的分隔符拼接
  - 若 placeholder 提示「单个答案包含多种解答形式用'|'分隔」，同一空有多个可接受答案时用数组元素内加'|'，如 ["red|Red","potatoes"]
  - 纯段落转述型的答案（单个长句/段落）直接用字符串
  - 多节合并为一道题时用 blanks 数组，每节一条；listening_script 填听力原文；顶层 question 填 ""
  - blanks[].question 必须是**试卷上印刷的原始内容**（如信息转述节的起始句），不要自己编造任务描述；若开头有题号也请去掉。
- explanation: **题目解析/详解**，填到页面「题目属性」里的「解析」框；不要与 listening_script 混淆，不要填到听力原文。**每条题目都必须包含 explanation 字段**：请结合题干内容和听力原文（如有）为该题生成解析/详解，不要省略该字段或填「略」。**多小题题目的解析中，引用各小题时用「第1小题」「第2小题」等相对编号**，不要用试卷原始的绝对题号（如第14题）——因为每道题录入后是独立的，绝对题号对学生无意义。

【多小题听力题「两层音频」结构说明（重要）】
录题平台中，多小题听力题（如听力选择、听后判断等）每个小题都有一个独立音频播放器（blanks[n].listening_script），大题本身也可以有一个共享音频（顶层 listening_script）。

请根据 Word 材料的结构判断如何分配：

情形 A：Word 中有**与小题数量相同的编号段落**（如 "1. W:... M:... 2. W:... M:... 3. W:..."），
  → 每段独立对话分别放入对应 blank 的 listening_script（去掉序号），顶层 listening_script 留 ""
  → 这是最常见情形，每小题播放各自独立的短对话

情形 B：Word 中只有**一段共享独白/对话**（无分段编号，所有小题围绕同一段材料提问），
  → 共享段落放入顶层 listening_script，blank.listening_script 填该小题的提问句（如有）或 ""
  → 例：一段人物独白 + 多个问题，问题各自作为小题的 listening_script

判断依据：若材料中有 "1." "2." "3." 等明确的分段序号且段数 ≥ 小题数 → 情形 A；
          若材料是一段连续文本（无对应小题的分段编号）→ 情形 B
          **严禁把多段内容合并成一段塞进 blank[0] 或顶层**

【听后应答题特殊说明】
- 听后应答（listening_response）题型：试题文件中的问题（如 "Who likes eagles?"）是听力音频会朗读的问题，放入 question 字段
- 答案材料中该题号对应的所有候选项（去掉每条最前面的★等符号）全部放入 candidates 数组
- answer 字段：优先从文档中查找该题的正确答案；若文档未明确标注，则根据问题从 candidates 中选择最合理的一个作为答案

【交际朗读题特殊说明】
- 录题页若显示为「交际朗读」，type 仍输出 "reading_aloud"
- question 填试卷上该题**专属**的内容：实际要朗读的短文/句子/单词，或该题特有的任务要求（背景情境、提示词、参考开头句等）。这些内容因题而异，是学生需要看到才能完成作答的文字，都应保留。**若录题页「设置题干」输入框的 placeholder 提示包含「音标」「单词+音标」，question 须填「单词 /音标/」，如 "deal /diːl/"；若提示仅为「单词」则只填单词；若提示为「句子」则只填句子。以页面实际 placeholder 为准。**
- **严禁**把「通用题型操作说明」填入 question。通用说明指每道同类题都完全相同的模板文字，例如："听一遍下面的短文，之后你有X秒钟的准备时间。请在听到录音提示后X秒钟内完成朗读/回答。现在你有X秒钟的准备时间。"——这类文字只是考试流程说明，与具体题目内容无关，**不得**填入 question。
- 判断方法：把这段文字原封不动放到另一道同类题，是否照样适用？若适用 → 通用说明（不填）；若必须配合本题内容才有意义 → 该题要求（填入）。
- 若 question 排除通用操作说明后确无文字内容，则填 ""（朗读内容将通过上传图片另行提供）。
- **重要**：排除通用说明后，若原材料中仍有学生需要实际朗读的内容（短文段落、单词、句子等），必须完整填入 question，不得遗漏、不得留空、不得用套话替代。仅当真的没有任何文字朗读内容时才填 ""。
- keyword 填「参考单词/送评单词」框需要的内容（**仅填单词本身，不含音标**，该字段只用于系统评分送审，不展示给学生）；若材料未提供则填 ""

【听后记录并转述题特殊说明】
- 题型 listening_fill_and_retell 含两部分：第一节「信息记录」（表格填空）和第二节「信息转述」（口头转述）。
- 输出1条，type 填 "listening_fill_and_retell"，用 blanks 数组：第一节各空每空一项（answer 为填空答案），第二节一项（question 填转述起始句原文，answer 填转述参考答案）。
- listening_script 填完整听力原文。

只输出 JSON 数组，不要 markdown 代码块，不要其他说明。"""

# 单文件解析：同样通用，目标结构由 paper_metadata 注入
SINGLE_FILE_PROMPT = """你是一个题目解析助手。请解析试卷内容，按**录题页结构**输出 JSON。

【通用规则】
- 若下方有【录题页结构】/【目标题数】：以之为准；某题多空时只输出一条，用 blanks，顺序与页面一致。
- 若未提供录题页结构：按文档题号逐条提取；同一表格/同一材料下多空可合并为一条并用 blanks。

【内嵌答案规则】答案可能以 (A)/(B)/(C) 形式写在题目行首（如 "(C)1.A.  B．  C．"），请提取括号内字母作为 answer，去掉行首的 (X) 前缀，只保留题干正文和选项内容。

【图片选项规则】若某题的选项在原卷中为图片（选项字母后内容为空白，如 "A.    B．    C．"），options 数组中每项填 "<<IMG>>"（系统在填充时会自动替换为默认图片文件名）。

【★ 符号处理规则（重要）】
- **candidates 字段**（听后应答候选项）：去掉每条候选项文字**最前面**的 ★ 等标注符号。
- **其他所有字段**（question、listening_script、blanks[].question 等）：**严禁删除** ★ 符号，原文有 ★ 就保留 ★。

【标点符号规则】英文内容中请统一使用半角 ASCII 标点（, . ? ! : ;），不要使用中文全角标点（，。？！：；）；中文内容保持原格式。

每道题的 JSON 结构：
- type: 题型，如 "listening_choice" | "listening_response" | "reading_aloud" | "listening_fill" | "listening_fill_and_retell" 等；其中 `reading_aloud` 对应录题页里的「交际朗读」
- listening_script: 仅听力材料正文，填「听力原文」框；无则 ""。若题目材料里提供了对应原文，即使页面该字段标注"非必填"也要输出，不要省略。对话型原文（W:/M: 格式）各说话人之间用 \n 换行分隔，不要写成一行。多小题（blanks）时按以下规则处理：① 若多小题共用同一段对话（最常见）：完整对话放**顶层 listening_script**，各 blank 的 listening_script 填该小题专属短内容（如该小题的提问句），无则 ""，**不要把共享对话复制到每个 blank**；② 若每小题有完全独立的对话：顶层留 ""，各 blank 填各自对应完整原文，不截断。
- question: 题干，必须是**试卷上印刷的原始内容**，不要自己编造任务描述；题干需要上传图片时填 ""。**若题干内容开头带有题号（如 1. 2. 一、(1) （1）等），请去掉题号，只保留题干正文。****题干应包含所有给学生看的文字，包括题目要求、提示内容（如"提示：1. 2. 3."）、参考起始句等，完整保留，不要截断或省略。**
- keyword: 仅在录题页存在「参考单词/送评单词」输入框时使用，填该题对应的送评单词/参考单词；无则 ""。
- options: 选项数组，无则 []
- candidates: **仅 listening_response（听后应答）题型使用**，所有候选答案的数组（去掉每条候选项最前面的★等标注符号），如 ["In the zoo.", "Mr Smith."]
- answer: 单题答案，**只填选项字母**（如 "A"）；多空时用 blanks
- blanks: 该题多空时使用 [{question, keyword, answer, options, listening_script}]，question 必须是**试卷上印刷的原始内容**（如信息转述起始句），若开头有题号也请去掉；options 为该小题选项数组（无则 []）；listening_script 为该小题对应的音频原文片段（详见下方听力原文规则）；**不要**把选项放到顶层 options 字段
- image_url: 若【录题页结构】标注本题需要上传图片，且试卷中有对应图片URL则填入，否则填 ""
- 填空类题型：若答案框 placeholder 提示「用'X'分隔」，answer 用数组（如 ["red","potatoes","count"]）
- explanation: 题目解析，填「题目属性」里的「解析」框；不要填到听力原文。**每条题目都必须包含 explanation 字段**：请结合题干内容和听力原文（如有）为该题生成解析/详解，不要省略该字段或填「略」。**多小题题目的解析中，用「第1小题」「第2小题」等相对编号**，不要用试卷原始的绝对题号（如第14题）。

【多小题听力题「两层音频」结构说明（重要）】
录题平台中，多小题听力题（如听力选择、听后判断等）每个小题都有一个独立音频播放器（blanks[n].listening_script），大题本身也可以有一个共享音频（顶层 listening_script）。

请根据 Word 材料的结构判断如何分配：

情形 A：Word 中有**与小题数量相同的编号段落**（如 "1. W:... M:... 2. W:... M:... 3. W:..."），
  → 每段独立对话分别放入对应 blank 的 listening_script（去掉序号），顶层 listening_script 留 ""
  → 这是最常见情形，每小题播放各自独立的短对话

情形 B：Word 中只有**一段共享独白/对话**（无分段编号，所有小题围绕同一段材料提问），
  → 共享段落放入顶层 listening_script，blank.listening_script 填该小题的提问句（如有）或 ""
  → 例：一段人物独白 + 多个问题，问题各自作为小题的 listening_script

判断依据：若材料中有 "1." "2." "3." 等明确的分段序号且段数 ≥ 小题数 → 情形 A；
          若材料是一段连续文本（无对应小题的分段编号）→ 情形 B
          **严禁把多段内容合并成一段塞进 blank[0] 或顶层**

【听后应答题特殊说明】
- 听后应答（listening_response）题型：问题（如 "Who likes eagles?"）放入 question 字段
- 该题号对应的所有候选项（去掉每条最前面的★等符号）全部放入 candidates 数组
- answer 字段：优先从文档中查找该题的正确答案；若文档未明确标注，则根据问题从 candidates 中选择最合理的一个作为答案

【交际朗读题特殊说明】
- 录题页若显示为「交际朗读」，type 仍输出 "reading_aloud"
- question 填试卷上该题**专属**的内容：实际要朗读的短文/句子/单词，或该题特有的任务要求（背景情境、提示词、参考开头句等）。这些内容因题而异，是学生需要看到才能完成作答的文字，都应保留。**若录题页「设置题干」输入框的 placeholder 提示包含「音标」「单词+音标」，question 须填「单词 /音标/」，如 "deal /diːl/"；若提示仅为「单词」则只填单词；若提示为「句子」则只填句子。以页面实际 placeholder 为准。**
- **严禁**把「通用题型操作说明」填入 question。通用说明指每道同类题都完全相同的模板文字，例如："听一遍下面的短文，之后你有X秒钟的准备时间。请在听到录音提示后X秒钟内完成朗读/回答。现在你有X秒钟的准备时间。"——这类文字只是考试流程说明，与具体题目内容无关，**不得**填入 question。
- 判断方法：把这段文字原封不动放到另一道同类题，是否照样适用？若适用 → 通用说明（不填）；若必须配合本题内容才有意义 → 该题要求（填入）。
- 若 question 排除通用操作说明后确无文字内容，则填 ""（朗读内容将通过上传图片另行提供）。
- **重要**：排除通用说明后，若原材料中仍有学生需要实际朗读的内容（短文段落、单词、句子等），必须完整填入 question，不得遗漏、不得留空、不得用套话替代。仅当真的没有任何文字朗读内容时才填 ""。
- keyword 填「参考单词/送评单词」框需要的内容（**仅填单词本身，不含音标**，该字段只用于系统评分送审，不展示给学生）；若材料未提供则填 ""

【听后记录并转述题特殊说明】
- 题型 listening_fill_and_retell 含两部分：第一节「信息记录」（表格填空）和第二节「信息转述」（口头转述）。
- 输出1条，type 填 "listening_fill_and_retell"，用 blanks 数组：第一节各空每空一项（answer 为填空答案），第二节一项（question 填转述起始句原文，answer 填转述参考答案）。
- listening_script 填完整听力原文。

只输出 JSON 数组，不要 markdown 代码块，不要其他说明。"""


def _system_prompt_with_total(base_prompt: str, paper_metadata: Dict[str, Any] | None) -> str:
    """根据「页面检测」得到的题数、每题空数，在 base prompt 前注入「目标结构」说明，供 AI 按页面对应输出。
    
    优化原则：
    1. 精简描述，只告诉 AI 必要信息
    2. 相同结构的小题合并描述，不逐个列出
    3. 图片选项只需告知「选项为图片」，AI 返回空时由系统用默认图片填充
    """
    if not paper_metadata:
        return base_prompt
    total = paper_metadata.get("total_questions")
    n = int(total) if total is not None and isinstance(total, (int, float)) and int(total) > 0 else 0
    slots = paper_metadata.get("question_slots")
    hint = ""
    if n > 0:
        hint = f"""【目标题数】录题页共 {n} 题。请严格只输出 {n} 条，按题号顺序一一对应。
"""
    if slots and isinstance(slots, list):
        parts = []
        for i, s in enumerate(slots[: (n or len(slots))]):
            idx = s.get("index", i + 1)
            sub = s.get("subCount", 1)
            try:
                sub = int(sub) if sub is not None else 1
            except (TypeError, ValueError):
                sub = 1
            
            # 从 sectionLabels 推断题型（最可靠）
            section_labels = s.get("sectionLabels") or []
            labels_set = set(section_labels) if isinstance(section_labels, list) else set()
            
            # 推断题型
            inferred_type = ""
            if "参考单词" in labels_set and "题干" in labels_set:
                inferred_type = "交际朗读"
            elif "听力原文" in labels_set and ("设置选项" in labels_set or "图片选项" in labels_set):
                inferred_type = "听后选择"
            elif "听力原文" in labels_set and "题干" in labels_set and "设置答案" in labels_set and "设置选项" not in labels_set:
                inferred_type = "听后记录并转述"
            elif "听力原文" in labels_set and "设置选项" not in labels_set and "设置答案" not in labels_set:
                inferred_type = "交际朗读"
            elif "设置选项" in labels_set and "听力原文" not in labels_set:
                inferred_type = "听后应答"
            elif "设置答案" in labels_set and "设置选项" not in labels_set:
                inferred_type = "填空/转述"
            
            # 使用页面提供的 typeHint 或推断的类型
            type_hint = (s.get("typeHint") or "").strip() or (s.get("typeCode") or "").strip() or inferred_type
            
            # 检测选项类型
            option_kind = (s.get("optionKind") or "").strip()
            is_image_option = option_kind == "image"
            
            # 构建简洁描述
            desc_parts = []
            if type_hint:
                desc_parts.append(type_hint)
            
            # 多小题处理
            if sub > 1:
                # 检查小题结构是否相同
                sub_questions = s.get("subQuestions") or []
                all_same_structure = True
                first_sq_labels = None
                for sq in sub_questions[:sub]:
                    sq_labels = set(sq.get("sectionLabels") or [])
                    if first_sq_labels is None:
                        first_sq_labels = sq_labels
                    elif sq_labels != first_sq_labels:
                        all_same_structure = False
                        break
                
                if all_same_structure and first_sq_labels:
                    # 小题结构相同，合并描述
                    desc_parts.append(f"共{sub}小题，结构相同")
                    # 列出每小题包含的所有关键组件（含题干/参考单词）
                    sq_components = [l for l in ["上传音频", "听力原文", "题干", "参考单词", "设置选项", "设置答案"] if l in first_sq_labels]
                    if sq_components:
                        desc_parts.append(f"每小题含：{'/'.join(sq_components)}")
                else:
                    desc_parts.append(f"共{sub}小题，一题多空")
                    # 只有结构不同时才逐个描述
                    if not all_same_structure and sub_questions:
                        sq_descs = []
                        for sq in sub_questions[:sub]:
                            sq_idx = sq.get("index", len(sq_descs) + 1)
                            sq_heading = (sq.get("heading") or "").strip()
                            if sq_heading:
                                # 简化标题，只取关键词
                                if "信息记录" in sq_heading:
                                    sq_descs.append(f"第{sq_idx}小题:信息记录")
                                elif "信息转述" in sq_heading:
                                    sq_descs.append(f"第{sq_idx}小题:信息转述")
                        if sq_descs:
                            desc_parts.append("；".join(sq_descs))
            
            # 图片选项：简洁告知
            if is_image_option:
                desc_parts.append("选项为图片(填\"<<IMG>>\")")
            
            # 大题级别的组件
            main_components = []
            if sub <= 1:
                # 单题：完整列出所有检测到的字段，让 AI 知道需要填哪些内容
                _FIELD_ORDER = [
                    ("上传音频",     "上传音频"),
                    ("听力原文",     "听力原文"),
                    ("题干",         "题干"),
                    ("参考单词",     "参考单词"),
                    ("设置选项",     "设置选项"),
                    ("设置答案",     "设置答案"),
                    ("上传图片",     "上传图片"),
                    ("题目属性信息", "题目属性"),
                    ("解析",         "解析"),
                ]
                main_components = [disp for key, disp in _FIELD_ORDER if key in labels_set]
            else:
                # 多小题：小题内已描述核心字段，大题层面只补充附加字段
                if "题目属性信息" in labels_set:
                    main_components.append("题目属性")
                if "解析" in labels_set:
                    main_components.append("解析")
                if "上传图片" in labels_set:
                    main_components.append("上传图片")
            if main_components:
                desc_parts.append(f"含：{'/'.join(main_components)}")
            
            desc = f"序号{idx}"
            if desc_parts:
                desc += "（" + "，".join(desc_parts) + "）"
            
            parts.append(desc)
        
        if parts:
            hint += "【录题页结构】" + "；".join(parts) + "。\n"
            hint += "【重要】多小题的题只输出一条，用 blanks 数组存放各小题的 {question, keyword, answer, options, listening_script}。\n"
    
    if not hint:
        return base_prompt
    return hint + "\n" + base_prompt


async def parse_exam_with_answers(
    exam_text: str,
    answer_text: str,
    *,
    base_url: str,
    api_key: str,
    model: str,
    field_structure: List[Dict[str, Any]] | None = None,
    paper_metadata: Dict[str, Any] | None = None,
    return_debug: bool = False,
    reasoning_effort: str = "medium",
):
    """将试题和答案材料合并解析，按题号对齐。
    
    exam_text: 试题文件内容
    answer_text: 答案材料文件内容
    paper_metadata: 可含 total_questions，用于约束只提取对应题数
    return_debug: 为 True 时返回 (questions, debug_info)，便于调试 prompt
    reasoning_effort: 思考程度 minimal/low/medium/high
    """
    key = _resolve_api_key(api_key, base_url)
    if not key:
        raise ValueError("未配置 api_key")
    
    # 用占位符保护 ★，防止 AI 误删题干中的标记
    user_content = f"""【试题】
{_encode_star(exam_text)}

【答案材料】
{_encode_star(answer_text)}

请按题号合并解析，输出 JSON 数组。"""
    system_content = _system_prompt_with_total(MERGED_PARSE_PROMPT, paper_metadata)
    client = AsyncOpenAI(base_url=base_url, api_key=key)
    kwargs = {"model": model, "messages": [{"role": "system", "content": system_content}, {"role": "user", "content": user_content}], "temperature": 0.1}
    if _is_volc(base_url):
        kwargs["extra_body"] = _volc_extra_body(reasoning_effort)
    resp = await client.chat.completions.create(**kwargs)
    content = resp.choices[0].message.content or ""
    questions = _normalize_questions(_decode_star_questions(extract_json_from_response(content)))
    if return_debug:
        return questions, {
            "system_prompt": system_content,
            "user_content": user_content,
            "strategy": "exam_with_answers",
        }
    return questions


# ─── 单文件解析 Prompt ─────────────────────────────────────────────────────────



async def parse_single_file(
    text: str,
    *,
    base_url: str,
    api_key: str,
    model: str,
    field_structure: List[Dict[str, Any]] | None = None,
    paper_metadata: Dict[str, Any] | None = None,
    return_debug: bool = False,
    reasoning_effort: str = "medium",
):
    """解析单个文件的题目。paper_metadata 可含 total_questions，用于约束只提取对应题数。
    return_debug 为 True 时返回 (questions, debug_info)，便于调试 prompt。"""
    key = _resolve_api_key(api_key, base_url)
    if not key:
        raise ValueError("未配置 api_key")
    
    encoded_text = _encode_star(text)
    system_content = _system_prompt_with_total(SINGLE_FILE_PROMPT, paper_metadata)
    client = AsyncOpenAI(base_url=base_url, api_key=key)
    kwargs = {"model": model, "messages": [{"role": "system", "content": system_content}, {"role": "user", "content": encoded_text}], "temperature": 0.1}
    if _is_volc(base_url):
        kwargs["extra_body"] = _volc_extra_body(reasoning_effort)
    resp = await client.chat.completions.create(**kwargs)
    content = resp.choices[0].message.content or ""
    questions = _normalize_questions(_decode_star_questions(extract_json_from_response(content)))
    if return_debug:
        return questions, {
            "system_prompt": system_content,
            "user_content": encoded_text,
            "strategy": "single_file",
        }
    return questions
