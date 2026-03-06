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
        for field in ("question", "listening_script", "explanation"):
            if field in q:
                q[field] = restore(q[field])
        # candidates：去掉占位符（原来的 ★ 标记正确答案，不需要显示）
        if "candidates" in q and isinstance(q["candidates"], list):
            q["candidates"] = [strip_star(c) for c in q["candidates"]]
        # blanks
        if "blanks" in q and isinstance(q["blanks"], list):
            for blank in q["blanks"]:
                for field in ("question", "answer"):
                    if field in blank:
                        blank[field] = restore(blank[field])
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
                desc = f"第{idx}题"
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
    
    # 构建动态字段说明
    field_descriptions = []
    
    # 根据检测到的字段生成说明
    role_to_desc = {
        "type": ('type', '题型，如 "single"（单选）、"multiple"（多选）、"judge"（判断）、"blank"（填空）等'),
        "question": ('question', '题干文字'),
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
        if role in role_to_desc:
            key, desc = role_to_desc[role]
            label = detected_labels.get(role, "")
            if label and label != key:
                field_descriptions.append(f"- {key}: {desc}（页面字段：{label}）")
            else:
                field_descriptions.append(f"- {key}: {desc}")
        else:
            # 未知字段也尝试提取
            label = detected_labels.get(role, role)
            field_descriptions.append(f"- {role}: 对应页面字段「{label}」的内容")
    
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


def _volc_extra_body() -> dict:
    """豆包：开启深度思考，级别 medium，兼顾推理质量与耗时。"""
    return {"thinking": {"type": "enabled", "level": "medium"}}


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
) -> List[Dict[str, Any]]:
    """将一块题目文本发给 LLM，返回解析后的题目列表。"""
    key = _resolve_api_key(api_key, base_url)
    if not key:
        raise ValueError("未配置 api_key，请在 config.yaml 的 llm.api_key 填写或设置环境变量 ARK_API_KEY")
    system_prompt = build_system_prompt(field_structure)
    client = OpenAI(base_url=base_url, api_key=key)
    kwargs = {"model": model, "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": text_chunk}], "temperature": 0.1}
    if _is_volc(base_url):
        kwargs["extra_body"] = _volc_extra_body()
    resp = client.chat.completions.create(**kwargs)
    content = resp.choices[0].message.content or ""
    return extract_json_from_response(content)


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
    return all_questions


async def _call_llm_extract_async(
    text_chunk: str,
    *,
    base_url: str,
    api_key: str,
    model: str,
    semaphore: asyncio.Semaphore,
    system_prompt: str,
) -> List[Dict[str, Any]]:
    """异步调用 LLM，受 semaphore 控制并发数。"""
    key = _resolve_api_key(api_key, base_url)
    async with semaphore:
        client = AsyncOpenAI(base_url=base_url, api_key=key)
        kwargs = {"model": model, "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": _encode_star(text_chunk)}], "temperature": 0.1}
        if _is_volc(base_url):
            kwargs["extra_body"] = _volc_extra_body()
        resp = await client.chat.completions.create(**kwargs)
    content = resp.choices[0].message.content or ""
    return _decode_star_questions(extract_json_from_response(content))


async def extract_questions_from_word_chunks_async(
    chunks: List[str],
    *,
    base_url: str,
    api_key: str,
    model: str,
    max_concurrency: int = 5,
    field_structure: List[Dict[str, Any]] | None = None,
    paper_metadata: Dict[str, Any] | None = None,
) -> List[Dict[str, Any]]:
    """并发调用 LLM 解析所有文本块，最多同时 max_concurrency 个请求。
    结果按原始顺序合并，失败的块跳过并打印警告。
    
    field_structure: 页面检测到的字段结构，用于生成动态 prompt
    paper_metadata: 试卷元信息（总题数、题型分布等），用于指导精确解析
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
    return all_questions


# ─── 快速文件类型识别（不调 LLM，根据文件名和内容特征）─────────────────────

def classify_file_fast(filename: str, text: str) -> Dict[str, Any]:
    """快速识别文件类型，不调用 LLM，根据文件名关键词和内容特征判断。
    
    返回：{"file_type": "exam|answer_material|other", "has_questions": bool, "has_answers": bool, "has_script": bool}
    """
    fn = filename.lower()
    
    # 内容特征检测
    has_options = bool(re.search(r'[A-D]\s*[.、．]\s*\w', text[:3000]))  # 有选项 A. B. C.
    has_answers = '参考答案' in text or bool(re.search(r'\d+[-~]\d+\s*[A-D]', text))  # 1-4 CAAC 格式
    has_script = '听力材料' in text or '听力原文' in text or bool(re.search(r'[MW][：:]\s*\w', text[:2000]))
    has_question_nums = bool(re.search(r'[(（]\s*[)）]\s*\d+\.', text[:1500]))  # (  )1. 格式
    
    # 根据文件名判断（优先级最高）
    # 注意：文件名同时包含"试题"和"答案"时，优先看更具体的关键词
    fn_has_exam = '试题' in fn or '试卷' in fn
    fn_has_answer = '答案' in fn or '材料' in fn
    
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

【★ 符号处理规则（重要）】
- **candidates 字段**（听后应答候选项）：去掉每条候选项文字**最前面**的 ★ 等标注符号（如 "★In the zoo." → "In the zoo."）。
- **其他所有字段**（question、listening_script、blanks[].question 等）：**严禁删除** ★ 符号，原文有 ★ 就保留 ★。

每道题（每条）的 JSON 结构：
- type: 题型，如 "listening_choice" | "listening_response" | "reading_aloud" | "listening_fill" | "listening_fill_and_retell" 等
- listening_script: **仅听力材料正文**，填到页面的「听力原文」输入框；无听力则 ""。**重要：同一段听力对话对应多道小题时，每道小题的 listening_script 必须填写完整的同一段对话原文，绝对不能截断或拆分，所有小题的 listening_script 必须完全一致。**
- question: 题干，必须是**试卷上印刷的原始内容**（学生看到的题目文字），不要自己编造任务描述；若题干需要上传图片则填 ""
- options: 选项数组，无则 []
- candidates: **仅 listening_response（听后应答）题型使用**，所有候选答案的数组（去掉每条候选项最前面的★等标注符号），如 ["In the zoo.", "Mr Smith."]
- answer: 单题时的答案，**只填选项字母**（如 "A"、"B"、"C"），不要带选项文字；若该题有多空则用 blanks，不填 answer
- blanks: 仅当该题在录题页有多个答案框时使用，数组每项 {question, answer, options}，options 为该小题的选项数组（无选项则 []），与页面顺序一致；**不要**把所有小题选项放到顶层 options 字段
- image_url: 若【录题页结构】标注本题「录入类型:image」或「包含：上传图片」，说明本题需要上传题干图片；若试卷中有对应的图片URL则填入，否则填 ""
- 填空类题型（如 listening_fill、listening_fill_and_retell）特别说明：
  - 若某个答案框的 placeholder 提示「各个答案之间用'X'分隔」，说明该框一次填多个答案，answer 用**数组**输出（如 ["red","potatoes","count"]），填充时会自动按 placeholder 的分隔符拼接
  - 若 placeholder 提示「单个答案包含多种解答形式用'|'分隔」，同一空有多个可接受答案时用数组元素内加'|'，如 ["red|Red","potatoes"]
  - 纯段落转述型的答案（单个长句/段落）直接用字符串
  - 多节合并为一道题时用 blanks 数组，每节一条；listening_script 填听力原文；顶层 question 填 ""
  - blanks[].question 必须是**试卷上印刷的原始内容**（如信息转述节的起始句），不要自己编造任务描述
- explanation: **题目解析/详解**，填到页面「题目属性」里的「解析」框；不要与 listening_script 混淆，不要填到听力原文。**每条题目都必须包含 explanation 字段**：请结合题干内容和听力原文（如有）为该题生成解析/详解，不要省略该字段或填「略」。

【听后应答题特殊说明】
- 听后应答（listening_response）题型：试题文件中的问题（如 "Who likes eagles?"）是听力音频会朗读的问题，放入 question 字段
- 答案材料中该题号对应的所有候选项（去掉每条最前面的★等符号）全部放入 candidates 数组
- answer 字段：优先从文档中查找该题的正确答案；若文档未明确标注，则根据问题从 candidates 中选择最合理的一个作为答案

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

【★ 符号处理规则（重要）】
- **candidates 字段**（听后应答候选项）：去掉每条候选项文字**最前面**的 ★ 等标注符号。
- **其他所有字段**（question、listening_script、blanks[].question 等）：**严禁删除** ★ 符号，原文有 ★ 就保留 ★。

每道题的 JSON 结构：
- type: 题型，如 "listening_choice" | "listening_response" | "reading_aloud" | "listening_fill" | "listening_fill_and_retell" 等
- listening_script: 仅听力材料正文，填「听力原文」框；无则 ""。**重要：同一段听力对话对应多道小题时，每道小题的 listening_script 必须填写完整的同一段对话原文，绝对不能截断或拆分，所有小题的 listening_script 必须完全一致。**
- question: 题干，必须是**试卷上印刷的原始内容**，不要自己编造任务描述；题干需要上传图片时填 ""
- options: 选项数组，无则 []
- candidates: **仅 listening_response（听后应答）题型使用**，所有候选答案的数组（去掉每条候选项最前面的★等标注符号），如 ["In the zoo.", "Mr Smith."]
- answer: 单题答案，**只填选项字母**（如 "A"）；多空时用 blanks
- blanks: 该题多空时使用 [{question, answer, options}]，question 必须是**试卷上印刷的原始内容**（如信息转述起始句）；options 为该小题选项数组（无则 []）；**不要**把选项放到顶层 options 字段
- image_url: 若【录题页结构】标注本题需要上传图片，且试卷中有对应图片URL则填入，否则填 ""
- 填空类题型：若答案框 placeholder 提示「用'X'分隔」，answer 用数组（如 ["red","potatoes","count"]）
- explanation: 题目解析，填「题目属性」里的「解析」框；不要填到听力原文。**每条题目都必须包含 explanation 字段**：请结合题干内容和听力原文（如有）为该题生成解析/详解，不要省略该字段或填「略」。

【听后应答题特殊说明】
- 听后应答（listening_response）题型：问题（如 "Who likes eagles?"）放入 question 字段
- 该题号对应的所有候选项（去掉每条最前面的★等符号）全部放入 candidates 数组
- answer 字段：优先从文档中查找该题的正确答案；若文档未明确标注，则根据问题从 candidates 中选择最合理的一个作为答案

【听后记录并转述题特殊说明】
- 题型 listening_fill_and_retell 含两部分：第一节「信息记录」（表格填空）和第二节「信息转述」（口头转述）。
- 输出1条，type 填 "listening_fill_and_retell"，用 blanks 数组：第一节各空每空一项（answer 为填空答案），第二节一项（question 填转述起始句原文，answer 填转述参考答案）。
- listening_script 填完整听力原文。

只输出 JSON 数组，不要 markdown 代码块，不要其他说明。"""


def _system_prompt_with_total(base_prompt: str, paper_metadata: Dict[str, Any] | None) -> str:
    """根据「页面检测」得到的题数、每题空数，在 base prompt 前注入「目标结构」说明，供 AI 按页面对应输出。"""
    if not paper_metadata:
        return base_prompt
    total = paper_metadata.get("total_questions")
    n = int(total) if total is not None and isinstance(total, (int, float)) and int(total) > 0 else 0
    slots = paper_metadata.get("question_slots")
    hint = ""
    if n > 0:
        hint = f"""【目标题数】录题页共 {n} 题。题号按页面左侧顺序（可能为大题或小题编号），请严格只输出 {n} 条、按该顺序一一对应，不重复、不遗漏、不多提。
"""
    if slots and isinstance(slots, list):
        parts = []
        for i, s in enumerate(slots[: (n or len(slots))]):
            idx = s.get("index", i + 1)
            part = (s.get("partName") or "").strip()
            type_hint = (s.get("typeHint") or "").strip()
            type_code = (s.get("typeCode") or "").strip()
            sub = s.get("subCount", 1)
            try:
                sub = int(sub) if sub is not None else 1
            except (TypeError, ValueError):
                sub = 1
            input_count = s.get("inputCount") or {}
            answer_count = input_count.get("answer", sub)
            try:
                answer_count = int(answer_count) if answer_count is not None else sub
            except (TypeError, ValueError):
                answer_count = sub
            question_count = input_count.get("question", 0)
            try:
                question_count = int(question_count) if question_count is not None else 0
            except (TypeError, ValueError):
                question_count = 0
            option_count = s.get("optionCount", 0)
            try:
                option_count = int(option_count) if option_count is not None else 0
            except (TypeError, ValueError):
                option_count = 0
            media = s.get("media") or {}
            input_kind = (media.get("inputKind") or "text").strip() or "text"

            desc_parts = []
            if part:
                desc_parts.append(part)

            # 从 sectionLabels 推断题型描述（比 typeHint 更可靠，直接来自页面区块标题）
            section_labels = s.get("sectionLabels") or []
            inferred_type = ""
            if isinstance(section_labels, list) and len(section_labels) > 0:
                labels_set = set(section_labels)
                if "听力原文" in labels_set and "设置选项" in labels_set:
                    inferred_type = "听后选择" if sub == 1 else "听后选择（多小题）"
                elif "听力原文" in labels_set and "题干" in labels_set and "设置答案" in labels_set and "设置选项" not in labels_set:
                    # 有听力原文 + 题干 + 设置答案，且无选项 → 听后记录并转述
                    inferred_type = "听后记录并转述"
                elif "听力原文" in labels_set and "设置选项" not in labels_set and "设置答案" not in labels_set:
                    inferred_type = "模仿朗读"
                elif "设置选项" in labels_set and "听力原文" not in labels_set:
                    inferred_type = "听后应答"
                elif "设置答案" in labels_set and "设置选项" not in labels_set:
                    inferred_type = "填空/转述"

            if type_hint or type_code:
                desc_parts.append(type_hint or type_code)
            elif inferred_type:
                desc_parts.append(inferred_type)

            # 用 subCount 作为答案框数（最准确），不显示选项数（与题数无关）
            if sub > 1:
                desc_parts.append(f"共{sub}小题")
                desc_parts.append("一题多空")
            else:
                desc_parts.append("答案框1个")
            if input_kind and input_kind != "text":
                desc_parts.append(f"录入类型:{input_kind}")
            # 明确标注是否需要图片/音频
            has_image = media.get("hasImage")
            has_audio = media.get("hasAudio")
            if has_image and input_kind != "image":
                desc_parts.append("需要上传图片")
            if has_audio and input_kind != "audio":
                desc_parts.append("需要上传音频")
            # 页面扫描到的区块标题，直接告诉大模型本题有哪些组件
            if isinstance(section_labels, list) and len(section_labels) > 0:
                desc_parts.append("包含：" + "、".join(section_labels))

            desc = f"第{idx}题"
            if desc_parts:
                desc += "（" + "，".join(desc_parts) + "）"

            # 多小题时，追加每个小题的逐条描述（来自 getSubQuestionDetails 扫描结果）
            sub_questions = s.get("subQuestions") or []
            if sub > 1 and sub_questions:
                sq_desc_list = []
                for sq in sub_questions[:sub]:
                    sq_idx = sq.get("index", len(sq_desc_list) + 1)
                    sq_heading = (sq.get("heading") or "").strip()
                    sq_labels = sq.get("sectionLabels") or []

                    # 从标题推断小题类型
                    sq_type = ""
                    if "信息记录" in sq_heading:
                        sq_type = "信息记录（表格填空）"
                    elif "信息转述" in sq_heading:
                        sq_type = "信息转述（口头转述，blanks[].question 填转述起始句原文）"
                    elif sq_heading:
                        sq_type = sq_heading.split("（")[0].strip()[:20]

                    sq_parts = []
                    if sq_type:
                        sq_parts.append(sq_type)
                    if sq_labels:
                        sq_parts.append("包含：" + "、".join(sq_labels))

                    sq_desc = f"第{sq_idx}小题"
                    if sq_parts:
                        sq_desc += "（" + "，".join(sq_parts) + "）"
                    sq_desc_list.append(sq_desc)

                if sq_desc_list:
                    desc += "：" + "；".join(sq_desc_list)

            parts.append(desc)
        if parts:
            hint += "【录题页结构】" + "；".join(parts) + "。\n"
            # 若某题有题目 JSON/配置，附上供 AI 精确对应
            for i, s in enumerate(slots[: (n or len(slots))]):
                qj = s.get("questionJson")
                if qj is not None and (isinstance(qj, dict) or (isinstance(qj, str) and qj.strip())):
                    idx = s.get("index", i + 1)
                    hint += f"第{idx}题页面题目配置参考：" + (json.dumps(qj, ensure_ascii=False)[:500] if isinstance(qj, dict) else str(qj)[:500]) + "\n"
            hint += "重要：若某题标注「一题多空」或「共N小题」或答案框数>1，则输出中该题只占**一条**，且该条必须用 blanks 数组，包含相应个数的 {question, answer, options}，顺序与页面一致；**不要**把同一题的多个小题拆成多条输出。\n"
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
):
    """将试题和答案材料合并解析，按题号对齐。
    
    exam_text: 试题文件内容
    answer_text: 答案材料文件内容
    paper_metadata: 可含 total_questions，用于约束只提取对应题数
    return_debug: 为 True 时返回 (questions, debug_info)，便于调试 prompt
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
        kwargs["extra_body"] = _volc_extra_body()
    resp = await client.chat.completions.create(**kwargs)
    content = resp.choices[0].message.content or ""
    questions = _decode_star_questions(extract_json_from_response(content))
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
        kwargs["extra_body"] = _volc_extra_body()
    resp = await client.chat.completions.create(**kwargs)
    content = resp.choices[0].message.content or ""
    questions = _decode_star_questions(extract_json_from_response(content))
    if return_debug:
        return questions, {
            "system_prompt": system_content,
            "user_content": encoded_text,
            "strategy": "single_file",
        }
    return questions
