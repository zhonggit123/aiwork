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
    import re as _re
    # 说话人标记前插入换行：支持有/无空格（W:Was 和 W: Was 两种格式）
    # 匹配：W: M: Boy: Girl: Man: Woman: Narrator: Q: Qs: A: 以及常见中文/英文人名缩写
    formatted = _re.sub(
        r'(?<=[^\n])\s*(W:|M:|Boy:|Girl:|Man:|Woman:|Narrator:|Qs?:|Q\d+\.|[A-Z][a-z]{1,9}:)',
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


def _slot_type_to_json_type(slot: Dict[str, Any]) -> str:
    """从录题页槽位的 typeHint/推断类型 映射为 JSON 的 type 字段。"""
    section_labels = slot.get("sectionLabels") or []
    labels_set = set(section_labels) if isinstance(section_labels, list) else set()
    hint = (slot.get("typeHint") or "").strip() or (slot.get("typeCode") or "").strip()
    if "交际朗读" in hint or "模仿朗读" in hint or ("参考单词" in labels_set and "题干" in labels_set):
        return "reading_aloud"
    if "听后记录并转述" in hint or "信息记录" in str(slot.get("heading") or ""):
        return "listening_fill_and_retell"
    if "听后选择" in hint or ("设置选项" in labels_set and "听力原文" in labels_set):
        return "listening_choice"
    if "听后应答" in hint:
        return "listening_response"
    if "听后判断" in hint:
        return "listening_judge"
    if "填空" in hint or "信息记录" in hint:
        return "listening_fill"
    return ""


def _reorder_questions_by_slots(
    questions: List[Dict[str, Any]],
    slots: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """当题目数量与槽位数量一致时，按槽位顺序重排题目，使第 i 题与第 i 个槽位题型一致。"""
    if not questions or not slots or len(questions) != len(slots):
        return questions
    slot_types = [_slot_type_to_json_type(s) for s in slots[: len(questions)]]
    if not any(slot_types):
        return questions
    # 为每道题找最匹配的槽位：优先按 type 完全匹配
    used = set()
    result = []
    for st in slot_types:
        if not st:
            if len(result) < len(questions):
                for j, q in enumerate(questions):
                    if j not in used:
                        result.append(q)
                        used.add(j)
                        break
            continue
        for j, q in enumerate(questions):
            if j in used:
                continue
            qt = (q.get("type") or "").strip()
            if qt == st:
                result.append(q)
                used.add(j)
                break
        else:
            for j, q in enumerate(questions):
                if j not in used:
                    result.append(q)
                    used.add(j)
                    break
    # 未匹配到的题目按原顺序追加
    for j, q in enumerate(questions):
        if j not in used:
            result.append(q)
    return result if len(result) == len(questions) else questions


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
            top_script = (q.get("listening_script") or "").strip()
            if (blank_scripts
                    and all(s == blank_scripts[0] for s in blank_scripts)
                    and len(blank_scripts[0]) > 30):
                shared = blank_scripts[0]
                # 仅在顶层确实为空时才提升，避免覆盖已正确填入的顶层原文
                if not top_script:
                    q["listening_script"] = shared
                for blank in q["blanks"]:
                    if isinstance(blank, dict):
                        # 用 blank.question 作为小题专属脚本（如果它比共享脚本短很多）
                        bq = (blank.get("question") or "").strip()
                        blank["listening_script"] = bq if (bq and len(bq) < len(shared) * 0.5) else ""
            # 情形B兜底：顶层为空 + blank[0]的script明显比其他blank长得多，说明AI把共享对话塞进了blank[0]
            elif (not top_script
                  and blank_scripts
                  and len(blank_scripts[0]) > 60
                  and len(blank_scripts[0]) > sum(len(s) for s in blank_scripts[1:]) * 2):
                shared = blank_scripts[0]
                q["listening_script"] = shared
                for i, blank in enumerate(q["blanks"]):
                    if isinstance(blank, dict):
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
            # 兜底：AI 仍按"每空一项"输出多个只有 answer 无 question 的 blank（信息记录格式）
            # → 将连续的 fill-only blank 合并到第一个，answer 改为数组
            if has_blanks:
                blanks = q["blanks"]
                fill_idxs = [i for i, b in enumerate(blanks)
                             if isinstance(b, dict)
                             and not (b.get("question") or "").strip()
                             and (b.get("answer") not in (None, "", []))]
                retell_idxs = [i for i, b in enumerate(blanks)
                               if isinstance(b, dict)
                               and (b.get("question") or "").strip()]
                if len(fill_idxs) > 1:
                    # 把多个 fill blank 合并为数组答案放入 fill_idxs[0]
                    combined = [str(blanks[i].get("answer") or "").strip() for i in fill_idxs]
                    blanks[fill_idxs[0]]["answer"] = combined
                    # 删除多余的 fill blank（倒序删避免索引变化）
                    for i in sorted(fill_idxs[1:], reverse=True):
                        del blanks[i]
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


# 录题页题型名称 -> JSON type 映射（用于按槽位重排题目顺序）
_SLOT_TYPE_TO_JSON = {
    "交际朗读": "reading_aloud",
    "模仿朗读": "reading_aloud",
    "听后记录并转述": "listening_fill_and_retell",
    "听后选择": "listening_choice",
    "听后应答": "listening_response",
    "听后判断": "listening_judge",
    "听力填空": "listening_fill",
    "reading_aloud": "reading_aloud",
    "listening_fill_and_retell": "listening_fill_and_retell",
    "listening_choice": "listening_choice",
    "listening_response": "listening_response",
    "listening_judge": "listening_judge",
    "listening_fill": "listening_fill",
}


def _slot_expected_type(slot: Dict[str, Any]) -> str:
    """从录题页槽位推断期望的题目 type（JSON 字段值）。"""
    hint = (slot.get("typeHint") or slot.get("typeCode") or "").strip()
    if hint and hint in _SLOT_TYPE_TO_JSON:
        return _SLOT_TYPE_TO_JSON[hint]
    section_labels = set(slot.get("sectionLabels") or [])
    if "参考单词" in section_labels and "题干" in section_labels:
        return "reading_aloud"
    if "听力原文" in section_labels and "设置选项" in section_labels:
        return "listening_choice"
    if "听力原文" in section_labels and "题干" in section_labels and "设置答案" in section_labels and "设置选项" not in section_labels:
        return "listening_fill_and_retell"
    if "听力原文" in section_labels and "设置选项" not in section_labels and "设置答案" not in section_labels:
        return "reading_aloud"
    if "设置选项" in section_labels and "听力原文" not in section_labels:
        return "listening_response"
    return ""


def reorder_questions_by_slots(
    questions: List[Dict[str, Any]],
    slots: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """按录题页槽位顺序重排题目。当 LLM 返回顺序与文档不一致时，用槽位顺序矫正。
    要求 len(questions)==len(slots)，且每题能按 type 唯一匹配到槽位。
    """
    if not slots or not questions or len(slots) != len(questions):
        return questions
    expected = [_slot_expected_type(s) for s in slots]
    used = [False] * len(questions)
    out = []
    for exp in expected:
        for i, q in enumerate(questions):
            if used[i]:
                continue
            qtype = (q.get("type") or "").strip()
            if qtype == exp or (exp and not qtype):
                out.append(q)
                used[i] = True
                break
        else:
            # 未找到匹配，按剩余题目顺序取第一个
            for i in range(len(questions)):
                if not used[i]:
                    out.append(questions[i])
                    used[i] = True
                    break
    return out if len(out) == len(questions) else questions


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


def _repair_json(text: str) -> str:
    """简单修复 LLM 常见 JSON 错误：字符串内的裸换行替换为 \\n。"""
    # 用状态机扫描，将 JSON 字符串内出现的裸换行/制表符替换为转义形式
    result = []
    in_string = False
    escape_next = False
    for ch in text:
        if escape_next:
            result.append(ch)
            escape_next = False
        elif ch == "\\":
            result.append(ch)
            escape_next = True
        elif ch == '"':
            result.append(ch)
            in_string = not in_string
        elif in_string and ch == "\n":
            result.append("\\n")
        elif in_string and ch == "\r":
            result.append("\\r")
        elif in_string and ch == "\t":
            result.append("\\t")
        else:
            result.append(ch)
    return "".join(result)


def extract_json_from_response(content: str) -> List[Dict[str, Any]]:
    """从模型回复中剥离并解析 JSON 数组。"""
    content = content.strip()
    # 去掉可能的 ```json ... ``` 包裹
    if "```" in content:
        match = re.search(r"```(?:json)?\s*([\s\S]*?)```", content)
        if match:
            content = match.group(1).strip()

    def _try_parse(s: str):
        try:
            data = json.loads(s)
        except json.JSONDecodeError:
            data = json.loads(_repair_json(s))
        if isinstance(data, list):
            return data
        if isinstance(data, dict) and "questions" in data:
            return data["questions"]
        return [data]

    try:
        return _try_parse(content)
    except json.JSONDecodeError:
        # 尝试找到第一个 [ 到最后一个 ] 之间的内容
        start = content.find("[")
        end = content.rfind("]") + 1
        if start >= 0 and end > start:
            try:
                return _try_parse(content[start:end])
            except json.JSONDecodeError as e2:
                print(f"[JSON解析失败] {e2}\n--- LLM 原始回复（前3000字）---\n{content[:3000]}\n--- end ---", flush=True)
                raise
        print(f"[JSON解析失败] 未找到JSON数组\n--- LLM 原始回复（前3000字）---\n{content[:3000]}\n--- end ---", flush=True)
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
MERGED_PARSE_PROMPT = """你是题目解析助手。将【试题】和【答案材料】（或其一）按**录题页结构**解析为 JSON 数组。
文件分工：试题 → 题干/选项；答案材料 → 听力原文/参考答案（听后应答类题干可能在答案材料里，选项在试题上，按题号对齐合并）。

【输出约束】
- 顺序：严格按文档出现顺序，第1条对应页面第1题；不得按题型/相似度换位
- 边界：相邻题严格分开，不得借用/混入彼此内容（朗读、话题表达、信息转述等长文本题尤其注意）
- 分题：以文档题号、序号、节标题、明显换段为准；出现新题号即开始新 JSON 对象
- 有【录题页结构】/【目标题数】时：以之为准；多空题只输出一条，用 blanks 数组，顺序与页面一致
- 无录题页结构时：按文档题号逐条提取；同一材料下多空可合并为一条

【格式规则】
- 内嵌答案：答案以 (A)/(C) 形式写在行首（如 "(C)1.A. B. C."），提取括号内字母为 answer，删除行首 (X) 前缀
- 图片选项：选项为图片时（字母后为空白），options 每项填 "<<IMG>>"
- ★符号：candidates 字段去掉每条最前面的★；其他所有字段保留★，严禁删除
- 标点：英文内容用半角 ASCII 标点（, . ? ! : ;），中文内容保持原格式

【JSON 字段】
- type: 题型，如 "listening_choice" | "listening_response" | "reading_aloud" | "listening_fill" | "listening_fill_and_retell" 等（reading_aloud 对应「交际朗读」）
- listening_script: 仅听力材料正文；无则 ""；有原文就输出，不要省略。对话型（W:/M:）各行用 \n 分隔。多小题时见下方「两层音频」说明
- question: 该题专属文字，去掉开头题号（如 1. 2. (1) 等）。**严禁填大节通用说明**——判断：把这段文字放到同节另一道题，若同样适用 → 通用说明，填 ""；若必须配合本题内容才有意义 → 填入。题干需上传图片时填 ""
- keyword: 仅「参考单词/送评单词」框存在时填对应单词；无则 ""
- options: 选项数组；无则 []
- candidates: 仅 listening_response 使用，候选答案数组（去掉每条最前面的★）
- answer: 单题答案，只填字母（如 "A"）；多空时用 blanks，此处留 ""
- blanks: 多空题时使用 [{question, keyword, answer, options, listening_script}]；question 去掉开头题号；options 无则 []；不要把小题选项放到顶层 options
- image_url: 【录题页结构】标注本题需上传图片时填对应 URL，否则填 ""
- explanation: 每题必须生成，结合题干/听力原文写解析；多小题用「第1小题」「第2小题」等相对编号，不用试卷绝对题号
- 填空类额外说明：placeholder 提示「用'X'分隔多个答案」→ answer 用数组（如 ["red","potatoes"]）；提示「用'|'分隔多种解答」→ 元素内加'|'（如 ["red|Red","potatoes"]）；纯转述型直接用字符串

【多小题两层音频】
情形A（每小题独立对话）：文档有编号分段（"1.W:... 2.W:..."），段数≥小题数，或各小题对话内容明显不同
  → 顶层 listening_script 留 ""，每段放入对应 blank.listening_script（去掉序号）
情形B（所有小题共享同一段材料）：文档为一段连续文本、无分段编号，且该段内容与所有小题均相关
  → 共享段落放顶层 listening_script，blank.listening_script 填该小题提问句（尽量从材料中找到对应提问句填入）
判断原则：优先看文档是否有编号分段——有编号分段且段数≥小题数 → 情形A；只有一段连续文本 → 情形B
严禁把多段独立对话合并成一段塞进顶层 listening_script 或 blank[0]

【听后应答题】
- question 填题目中的问题句（音频会朗读该句）
- 答案材料中的候选项（去掉每条最前面的★）全部放入 candidates
- answer 优先从文档查找；无明确标注则从 candidates 中选最合理的

【交际朗读题】
- type 输出 "reading_aloud"
- 【录题页结构】「含：」有「听力原文」→ listening_script 填原文，question 填题干/要求；无「听力原文」→ listening_script 必须为 ""，朗读内容完整填入 question
- question 填该题专属内容（朗读短文/单词/句子或特有任务要求）；placeholder 含「音标」→ 填「单词 /音标/」格式（如 "deal /diːl/"）；严禁填通用操作说明（"现在你有X秒钟的准备时间"等）；排除通用说明后有朗读内容必须完整保留，确无内容才填 ""
- keyword 只填单词本身，不含音标

【听后记录并转述题】type = "listening_fill_and_retell"
- 输出1条，blanks 按录题页结构的小题数输出（通常为2项）：
  - 第1小题（信息记录）：若答案框提示含"'#'分隔"或"#分隔"→ answer 用数组把所有填空答案合并（如 ["red","potatoes","count"]），question 填""；若无此提示则每空一项
  - 第2小题（信息转述）：question 填转述起始句原文，answer 填参考答案字符串
- 顶层 listening_script 填完整听力原文

只输出 JSON 数组，不要 markdown 代码块，不要其他说明。"""

# 单文件解析：同样通用，目标结构由 paper_metadata 注入
SINGLE_FILE_PROMPT = """你是题目解析助手。请解析试卷内容，按**录题页结构**输出 JSON 数组。

【输出约束】
- 顺序：严格按文档出现顺序，第1条对应页面第1题；不得按题型/相似度换位
- 边界：相邻题严格分开，不得借用/混入彼此内容（朗读、话题表达、信息转述等长文本题尤其注意）
- 分题：以文档题号、序号、节标题、明显换段为准；出现新题号即开始新 JSON 对象
- 有【录题页结构】/【目标题数】时：以之为准；多空题只输出一条，用 blanks 数组，顺序与页面一致
- 无录题页结构时：按文档题号逐条提取；同一材料下多空可合并为一条

【格式规则】
- 内嵌答案：答案以 (A)/(C) 形式写在行首（如 "(C)1.A. B. C."），提取括号内字母为 answer，删除行首 (X) 前缀
- 图片选项：选项为图片时（字母后为空白），options 每项填 "<<IMG>>"
- ★符号：candidates 字段去掉每条最前面的★；其他所有字段保留★，严禁删除
- 标点：英文内容用半角 ASCII 标点（, . ? ! : ;），中文内容保持原格式

【JSON 字段】
- type: 题型，如 "listening_choice" | "listening_response" | "reading_aloud" | "listening_fill" | "listening_fill_and_retell" 等（reading_aloud 对应「交际朗读」）
- listening_script: 仅听力材料正文；无则 ""；有原文就输出，不要省略。对话型（W:/M:）各行用 \n 分隔。多小题时见下方「两层音频」说明
- question: 该题专属文字，去掉开头题号（如 1. 2. (1) 等）。**严禁填大节通用说明**——判断：把这段文字放到同节另一道题，若同样适用 → 通用说明，填 ""；若必须配合本题内容才有意义 → 填入。题干需上传图片时填 ""
- keyword: 仅「参考单词/送评单词」框存在时填对应单词；无则 ""
- options: 选项数组；无则 []
- candidates: 仅 listening_response 使用，候选答案数组（去掉每条最前面的★）
- answer: 单题答案，只填字母（如 "A"）；多空时用 blanks，此处留 ""
- blanks: 多空题时使用 [{question, keyword, answer, options, listening_script}]；question 去掉开头题号；options 无则 []；不要把小题选项放到顶层 options
- image_url: 【录题页结构】标注本题需上传图片时填对应 URL，否则填 ""
- explanation: 每题必须生成，结合题干/听力原文写解析；多小题用「第1小题」「第2小题」等相对编号，不用试卷绝对题号
- 填空类额外说明：placeholder 提示「用'X'分隔多个答案」→ answer 用数组（如 ["red","potatoes"]）；提示「用'|'分隔多种解答」→ 元素内加'|'（如 ["red|Red","potatoes"]）；纯转述型直接用字符串

【多小题两层音频】
情形A（每小题独立对话）：文档有编号分段（"1.W:... 2.W:..."），段数≥小题数，或各小题对话内容明显不同
  → 顶层 listening_script 留 ""，每段放入对应 blank.listening_script（去掉序号）
情形B（所有小题共享同一段材料）：文档为一段连续文本、无分段编号，且该段内容与所有小题均相关
  → 共享段落放顶层 listening_script，blank.listening_script 填该小题提问句（尽量从材料中找到对应提问句填入）
判断原则：优先看文档是否有编号分段——有编号分段且段数≥小题数 → 情形A；只有一段连续文本 → 情形B
严禁把多段独立对话合并成一段塞进顶层 listening_script 或 blank[0]

【听后应答题】
- question 填题目中的问题句（音频会朗读该句）
- 答案材料中的候选项（去掉每条最前面的★）全部放入 candidates
- answer 优先从文档查找；无明确标注则从 candidates 中选最合理的

【交际朗读题】
- type 输出 "reading_aloud"
- 【录题页结构】「含：」有「听力原文」→ listening_script 填原文，question 填题干/要求；无「听力原文」→ listening_script 必须为 ""，朗读内容完整填入 question
- question 填该题专属内容（朗读短文/单词/句子或特有任务要求）；placeholder 含「音标」→ 填「单词 /音标/」格式（如 "deal /diːl/"）；严禁填通用操作说明（"现在你有X秒钟的准备时间"等）；排除通用说明后有朗读内容必须完整保留，确无内容才填 ""
- keyword 只填单词本身，不含音标

【听后记录并转述题】type = "listening_fill_and_retell"
- 输出1条，blanks 按录题页结构的小题数输出（通常为2项）：
  - 第1小题（信息记录）：若答案框提示含"'#'分隔"或"#分隔"→ answer 用数组把所有填空答案合并（如 ["red","potatoes","count"]），question 填""；若无此提示则每空一项
  - 第2小题（信息转述）：question 填转述起始句原文，answer 填参考答案字符串
- 顶层 listening_script 填完整听力原文

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
【顺序约束】请只按原文出现顺序输出，不要依据题型对题目换位；第 1 条对应页面第 1 题，第 2 条对应页面第 2 题。
【边界约束】相邻题目必须分开输出；不要把后一题的内容并入前一题，也不要把前一题的内容延续到后一题。
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
            
            # ── 调试日志：打印每题关键字段，便于排查 placeholder 是否传入 ──
            print(
                f"[slot#{idx}] keys={list(s.keys())} "
                f"answerPhs={s.get('answerPlaceholders')} "
                f"qPhs={s.get('qPlaceholders')} "
                f"currentSlotFields(roles)={[f.get('role') for f in (s.get('currentSlotFields') or [])]}",
                flush=True,
            )

            # 从 sectionLabels 建立标签集合
            section_labels = s.get("sectionLabels") or []
            labels_set = set(section_labels) if isinstance(section_labels, list) else set()

            # 题型只使用页面 HTML 中明确读到的值（typeHint / typeCode），
            # 不做任何推断——平台有200+题型，靠 sectionLabels 猜类型容易出错，
            # 含：/禁：字段约束已经足够告诉 AI 该填什么。
            type_hint = (s.get("typeHint") or "").strip() or (s.get("typeCode") or "").strip()
            
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
                
                # 从本题的 currentSlotFields 建立 role→label 映射（label 通常含 placeholder 文字）
                slot_fields_raw = s.get("currentSlotFields") or []
                role_label = {f.get("role", ""): (f.get("label") or "").strip() for f in slot_fields_raw if f.get("role")}
                
                _SQ_FIELD_ORDER = [
                    ("上传音频", "上传音频"),
                    ("听力原文", "听力原文"),
                    ("题干",     "题干"),
                    ("参考单词", "参考单词"),
                    ("设置选项", "设置选项"),
                    ("设置答案", "设置答案"),
                ]
                
                # 优先用 answerPlaceholders（来自输入框 placeholder 属性）
                _slot_ans_phs = [p for p in (s.get("answerPlaceholders") or []) if p and len(p) > 4]

                def _sq_ph(n: int) -> str:
                    """取第 n 个小题答案框的 placeholder，优先用直接采集的 answerPlaceholders"""
                    # 先用直接采集的 placeholder（最可靠，与题无关只按输入框顺序）
                    if _slot_ans_phs:
                        idx0 = n - 1
                        return _slot_ans_phs[idx0][:60] if idx0 < len(_slot_ans_phs) else _slot_ans_phs[0][:60]
                    # 兜底：从 currentSlotFields label 里找
                    lbl = role_label.get(f"blank_answer_{n}") or role_label.get("answer") or ""
                    return lbl[:60] if lbl and len(lbl) > 4 and lbl.strip().upper() not in {"A","B","C","D"} else ""
                
                # 检查小题是否也有听力原文字段（与顶层同名但含义不同）
                sq_has_script = first_sq_labels and "听力原文" in first_sq_labels
                top_has_script = "听力原文" in labels_set

                if all_same_structure and first_sq_labels:
                    sq_components = [l for l in ["上传音频", "听力原文", "题干", "参考单词", "设置选项", "设置答案"] if l in first_sq_labels]
                    sq_part = f"共{sub}小题，每小题含：{'/'.join(sq_components)}" if sq_components else f"共{sub}小题，结构相同"
                    ph1 = _sq_ph(1)
                    if ph1:
                        sq_part += f"，答案框提示：{ph1}"
                    desc_parts.append(sq_part)
                else:
                    desc_parts.append(f"共{sub}小题，一题多空")
                    if sub_questions:
                        sq_descs = []
                        for sq in sub_questions[:sub]:
                            sq_idx = sq.get("index", len(sq_descs) + 1)
                            sq_labels = set(sq.get("sectionLabels") or [])
                            sq_fields = [disp for key, disp in _SQ_FIELD_ORDER if key in sq_labels]
                            sq_desc = f"第{sq_idx}小题"
                            if sq_fields:
                                sq_desc += f"含：{'/'.join(sq_fields)}"
                            ph = _sq_ph(sq_idx)
                            if ph:
                                sq_desc += f"，答案框提示：{ph}"
                            sq_descs.append(sq_desc)
                        if sq_descs:
                            desc_parts.append("；".join(sq_descs))

                # 两层都有听力原文时，提示 AI 区分情形A/B，不要强制假设共享对话
                if sq_has_script and top_has_script:
                    desc_parts.append("注：大题与小题均有听力原文字段——若所有小题共享同一段对话（情形B）→ 共享段放大题listening_script，blank.listening_script填该小题提问句（尽量从材料中找到对应提问句填入）；若每小题各有独立对话（情形A）→ 大题listening_script留空，每个blank.listening_script填各自对话；严禁把多段独立对话合并塞进大题listening_script")
            
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
                # 多小题：大题层面显示顶层专属字段（音频/原文/图片/属性/解析）
                # 注意：上传音频/听力原文 必须包含，否则「禁：」规则会错误禁止 AI 填写共享对话
                for key, disp in [
                    ("上传音频",     "上传音频"),
                    ("听力原文",     "听力原文"),
                    ("上传图片",     "上传图片"),
                    ("题目属性信息", "题目属性"),
                    ("解析",         "解析"),
                ]:
                    if key in labels_set:
                        main_components.append(disp)
            if main_components:
                # 多小题时用"大题含："，明确区分顶层字段与小题字段
                prefix = "大题含：" if sub > 1 else "含："
                desc_parts.append(f"{prefix}{'/'.join(main_components)}")
            
            # 针对本槽位，将"含："里没有的关键字段明确标注为禁止填写
            # 这样 AI 拿到的是每道题专属的字段禁令，而不只是全局通用规则
            _LABEL_TO_FORBIDDEN = [
                ("听力原文",  'listening_script=""'),
                ("设置选项",  "options=[]"),
                ("设置答案",  'answer=""'),
                ("参考单词",  'keyword=""'),
            ]
            forbidden_hints = [fh for lbl, fh in _LABEL_TO_FORBIDDEN if lbl not in labels_set]
            if forbidden_hints:
                desc_parts.append("禁：" + "，".join(forbidden_hints))
            
            # 单题：优先用 answerPlaceholders / qPlaceholders（来自输入框 placeholder 属性，最直接）
            if sub <= 1:
                ans_phs = [p for p in (s.get("answerPlaceholders") or []) if p and len(p) > 4]
                q_phs   = [p for p in (s.get("qPlaceholders")      or []) if p and len(p) > 4]
                if ans_phs:
                    desc_parts.append(f"答案框提示：{ans_phs[0][:60]}")
                else:
                    # 兜底：从 currentSlotFields label 里找（部分平台会把 placeholder 写进 label）
                    for _f in (s.get("currentSlotFields") or []):
                        _r   = _f.get("role", "")
                        _lbl = (_f.get("label") or "").strip()
                        if _r == "answer" and len(_lbl) > 4 and _lbl.upper() not in {"A","B","C","D"}:
                            desc_parts.append(f"答案框提示：{_lbl[:60]}")
                            break
                if q_phs:
                    desc_parts.append(f"题干框提示：{q_phs[0][:40]}")
                else:
                    for _f in (s.get("currentSlotFields") or []):
                        _r   = _f.get("role", "")
                        _lbl = (_f.get("label") or "").strip()
                        if _r == "question" and len(_lbl) > 4:
                            desc_parts.append(f"题干框提示：{_lbl[:40]}")
                            break
            
            desc = f"序号{idx}"
            if desc_parts:
                desc += "（" + "，".join(desc_parts) + "）"
            
            parts.append(desc)
        
        if parts:
            hint += "【录题页结构】" + "；".join(parts) + "。\n"
            hint += "【字段约束（重要）】每道题只能输出其「含：」或「大题含：」列表中列出的字段内容；列表中**没有**的字段必须输出空值（listening_script=\"\"，options=[]，answer=\"\"，blanks=[]，keyword=\"\"）。例如某题「含：上传音频/题干/题目属性/解析」，则该题 listening_script 必须为 \"\"，不得填入任何内容。\n"
            hint += "【重要】多小题的题只输出一条，用 blanks 数组存放各小题的 {question, keyword, answer, options, listening_script}；「大题含：听力原文」表示共享对话填顶层 listening_script，各 blank.listening_script 填该小题提问句（尽量从材料中找到对应提问句填入）。\n"
    
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
