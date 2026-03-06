# -*- coding: utf-8 -*-
"""
从 Word 文档中读取题目文本，按块切分供 LLM 解析。
"""
from pathlib import Path
from typing import List

from docx import Document


def read_word_text(docx_path: str) -> str:
    """读取 Word 全文（仅文字，不含图片）。
    
    特殊处理：若某段落的 run 有下划线格式且文字全为空白，将其转换为 '___'，
    这样下划线空白行（如转述节的填写区域）不会因 strip() 丢失。
    """
    path = Path(docx_path)
    if not path.exists():
        raise FileNotFoundError(f"Word 文件不存在: {docx_path}")
    doc = Document(path)
    parts = []
    for para in doc.paragraphs:
        # 逐 run 处理：下划线+纯空白 → 替换为下划线符号
        run_texts = []
        for run in para.runs:
            if run.underline and run.text and not run.text.strip():
                # 下划线空白 run（学生填写区域）→ 用 ___ 占位
                run_texts.append("_" * max(len(run.text), 8))
            else:
                run_texts.append(run.text)
        text = "".join(run_texts).strip()
        if text:
            parts.append(text)
    return "\n\n".join(parts)


def chunk_text(
    text: str,
    questions_per_batch: int = 5,
    sep: str = "\n\n",
) -> List[str]:
    """
    将长文本按「题」切分成多块。这里用简单策略：按空行分段，每 N 段为一块。
    若你的 Word 里每题格式固定（如都有「题目」「选项」「答案」），可在这里加强逻辑。
    """
    blocks = [b.strip() for b in text.split(sep) if b.strip()]
    if not blocks:
        return []
    chunks = []
    for i in range(0, len(blocks), questions_per_batch):
        chunk = sep.join(blocks[i : i + questions_per_batch])
        chunks.append(chunk)
    return chunks
