# -*- coding: utf-8 -*-
"""
从 Word 文档中读取题目文本，按块切分供 LLM 解析。
"""
import re
from pathlib import Path
from typing import List

from docx import Document
from docx.oxml.ns import qn


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


_OPTION_LABEL_RE = re.compile(r'^[（(]?([A-Da-d])[)）.\s、。，]')


def _convert_to_web_image(image_bytes: bytes) -> tuple:
    """将任意格式图片（如 TIFF/CMYK/调色板）转换为标准 JPEG，返回 (jpeg_bytes, 'jpg')。

    强制转为 RGB 模式，确保浏览器和各平台均可正常渲染。
    若转换失败则尝试输出 PNG，再失败则原样返回（防崩溃）。
    """
    try:
        import io
        from PIL import Image
        img = Image.open(io.BytesIO(image_bytes))
        # 统一转为 RGB（处理 CMYK / LAB / 调色板 / RGBA 等所有色彩模式）
        if img.mode == 'RGBA':
            # 透明通道合成白底
            bg = Image.new('RGB', img.size, (255, 255, 255))
            bg.paste(img, mask=img.split()[3])
            img = bg
        elif img.mode != 'RGB':
            img = img.convert('RGB')
        out = io.BytesIO()
        img.save(out, format='JPEG', quality=85, optimize=True)
        return out.getvalue(), 'jpg'
    except Exception:
        try:
            import io
            from PIL import Image
            img = Image.open(io.BytesIO(image_bytes))
            img = img.convert('RGB')
            out = io.BytesIO()
            img.save(out, format='PNG')
            return out.getvalue(), 'png'
        except Exception:
            return image_bytes, 'jpg'


def _convert_to_png(image_bytes: bytes) -> tuple:
    """将不被浏览器支持的图片格式（如 TIFF）转换为 PNG。
    返回 (bytes, 'png')；若转换失败则返回原始数据和 'png'（扩展名统一）。
    """
    try:
        import io
        from PIL import Image
        img = Image.open(io.BytesIO(image_bytes))
        out = io.BytesIO()
        img.save(out, format='PNG')
        return out.getvalue(), 'png'
    except Exception:
        return image_bytes, 'png'


def extract_word_images(docx_path: str) -> list:
    """从 Word 文档中提取内嵌图片（inline shapes），返回图片信息列表。

    每项包含：
      para_index   - 所在段落索引（主体段落顺序，表格内段落为 -1）
      para_text    - 该段落文字（可能为空，例如图片独占一段）
      prev_para_text - 前一段落文字（用于推断选项归属）
      option_label - 推断出的选项字母 A/B/C/D，或 None
      image_bytes  - 图片原始字节
      image_ext    - 扩展名（jpg/png/gif/bmp/webp）
      image_index  - 在本文档中的全局顺序（0-based）
    """
    path = Path(docx_path)
    if not path.exists():
        raise FileNotFoundError(f"Word 文件不存在: {docx_path}")

    doc = Document(path)
    results: list = []
    img_idx = 0

    def _extract_from_para(para, para_i: int, prev_text: str) -> None:
        nonlocal img_idx
        drawings = para._element.findall('.//' + qn('w:drawing'))
        if not drawings:
            return
        para_text = para.text.strip()
        _OPT_LETTERS = "ABCD"
        for draw_pos, drawing in enumerate(drawings):
            blip = drawing.find('.//' + qn('a:blip'))
            if blip is None:
                continue
            r_embed = blip.get(qn('r:embed'))
            if not r_embed:
                continue
            try:
                image_part = doc.part.related_parts[r_embed]
                image_bytes = image_part.blob
                content_type = getattr(image_part, 'content_type', 'image/png')
                ext = content_type.split('/')[-1].lower()
                if ext == 'jpeg':
                    ext = 'jpg'
                elif ext not in ('jpg', 'png', 'gif', 'webp'):
                    # TIFF/BMP 等浏览器兼容性差的格式 → 转换为 JPEG
                    image_bytes, ext = _convert_to_web_image(image_bytes)
            except (KeyError, AttributeError):
                continue

            # 推断选项归属：
            # 若同一段落有多张图片（如 A/B/C 都在一行），按段落内顺序 0→A, 1→B, 2→C
            if len(drawings) > 1:
                option_label = _OPT_LETTERS[draw_pos] if draw_pos < len(_OPT_LETTERS) else None
            else:
                # 单图：尝试从本段/前段文字中匹配选项字母
                option_label = None
                for candidate in (para_text, prev_text):
                    m = _OPTION_LABEL_RE.match(candidate)
                    if m:
                        option_label = m.group(1).upper()
                        break

            results.append({
                "para_index": para_i,
                "para_text": para_text,
                "prev_para_text": prev_text,
                "option_label": option_label,
                "image_bytes": image_bytes,
                "image_ext": ext,
                "image_index": img_idx,
            })
            img_idx += 1

    # 遍历主体段落（保留顺序）
    paragraphs = doc.paragraphs
    for para_i, para in enumerate(paragraphs):
        prev_text = paragraphs[para_i - 1].text.strip() if para_i > 0 else ""
        _extract_from_para(para, para_i, prev_text)

    # 遍历表格内段落（表格中的图片）
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell_paras = cell.paragraphs
                for pi, para in enumerate(cell_paras):
                    prev_text = cell_paras[pi - 1].text.strip() if pi > 0 else ""
                    _extract_from_para(para, -1, prev_text)

    return results


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
