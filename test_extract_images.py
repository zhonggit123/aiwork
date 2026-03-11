# -*- coding: utf-8 -*-
"""
最小化测试脚本：验证从 Word 文档中提取图片及选项匹配是否正确。
用法：
    cd /Users/zhong/Downloads/AI录题
    python test_extract_images.py
"""
import sys
import os

# 确保能 import word_reader
sys.path.insert(0, os.path.dirname(__file__))

from word_reader import extract_word_images

DOCX_PATH = "/Users/zhong/Downloads/贵阳/Unit 3人机对话.docx"

def main():
    print(f"正在提取图片：{DOCX_PATH}")
    try:
        images = extract_word_images(DOCX_PATH)
    except FileNotFoundError as e:
        print(f"[错误] {e}")
        sys.exit(1)

    if not images:
        print("未提取到任何图片。请检查 Word 文档是否包含内嵌图片。")
        sys.exit(0)

    print(f"\n共提取到 {len(images)} 张图片：\n")
    print(f"{'序号':<4} {'选项':<4} {'图片大小(B)':<12} {'格式':<6} {'段落文字(前40字)'}")
    print("-" * 70)
    for img in images:
        size = len(img["image_bytes"])
        para_preview = (img["para_text"] or "(空)").replace("\n", " ")[:40]
        print(f"{img['image_index']:<4} {img['option_label'] or '?':<4} {size:<12} {img['image_ext']:<6} {para_preview}")

    # 将图片保存到临时目录供人工检验
    out_dir = "/tmp/test_word_images"
    os.makedirs(out_dir, exist_ok=True)
    for img in images:
        fname = f"img_{img['image_index']:04d}_opt{img['option_label'] or 'unknown'}.{img['image_ext']}"
        fpath = os.path.join(out_dir, fname)
        with open(fpath, "wb") as f:
            f.write(img["image_bytes"])
    print(f"\n图片已保存到 {out_dir}，可用 Finder 预览验证。")

if __name__ == "__main__":
    main()
