# -*- coding: utf-8 -*-
"""
Word 题库自动录入：读取 Word -> AI 解析为 JSON -> 提交到录题页面/接口。
用法：
  python main.py                    # 使用 config.yaml
  python main.py --word 其他.docx   # 指定 Word 路径
  python main.py --dry-run          # 只解析不提交，打印 JSON
"""
import argparse
from pathlib import Path

import yaml

from word_reader import read_word_text, chunk_text
from llm_extract import extract_questions_from_word_chunks
from submit_api import submit_all as submit_all_api
from submit_playwright import submit_all_playwright


def load_config(config_path: str = "config.yaml") -> dict:
    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(
            f"请复制 config.example.yaml 为 {config_path} 并填写配置"
        )
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def run(
    word_path: str,
    config: dict,
    dry_run: bool = False,
) -> None:
    # 1. 读 Word、分块
    text = read_word_text(word_path)
    chunks = chunk_text(
        text,
        questions_per_batch=config.get("word", {}).get("questions_per_batch", 5),
    )
    if not chunks:
        print("Word 中未解析到有效文本块，请检查文档格式。")
        return

    # 2. LLM 解析
    llm = config.get("llm", {})
    questions = extract_questions_from_word_chunks(
        chunks,
        base_url=llm.get("base_url", "https://api.openai.com/v1"),
        api_key=llm.get("api_key", ""),
        model=llm.get("model", "gpt-4o-mini"),
    )
    print(f"共解析出 {len(questions)} 道题。")

    if dry_run:
        import json
        print(json.dumps(questions, ensure_ascii=False, indent=2))
        return

    # 3. 提交
    submit_cfg = config.get("submit", {})
    mode = submit_cfg.get("mode", "api")

    if mode == "api":
        api_cfg = submit_cfg.get("api", {})
        url = api_cfg.get("url")
        if not url:
            print("config.submit.api.url 未配置，请填写抓包得到的接口地址。")
            return
        headers = api_cfg.get("headers", {})
        body_mapping = api_cfg.get("body_mapping", {})
        method = api_cfg.get("method", "POST")
        results = submit_all_api(
            questions,
            url=url,
            headers=headers,
            body_mapping=body_mapping,
            method=method,
        )
    else:
        pw_cfg = submit_cfg.get("playwright", {})
        page_url = pw_cfg.get("page_url")
        selectors = pw_cfg.get("selectors", {})
        if not page_url or not selectors:
            print("config.submit.playwright 需配置 page_url 和 selectors。")
            return
        results = submit_all_playwright(
            questions,
            page_url=page_url,
            selectors=selectors,
        )

    ok = sum(results)
    print(f"提交完成：成功 {ok}/{len(results)} 题。")


def main():
    parser = argparse.ArgumentParser(description="Word 题库自动录入")
    parser.add_argument("--config", "-c", default="config.yaml", help="配置文件路径")
    parser.add_argument("--word", "-w", default=None, help="Word 文件路径（默认用 config 里的 word.path）")
    parser.add_argument("--dry-run", action="store_true", help="只解析不提交，输出 JSON")
    args = parser.parse_args()

    config = load_config(args.config)
    word_path = args.word or config.get("word", {}).get("path", "题库.docx")
    run(word_path=word_path, config=config, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
