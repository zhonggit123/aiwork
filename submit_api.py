# -*- coding: utf-8 -*-
"""
通过抓包得到的接口，直接 POST 题目 JSON 到服务器。
速度快、不依赖页面结构，强烈推荐。
"""
import json
from typing import Any, Dict, List

import requests


def build_body(
    question: Dict[str, Any],
    body_mapping: Dict[str, str],
) -> Dict[str, Any]:
    """根据配置的字段映射，把 AI 提取的题目转成接口需要的 JSON。"""
    out = {}
    # 默认映射：我们的字段 -> 常见接口字段名
    default_map = {
        "question": "content",
        "options": "options",
        "answer": "answer",
        "explanation": "explanation",
        "type": "type",
    }
    mapping = {**default_map, **(body_mapping or {})}
    for our_key, api_key in mapping.items():
        if our_key in question and api_key:
            out[api_key] = question[our_key]
    return out


def submit_one(
    url: str,
    question: Dict[str, Any],
    headers: Dict[str, str],
    body_mapping: Dict[str, str],
    method: str = "POST",
) -> bool:
    """提交单道题。"""
    body = build_body(question, body_mapping)
    r = requests.request(
        method=method.upper(),
        url=url,
        headers=headers,
        json=body,
        timeout=30,
    )
    r.raise_for_status()
    return True


def submit_all(
    questions: List[Dict[str, Any]],
    url: str,
    headers: Dict[str, str],
    body_mapping: Dict[str, str],
    method: str = "POST",
) -> List[bool]:
    """批量提交，返回每道题是否成功。"""
    results = []
    for q in questions:
        try:
            submit_one(url, q, headers, body_mapping, method)
            results.append(True)
        except Exception as e:
            print(f"提交失败: {q.get('question', '')[:50]}... 错误: {e}")
            results.append(False)
    return results
