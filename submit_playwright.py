# -*- coding: utf-8 -*-
"""
使用 Playwright 模拟浏览器操作，在录题页面逐个输入框填写并提交。
适合没有接口文档、只能通过网页操作的场景。需在 config 里配置各输入框的选择器。
"""
from typing import Any, Dict, List

from playwright.sync_api import sync_playwright


def _fill_one(
    page,
    question: Dict[str, Any],
    selectors: Dict[str, str],
) -> None:
    """在当前页面上填一道题并点提交。"""
    sel = selectors
    # 题干
    if sel.get("question"):
        page.locator(sel["question"]).fill(question.get("question", "") or "")
    # 选项 A/B/C/D（若页面是分开的输入框）
    for key in ["option_a", "option_b", "option_c", "option_d"]:
        selector = sel.get(key)
        if not selector:
            continue
        opts = question.get("options") or []
        idx = ord(key.split("_")[1].upper()) - ord("A")
        value = opts[idx] if idx < len(opts) else ""
        page.locator(selector).fill(value)
    # 答案、解析
    if sel.get("answer"):
        page.locator(sel["answer"]).fill(question.get("answer", "") or "")
    if sel.get("explanation"):
        page.locator(sel["explanation"]).fill(question.get("explanation", "") or "")
    # 提交
    if sel.get("submit_btn"):
        page.locator(sel["submit_btn"]).click()


def submit_all_playwright(
    questions: List[Dict[str, Any]],
    page_url: str,
    selectors: Dict[str, str],
    headless: bool = False,
) -> List[bool]:
    """
    打开录题页面，循环：填一题 -> 提交 -> 等待（如有“新增下一题”再点）。
    返回每道题是否操作成功。若页面有“保存并继续”需在 selectors 里加对应按钮。
    """
    results = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context()
        page = context.new_page()
        page.goto(page_url)
        for i, q in enumerate(questions):
            try:
                _fill_one(page, q, selectors)
                results.append(True)
                # 若提交后需点“下一题”才能继续填，可在这里加等待和点击
                # page.wait_for_timeout(500)
                # page.locator(selectors.get("next_btn", "")).click()
            except Exception as e:
                print(f"第 {i+1} 题填写失败: {e}")
                results.append(False)
        browser.close()
    return results
