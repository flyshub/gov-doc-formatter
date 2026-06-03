#!/usr/bin/env python3
"""层级分类与合规检查 — 纯函数模块，可直接单元测试。"""
import re

from constants import LEVEL_PATTERNS, SIZE_3

LEVEL_ORDER = {"h1": 1, "h2": 2, "h3": 3, "h4": 4}
LEVEL_NUM_TO_KEY = {1: "h1", 2: "h2", 3: "h3", 4: "h4"}
LEVEL_NUM_TO_NAME = {1: "一级", 2: "二级", 3: "三级", 4: "四级"}
LEVEL_NAMES = {"h1": ("一级", "一、"), "h2": ("二级", "（一）"),
               "h3": ("三级", "1."),   "h4": ("四级", "(1)")}

STYLE_SALUTATION = {"font_name": "仿宋_GB2312", "size": SIZE_3, "bold": False,
                    "is_heading": False, "level": "salutation"}


def classify_paragraph(text: str) -> dict:
    for pattern, level, font_name, bold in LEVEL_PATTERNS:
        if re.match(pattern, text.strip()):
            return {"font_name": font_name, "size": SIZE_3, "bold": bold,
                    "is_heading": True, "level": level}
    return {"font_name": "仿宋_GB2312", "size": SIZE_3, "bold": False,
            "is_heading": False, "level": "body"}


def detect_hierarchy_issues(body_items: list) -> list:
    issues = []
    last_level_num = 0
    for i, item in enumerate(body_items):
        if isinstance(item, str):
            text = item.strip()
        elif isinstance(item, dict) and "text" in item:
            text = item["text"].strip()
        else:
            continue
        if not text:
            continue
        style = classify_paragraph(text)
        if not style["is_heading"]:
            continue
        curr_num = LEVEL_ORDER[style["level"]]
        if last_level_num > 0 and curr_num > last_level_num + 1:
            skipped = []
            for lv in range(last_level_num + 1, curr_num):
                key = LEVEL_NUM_TO_KEY[lv]
                name, example = LEVEL_NAMES[key]
                skipped.append(f"{name}（{example}）")
            prev_key = LEVEL_NUM_TO_KEY[last_level_num]
            prev_name, _ = LEVEL_NAMES[prev_key]
            curr_name, _ = LEVEL_NAMES[style["level"]]
            suggestion = (
                f"层级跳级：{prev_name}标题后直接出现{curr_name}标题"
                f"（\"{text[:30]}\"），跳过了{"、".join(skipped)}。"
                f"建议在中间插入被跳过的层级标题，"
                f"或将当前段落降级为被跳过的层级。"
            )
            issues.append({
                "index": i,
                "text": text[:60],
                "level": style["level"],
                "prev_level": prev_key,
                "skipped": [LEVEL_NUM_TO_NAME[lv]
                            for lv in range(last_level_num + 1, curr_num)],
                "suggestion": suggestion,
            })
        last_level_num = curr_num
    return issues
