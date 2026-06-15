#!/usr/bin/env python3
"""公文排版常量 — GB/T 9704-2012。纯数据模块，无逻辑。"""
from docx.shared import Pt

# ── 页面设置 ──────────────────────────────────────────
PAGE_TOP_MM = 37
PAGE_BOTTOM_MM = 28
PAGE_LEFT_MM = 28
PAGE_RIGHT_MM = 26
LINE_SPACING_PT = 28
INDENT_CHARS = 2
PAGE_CONTENT_WIDTH_MM = 210 - PAGE_LEFT_MM - PAGE_RIGHT_MM  # 版心 156mm

# ── 字号对照 ──────────────────────────────────────────
SIZE_2 = Pt(22)  # 二号
SIZE_3 = Pt(16)  # 三号
SIZE_4 = Pt(14)  # 四号

# ── 层级正则 ──────────────────────────────────────────
LEVEL_PATTERNS = [
    (r"^[一二三四五六七八九十]+、",     "h1", "黑体",         False),
    (r"^[（(][一二三四五六七八九十]+[）)]", "h2", "楷体_GB2312",  False),
    (r"^\d+[.、]",                      "h3", "仿宋_GB2312",  False),
    (r"^[（(]\d+[）)]",                 "h4", "仿宋_GB2312",  False),
]

# 用于 adapt_extract_output / --from-extract 的标题识别正则（仅 pattern 部分）
HEAD_RE_PATTERNS = [
    r"^[一二三四五六七八九十]+、",
    r"^[（(][一二三四五六七八九十]+[）)]",
    r"^\d+[.、]",
    r"^[（(]\d+[）)]",
]

# ── 字体映射 ──────────────────────────────────────────
FONT_CANDIDATES = {
    "方正小标宋简体": ["方正小标宋简体", "FZXiaoBiaoSong-B05S", "FZXiaoBiaoSong"],
    "黑体":           ["黑体", "SimHei"],
    "楷体_GB2312":    ["楷体_GB2312", "KaiTi_GB2312", "KaiTi", "楷体"],
    "仿宋_GB2312":    ["仿宋_GB2312", "FangSong_GB2312", "FangSong", "仿宋"],
}

FONT_FALLBACK = {
    "方正小标宋简体": "宋体",
    "黑体":           "SimHei",
    "楷体_GB2312":    "楷体",
    "仿宋_GB2312":    "仿宋",
}
