#!/usr/bin/env python3
"""字体工具模块 — 检测、解析、设置字体。"""
import os, sys
from pathlib import Path
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from constants import (FONT_CANDIDATES, FONT_FALLBACK, LINE_SPACING_PT,
                         SIZE_3, INDENT_CHARS)

def warn(warnings, msg):
    warnings.append(msg)
    print(f"  ⚠ {msg}", file=sys.stderr)

def resolve_font(desired_name, available_fonts, warnings):
    for name in FONT_CANDIDATES.get(desired_name, [desired_name]):
        if name in available_fonts:
            return name
    fb = FONT_FALLBACK.get(desired_name, "宋体")
    warn(warnings, f"字体 '{desired_name}' 未找到，回退为 '{fb}'")
    return fb

def set_cn_font(run, font_name, size, bold=False):
    run.font.size = size
    run.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for k, v in [("w:eastAsia", font_name),
                 ("w:ascii", "Times New Roman"),
                 ("w:hAnsi", "Times New Roman"),
                 ("w:cs", "Times New Roman")]:
        rFonts.set(qn(k), v)

def set_line_spacing_28(para):
    pPr = para._element.get_or_add_pPr()
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        pPr.append(spacing)
    spacing.set(qn("w:line"), str(int(LINE_SPACING_PT * 20)))
    spacing.set(qn("w:lineRule"), "exact")

def set_first_line_indent(para, chars=INDENT_CHARS):
    para.paragraph_format.first_line_indent = Pt(SIZE_3.pt * chars)

def add_spacer(doc, font_name, available_fonts, warnings, count=1):
    af = resolve_font(font_name, available_fonts, warnings)
    for _ in range(count):
        para = doc.add_paragraph()
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        set_line_spacing_28(para)
        run = para.add_run("")
        set_cn_font(run, af, SIZE_3, False)

def detect_available_fonts(font_dir=None):
    available = set()
    import platform
    system = platform.system()
    font_paths = []
    if system == "Windows":
        windir = Path(os.environ.get("WINDIR", "C:\\Windows"))
        font_paths = [windir / "Fonts",
                      Path.home() / "AppData" / "Local" / "Microsoft" / "Windows" / "Fonts"]
    elif system == "Darwin":
        font_paths = [Path("/Library/Fonts"), Path("/System/Library/Fonts"),
                      Path.home() / "Library" / "Fonts"]
    else:
        font_paths = [Path("/usr/share/fonts"), Path("/usr/local/share/fonts"),
                      Path.home() / ".fonts"]
    if font_dir:
        font_paths.append(Path(font_dir))
    try:
        from matplotlib.font_manager import fontManager
        for f in fontManager.ttflist:
            available.add(f.name)
    except ImportError:
        pass
    import glob
    for fp in font_paths:
        if fp.exists():
            for ext in ("*.ttf", "*.ttc", "*.otf", "*.TTF", "*.TTC", "*.OTF"):
                for f in fp.rglob(ext):
                    available.add(f.stem)
    ALIASES = [
        ("FZXiaoBiaoSong-B05S", "方正小标宋简体"),
        ("FZXBSJW", "方正小标宋简体"),
        ("SimHei", "黑体"),
        ("KaiTi_GB2312", "楷体_GB2312"),
        ("KaiTi", "楷体_GB2312"),
        ("FangSong_GB2312", "仿宋_GB2312"),
        ("FangSong", "仿宋_GB2312"),
    ]
    for fn, alias in ALIASES:
        if fn in available:
            available.add(alias)
    return available
