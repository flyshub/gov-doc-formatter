#!/usr/bin/env python3
"""shared type definitions — types contract between extract and format modules."""

from pathlib import Path
from typing import TypedDict, Union, List, Dict


class ExtractedImage(TypedDict, total=False):
    image: str
    width_mm: float
    height_mm: float


class ExtractedTable(TypedDict, total=False):
    table: List[List[str]]
    table_xml: str  # 完整 <w:tbl> XML 字符串，保留全部格式


ExtractedItem = Union[str, ExtractedImage, ExtractedTable]
ExtractOutput = List[ExtractedItem]


class FormatInput(TypedDict, total=False):
    标题: str
    称谓: str
    正文: List[Union[str, dict]]
    附件: List[str]
    署名: str
    日期: str


def adapt_extract_output(raw: list, fallback_title: str = "未命名公文") -> FormatInput:
    """Adapt extract_docx.py output to format_body.py input."""
    import re
    HEADING_PATTERNS = [
        r"^[一二三四五六七八九十]+、",
        r"^[（(][一二三四五六七八九十]+[）)]",
        r"^\d+[.、]",
        r"^[（(]\d+[）)]",
    ]
    if not isinstance(raw, list) or len(raw) == 0:
        return {"标题": fallback_title, "正文": []}
    first = raw[0]
    if isinstance(first, str) and not any(
        re.match(p, first.strip()) for p in HEADING_PATTERNS
    ):
        return {"标题": first, "正文": raw[1:]}
    return {"标题": f"{fallback_title} [AI自动生成，请核实]", "正文": raw}


# ── 图片生命周期管理 ──────────────────────────────────

class ImageContext:
    """上下文管理器：自动管理图片临时目录的生命周期。"""
    def __init__(self, directory=None, cleanup=True):
        import tempfile
        self.cleanup = cleanup
        self.dir = directory if directory else tempfile.mkdtemp(prefix="govdoc_imgs_")

    def __enter__(self):
        Path(self.dir).mkdir(parents=True, exist_ok=True)
        return self

    def __exit__(self, *args):
        if self.cleanup:
            import shutil
            try:
                shutil.rmtree(self.dir)
            except Exception:
                pass