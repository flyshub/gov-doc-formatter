#!/usr/bin/env python3
"""公文正文排版 — CLI 入口与文档组装。

使用方式：
    python -m scripts.format_body --input data.json --output out.docx
    或直接调用 main() 编排完整排版流程。
"""

import argparse, json, sys, re
from pathlib import Path

from docx import Document

from constants import SIZE_3, LEVEL_PATTERNS
from fonts import resolve_font, detect_available_fonts
from hierarchy import detect_hierarchy_issues
from render import setup_page, add_page_number, ELEMENTS


def assemble_document(data, available_fonts):
    """组装公文文档，返回 (Document, warnings)。"""
    warnings = []
    bf = resolve_font("仿宋_GB2312", available_fonts, warnings)

    doc = Document()
    doc.styles["Normal"].font.size = SIZE_3
    setup_page(doc)
    add_page_number(doc)

    for render in ELEMENTS:
        render(doc, data, available_fonts, warnings, bf)

    return doc, warnings


def main():
    parser = argparse.ArgumentParser(description="公文正文排版脚本")
    parser.add_argument("--input", default=None, help="输入 JSON 文件路径")
    parser.add_argument("--stdin", action="store_true", help="从 stdin 读取 JSON")
    parser.add_argument("--output", required=False, help="输出 .docx 文件路径")
    parser.add_argument("--font-dir", default=None, help="额外字体目录")
    parser.add_argument("--check", action="store_true", help="仅校验层级，不生成文档")
    parser.add_argument("--from-extract", action="store_true",
                        help="输入为 extract_docx.py 原始输出（列表），自动提取标题和正文")
    args = parser.parse_args()

    if not args.input and not args.stdin:
        parser.error("必须指定 --input 或 --stdin")
    if not args.check and not args.output:
        parser.error("生成模式下必须指定 --output")

    if args.stdin:
        raw = json.load(sys.stdin)
    else:
        with open(args.input, "r", encoding="utf-8") as f:
            raw = json.load(f)

    # --from-extract: 智能提取标题
    if args.from_extract:
        if isinstance(raw, list) and len(raw) > 0:
            first = raw[0]
            if isinstance(first, str) and not any(
                re.match(p, first.strip()) for p, *_ in LEVEL_PATTERNS
            ):
                raw = {"标题": first, "正文": raw[1:]}
            else:
                basename = Path(args.output).stem if args.output else "未命名公文"
                raw = {"标题": f"{basename} [AI自动生成，请核实]", "正文": raw}
        else:
            print(json.dumps(
                {"status": "error", "message": "--from-extract 要求输入为非空列表"},
                ensure_ascii=False))
            sys.exit(1)

    data = raw

    # ── 输入校验 ──
    errors = []
    if not isinstance(data, dict):
        errors.append("输入 JSON 必须是对象")
    else:
        paragraphs = data.get("正文", [])
        if not isinstance(paragraphs, list) or len(paragraphs) == 0:
            errors.append('"正文" 必须是非空数组')
        else:
            for i, item in enumerate(paragraphs):
                if isinstance(item, str):
                    if not item.strip():
                        errors.append(f"正文第{i+1}段为空字符串")
                elif isinstance(item, dict):
                    if "image" in item:
                        if not isinstance(item["image"], str) or not item["image"].strip():
                            errors.append(f"正文第{i+1}段图片路径无效")
                    elif "table" in item:
                        if not isinstance(item["table"], list):
                            errors.append(f"正文第{i+1}段表格格式无效")
                    elif "text" in item:
                        if not isinstance(item["text"], str) or not item["text"].strip():
                            errors.append(f'正文第{i+1}段 text 为空')
                        # 校验 runs 字段（可选）
                        runs = item.get("runs")
                        if runs is not None:
                            if not isinstance(runs, list):
                                errors.append(f'正文第{i+1}段 runs 必须是数组')
                            else:
                                for ri, r in enumerate(runs):
                                    if not isinstance(r, dict) or "text" not in r:
                                        errors.append(f'正文第{i+1}段 runs[{ri}] 格式无效')
                    else:
                        errors.append(f'正文第{i+1}段缺少 text、image 或 table 字段')
                else:
                    errors.append(f"正文第{i+1}段格式无效")

        for field in ["标题", "称谓", "署名", "日期"]:
            val = data.get(field, "")
            if val and not isinstance(val, str):
                errors.append(f'"{field}" 必须是字符串')

        if "附件" in data:
            att = data["附件"]
            if not isinstance(att, list):
                errors.append('"附件" 必须是数组')
            else:
                for i, item in enumerate(att):
                    if not isinstance(item, str) or not item.strip():
                        errors.append(f"附件第{i+1}项为空或格式无效")

    if errors:
        print(json.dumps(
            {"status": "error", "message": "；".join(errors)}, ensure_ascii=False))
        sys.exit(1)

    # ── 层级合规检查 ──
    hierarchy_issues = detect_hierarchy_issues(data.get("正文", []))

    if args.check:
        result = {"status": "ok", "check_only": True, "hierarchy_issues": hierarchy_issues}
        result["message"] = (
            f"发现 {len(hierarchy_issues)} 处层级跳级问题"
            if hierarchy_issues else "层级顺序符合公文规范"
        )
        print(json.dumps(result, ensure_ascii=False))
        return

    # ── 生成文档 ──
    available_fonts = detect_available_fonts(args.font_dir)
    doc, warnings = assemble_document(data, available_fonts)

    for issue in hierarchy_issues:
        warnings.append(issue["suggestion"])

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))

    result = {
        "status": "ok",
        "path": str(output_path.resolve()),
        "warnings": warnings,
    }
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()