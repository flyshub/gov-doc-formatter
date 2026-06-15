#!/usr/bin/env python3
"""批量公文排版 — import 式调用，无 subprocess 开销。"""
import sys, os, json, argparse
from pathlib import Path

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

from extract_docx import extract_with_numbering
from format_body import assemble_document
from fonts import detect_available_fonts
from shared import ImageContext, adapt_extract_output


def _doc_to_docx(filepath, tmpdir):
    """将 .doc 通过 Word COM 转换为 .docx，返回 .docx 路径。"""
    import win32com.client
    abs_path = os.path.abspath(filepath)
    docx_path = os.path.join(tmpdir, "_converted.docx")
    word = win32com.client.Dispatch("Word.Application")
    try:
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(abs_path)
        doc.SaveAs2(docx_path, FileFormat=16)
        doc.Close()
    finally:
        try:
            word.Quit()
        except Exception:
            pass
    return docx_path


def process_one(filepath, output_dir, available_fonts):
    filepath = Path(filepath).resolve()
    ext = filepath.suffix.lower()
    if ext not in (".docx", ".doc", ".txt"):
        return ("skipped", None, [f"不支持的文件格式: {ext}"])
    try:
        with ImageContext() as ctx:
            if ext == ".docx":
                items = extract_with_numbering(str(filepath), ctx.dir)
            elif ext == ".doc":
                docx_path = _doc_to_docx(str(filepath), ctx.dir)
                ctx.register_temp(docx_path)
                items = extract_with_numbering(docx_path, ctx.dir)
            else:
                text = filepath.read_text(encoding="utf-8")
                lines = [l.strip() for l in text.split("\n") if l.strip()]
                items = lines
            data = adapt_extract_output(items, fallback_title=filepath.stem)
            doc, warnings = assemble_document(data, available_fonts)
            output_name = filepath.stem + ".docx"
            output_path = Path(output_dir) / output_name
            output_path.parent.mkdir(parents=True, exist_ok=True)
            doc.save(str(output_path))
            return ("ok", str(output_path), warnings)
    except Exception as e:
        return ("error", None, [str(e)])


def main():
    parser = argparse.ArgumentParser(description="批量公文排版")
    parser.add_argument("files", nargs="+", help="输入文件 (.docx/.doc/.txt)")
    parser.add_argument("--output-dir", "-o", default=".", help="输出目录")
    args = parser.parse_args()
    available_fonts = detect_available_fonts()
    results = []
    for fp in args.files:
        if not os.path.exists(fp):
            print(f"  ✗ {fp}: 文件不存在")
            results.append(("missing", fp, None))
            continue
        status, path, warnings = process_one(fp, args.output_dir, available_fonts)
        results.append((status, fp, path, warnings))
        if status == "ok":
            print(f"  ✓ {os.path.basename(fp)} → {path}")
            for w in warnings:
                print(f"    ⚠ {w}")
        elif status == "skipped":
            msg = warnings[0] if warnings else ""
            print(f"  — {os.path.basename(fp)}: 跳过 ({msg})")
        else:
            msg = warnings[0] if warnings else "失败"
            print(f"  ✗ {os.path.basename(fp)}: {msg}")
    ok = sum(1 for r in results if r[0] == "ok")
    err = sum(1 for r in results if r[0] == "error")
    skip = sum(1 for r in results if r[0] in ("skipped", "missing"))
    print(f"\n完成: {ok} 成功, {err} 失败, {skip} 跳过 (共 {len(results)} 个)")


if __name__ == "__main__":
    main()
