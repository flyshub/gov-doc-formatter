"""从旧格式 .doc 提取文本和图片（通过 Word COM 转换 → .docx → 解析）"""
import sys, os, tempfile, subprocess, json

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

def doc_to_items(filepath, image_dir=None):
    import win32com.client
    filepath = os.path.abspath(filepath)
    tmpdir = tempfile.mkdtemp()
    docx_path = os.path.join(tmpdir, '_temp.docx')
    if image_dir is None:
        image_dir = os.path.join(tmpdir, 'images')
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(filepath)
        doc.SaveAs2(docx_path, FileFormat=16)
        doc.Close()
        try:
            word.Quit()
        except Exception:
            pass
        script_dir = os.path.dirname(os.path.abspath(__file__))
        extract_script = os.path.join(script_dir, 'extract_docx.py')
        result = subprocess.run(
            [sys.executable, extract_script, docx_path, image_dir],
            capture_output=True, text=True, encoding='utf-8'
        )
        if result.returncode != 0:
            raise RuntimeError(f"extract_docx failed: {result.stderr}")
        items = json.loads(result.stdout)
        return items, image_dir
    finally:
        try:
            os.remove(docx_path)
        except Exception:
            pass

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python extract_doc.py <input.doc> [image_dir]", file=sys.stderr)
        sys.exit(1)
    filepath = sys.argv[1]
    image_dir = sys.argv[2] if len(sys.argv) > 2 else None
    items, _ = doc_to_items(filepath, image_dir)
    print(json.dumps(items, ensure_ascii=False))