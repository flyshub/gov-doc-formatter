"""批量公文排版 — 支持 .docx / .doc / .txt 混合输入"""
import sys, os, json, subprocess, tempfile, argparse
from pathlib import Path

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = Path(__file__).resolve().parent
EXTRACT_DOCX = SCRIPT_DIR / 'extract_docx.py'
EXTRACT_DOC  = SCRIPT_DIR / 'extract_doc.py'
FORMAT_BODY  = SCRIPT_DIR / 'format_body.py'

def process_one(filepath, output_dir, keep_images=False):
    filepath = Path(filepath).resolve()
    ext = filepath.suffix.lower()
    tmpdir = Path(tempfile.mkdtemp(prefix='govdoc_batch_'))
    img_dir = tmpdir / 'images'
    try:
        if ext == '.docx':
            result = subprocess.run(
                [sys.executable, str(EXTRACT_DOCX), str(filepath), str(img_dir)],
                capture_output=True, text=True, encoding='utf-8')
        elif ext == '.doc':
            result = subprocess.run(
                [sys.executable, str(EXTRACT_DOC), str(filepath), str(img_dir)],
                capture_output=True, text=True, encoding='utf-8')
        elif ext == '.txt':
            text = filepath.read_text(encoding='utf-8')
            items = [l.strip() for l in text.split('\n') if l.strip()]
        else:
            return ('skipped', None, [f'不支持的文件格式: {ext}'])
        if ext in ('.docx', '.doc'):
            if result.returncode != 0:
                return ('error', None, [result.stderr.strip()])
            items = json.loads(result.stdout)
        if ext == '.txt':
            title = filepath.stem + ' [AI自动生成]'
        else:
            title = items[0] if items else filepath.stem
        body = items[1:] if len(items) > 1 else items
        data = {'标题': title, '正文': body}
        json_path = tmpdir / 'input.json'
        json_path.write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')
        output_name = filepath.stem + '.docx'
        output_path = Path(output_dir) / output_name
        output_path.parent.mkdir(parents=True, exist_ok=True)
        result = subprocess.run(
            [sys.executable, str(FORMAT_BODY),
             '--input', str(json_path),
             '--output', str(output_path)],
            capture_output=True, text=True, encoding='utf-8')
        if result.returncode != 0:
            return ('error', None, [result.stderr.strip()])
        resp = json.loads(result.stdout)
        if resp.get('status') == 'ok':
            return ('ok', str(output_path), resp.get('warnings', []))
        else:
            return ('error', None, [resp.get('message', 'Unknown')])
    except Exception as e:
        return ('error', None, [str(e)])
    finally:
        if not keep_images:
            import shutil
            try:
                shutil.rmtree(tmpdir)
            except Exception:
                pass

def main():
    parser = argparse.ArgumentParser(description='批量公文排版')
    parser.add_argument('files', nargs='+', help='输入文件 (.docx/.doc/.txt)')
    parser.add_argument('--output-dir', '-o', default='.', help='输出目录')
    parser.add_argument('--keep-images', action='store_true', help='保留临时图片')
    args = parser.parse_args()
    results = []
    for fp in args.files:
        if not os.path.exists(fp):
            print(f'  {fp}: 文件不存在')
            results.append(('missing', fp, None))
            continue
        status, path, warnings = process_one(fp, args.output_dir, args.keep_images)
        results.append((status, fp, path, warnings))
        if status == 'ok':
            print(f'  {os.path.basename(fp)} -> {path}')
            for w in warnings:
                print(f'    {w}')
        elif status == 'skipped':
            print(f'  {os.path.basename(fp)}: 跳过')
        else:
            print(f'  {os.path.basename(fp)}: {warnings[0] if warnings else "失败"}')
    ok = sum(1 for r in results if r[0] == 'ok')
    err = sum(1 for r in results if r[0] == 'error')
    skip = sum(1 for r in results if r[0] in ('skipped', 'missing'))
    print(f'\n完成: {ok} 成功, {err} 失败, {skip} 跳过 (共 {len(results)} 个)')

if __name__ == '__main__':
    main()