#!/usr/bin/env python3
"""公文正文排版脚本 — 按 GB/T 9704-2012 格式生成 .docx"""

import argparse
import json
import os
import re
import sys
from pathlib import Path

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

from docx import Document
from docx.shared import Pt, Cm, Mm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PAGE_TOP_MM    = 37
PAGE_BOTTOM_MM = 28
PAGE_LEFT_MM   = 28
PAGE_RIGHT_MM  = 26
LINE_SPACING_PT = 28
SIZE_2 = Pt(22)
SIZE_3 = Pt(16)
SIZE_4 = Pt(14)
INDENT_CHARS = 2

LEVEL_PATTERNS = [
    (r'^[一二三四五六七八九十]+、',     'h1', '黑体',         False),
    (r'^[（(][一二三四五六七八九十]+[）)]', 'h2', '楷体_GB2312',  False),
    (r'^\d+[.、]',                      'h3', '仿宋_GB2312',  False),
    (r'^[（(]\d+[）)]',                 'h4', '仿宋_GB2312',  False),
]

_h1 = LEVEL_PATTERNS[0][0].lstrip('^')
_h2 = LEVEL_PATTERNS[1][0].lstrip('^')
HEADING_SPLIT_RE = re.compile(r'^(' + _h1 + '|' + _h2 + ')')

FONT_CANDIDATES = {
    '方正小标宋简体': ['方正小标宋简体', 'FZXiaoBiaoSong-B05S', 'FZXiaoBiaoSong'],
    '黑体':           ['黑体', 'SimHei'],
    '楷体_GB2312':    ['楷体_GB2312', 'KaiTi_GB2312', 'KaiTi', '楷体'],
    '仿宋_GB2312':    ['仿宋_GB2312', 'FangSong_GB2312', 'FangSong', '仿宋'],
}

FONT_FALLBACK = {
    '方正小标宋简体': '宋体',
    '黑体':           'SimHei',
    '楷体_GB2312':    '楷体',
    '仿宋_GB2312':    '仿宋',
}

def warn(warnings, msg):
    warnings.append(msg)
    print(f"  {msg}", file=sys.stderr)

def resolve_font(desired_name, available_fonts, warnings):
    candidates = FONT_CANDIDATES.get(desired_name, [desired_name])
    for name in candidates:
        if name in available_fonts:
            return name
    fallback = FONT_FALLBACK.get(desired_name, '宋体')
    warn(warnings, f"字体 '{desired_name}' 未找到，回退为 '{fallback}'")
    return fallback

def set_cn_font(run, font_name, size, bold=False):
    run.font.size = size
    run.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'), 'Times New Roman')

def set_line_spacing_28(paragraph):
    pPr = paragraph._element.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    spacing.set(qn('w:line'), str(int(LINE_SPACING_PT * 20)))
    spacing.set(qn('w:lineRule'), 'exact')

def set_first_line_indent(paragraph, chars=INDENT_CHARS):
    paragraph.paragraph_format.first_line_indent = Pt(SIZE_3.pt * chars)

def add_spacer(doc, font_name, available_fonts, warnings, count=1):
    actual_font = resolve_font(font_name, available_fonts, warnings)
    for _ in range(count):
        para = doc.add_paragraph()
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after  = Pt(0)
        set_line_spacing_28(para)
        run = para.add_run('')
        set_cn_font(run, actual_font, SIZE_3, False)

def setup_page(doc):
    section = doc.sections[0]
    section.page_width  = Mm(210)
    section.page_height = Mm(297)
    section.top_margin    = Mm(PAGE_TOP_MM)
    section.bottom_margin = Mm(PAGE_BOTTOM_MM)
    section.left_margin   = Mm(PAGE_LEFT_MM)
    section.right_margin  = Mm(PAGE_RIGHT_MM)

def add_page_number(doc):
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False
    para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(0)
    for prefix in ['— ', '', ' —']:
        run = para.add_run(prefix)
        run.font.size = SIZE_4
        rPr = run._element.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:eastAsia'), '宋体')
        rPr.insert(0, rFonts)
        if prefix == '':
            fldChar_begin = OxmlElement('w:fldChar')
            fldChar_begin.set(qn('w:fldCharType'), 'begin')
            run._element.append(fldChar_begin)
            instrText = OxmlElement('w:instrText')
            instrText.set(qn('xml:space'), 'preserve')
            instrText.text = ' PAGE '
            run._element.append(instrText)
            fldChar_end = OxmlElement('w:fldChar')
            fldChar_end.set(qn('w:fldCharType'), 'end')
            run._element.append(fldChar_end)

def add_title(doc, text, available_fonts, warnings):
    actual_font = resolve_font('方正小标宋简体', available_fonts, warnings)
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(0)
    set_line_spacing_28(para)
    para.paragraph_format.first_line_indent = Pt(0)
    run = para.add_run(text)
    set_cn_font(run, actual_font, SIZE_2, bold=False)

def classify_paragraph(text):
    for pattern, level, font_name, bold in LEVEL_PATTERNS:
        if re.match(pattern, text.strip()):
            return {'font_name': font_name, 'size': SIZE_3, 'bold': bold,
                    'is_heading': True, 'level': level}
    return {'font_name': '仿宋_GB2312', 'size': SIZE_3, 'bold': False,
            'is_heading': False, 'level': 'body'}

STYLE_SALUTATION = {'font_name': '仿宋_GB2312', 'size': SIZE_3, 'bold': False,
                    'is_heading': False, 'level': 'salutation'}

def add_body_paragraph(doc, text, style, available_fonts, warnings,
                       alignment=None, no_indent=False):
    actual_font = resolve_font(style['font_name'], available_fonts, warnings)
    body_font = resolve_font('仿宋_GB2312', available_fonts, warnings)
    is_heading = style.get('is_heading', False)
    bold       = style.get('bold', False)
    size       = style.get('size', SIZE_3)
    para = doc.add_paragraph()
    para.alignment = alignment if alignment is not None else WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(0)
    set_line_spacing_28(para)
    if no_indent:
        para.paragraph_format.first_line_indent = Pt(0)
    else:
        set_first_line_indent(para)
    stripped = text.strip()
    if re.match(r'^\d+\.', stripped) and not re.match(r'^\d+\.\s', stripped):
        m = re.match(r'^(\d+\.)(.*)', stripped)
        if m:
            stripped = m.group(1) + ' ' + m.group(2)
    heading_split = None
    if is_heading:
        m = HEADING_SPLIT_RE.match(stripped)
        if m:
            after_marker = stripped[m.end():]
            split_match = re.search(r'[：。]', after_marker)
            if split_match and len(after_marker[:split_match.start()].strip()) >= 2:
                cut = m.end() + split_match.end()
                heading_split = cut
    if heading_split:
        prefix = stripped[:heading_split]
        suffix = stripped[heading_split:].lstrip()
        run1 = para.add_run(prefix)
        set_cn_font(run1, actual_font, size, bold)
        if suffix:
            run2 = para.add_run(suffix)
            set_cn_font(run2, body_font, size, False)
    else:
        run = para.add_run(stripped)
        set_cn_font(run, actual_font, size, bold)

def detect_available_fonts(font_dir=None):
    available = set()
    import platform
    system = platform.system()
    font_paths = []
    if system == 'Windows':
        windir = Path(os.environ.get('WINDIR', 'C:\\Windows'))
        font_paths = [
            windir / 'Fonts',
            Path.home() / 'AppData' / 'Local' / 'Microsoft' / 'Windows' / 'Fonts',
        ]
    elif system == 'Darwin':
        font_paths = [
            Path('/Library/Fonts'),
            Path('/System/Library/Fonts'),
            Path.home() / 'Library' / 'Fonts',
        ]
    else:
        font_paths = [
            Path('/usr/share/fonts'),
            Path('/usr/local/share/fonts'),
            Path.home() / '.fonts',
        ]
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
            for ext in ('*.ttf', '*.ttc', '*.otf', '*.TTF', '*.TTC', '*.OTF'):
                for f in fp.rglob(ext):
                    available.add(f.stem)
    ALIASES = [
        ('FZXiaoBiaoSong-B05S', '方正小标宋简体'),
        ('FZXBSJW',             '方正小标宋简体'),
        ('SimHei',              '黑体'),
        ('KaiTi_GB2312',        '楷体_GB2312'),
        ('KaiTi',               '楷体_GB2312'),
        ('FangSong_GB2312',     '仿宋_GB2312'),
        ('FangSong',            '仿宋_GB2312'),
    ]
    for from_name, to_alias in ALIASES:
        if from_name in available:
            available.add(to_alias)
    return available

def _render_title(doc, data, available_fonts, warnings, body_font):
    title = data.get('标题', '').strip()
    if title:
        add_title(doc, title, available_fonts, warnings)

def _render_salutation(doc, data, available_fonts, warnings, body_font):
    称谓 = data.get('称谓', '').strip()
    if 称谓:
        add_body_paragraph(doc, 称谓, STYLE_SALUTATION,
                          available_fonts, warnings, no_indent=True)

def _render_body(doc, data, available_fonts, warnings, body_font):
    PAGE_WIDTH_MM = 210 - PAGE_LEFT_MM - PAGE_RIGHT_MM
    for item in data.get('正文', []):
        if isinstance(item, dict) and 'image' in item:
            img_path = item['image']
            if not os.path.exists(img_path):
                warn(warnings, f"图片不存在: {img_path}")
                continue
            w_mm = min(item.get('width_mm', PAGE_WIDTH_MM), PAGE_WIDTH_MM)
            h_mm = item.get('height_mm', w_mm * 0.75)
            if item.get('width_mm') and item['width_mm'] > PAGE_WIDTH_MM:
                ratio = PAGE_WIDTH_MM / item['width_mm']
                h_mm = item['height_mm'] * ratio
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_before = Pt(6)
            para.paragraph_format.space_after  = Pt(6)
            para.paragraph_format.first_line_indent = Pt(0)
            run = para.add_run()
            try:
                run.add_picture(img_path, width=Mm(w_mm), height=Mm(h_mm))
            except Exception as e:
                warn(warnings, f"图片插入失败 ({os.path.basename(img_path)}): {e}")
                p = para._element
                p.getparent().remove(p)
            continue
        if isinstance(item, dict) and 'table' in item:
            rows = item['table']
            if not rows:
                continue
            tbl = doc.add_table(rows=len(rows), cols=len(rows[0]))
            tbl.alignment = WD_ALIGN_PARAGRAPH.CENTER
            tbl.autofit = True
            tbl.style = 'Table Grid'
            header_font_name = resolve_font('黑体', available_fonts, warnings)
            cell_font_name   = resolve_font('仿宋_GB2312', available_fonts, warnings)
            for ri, row_data in enumerate(rows):
                for ci, cell_text in enumerate(row_data):
                    if ci >= len(tbl.rows[ri].cells):
                        break
                    cell = tbl.rows[ri].cells[ci]
                    cell.paragraphs[0].clear()
                    run = cell.paragraphs[0].add_run(cell_text)
                    if ri == 0:
                        set_cn_font(run, header_font_name, Pt(10.5), bold=False)
                        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        set_cn_font(run, cell_font_name, Pt(10.5), bold=False)
            add_spacer(doc, '仿宋_GB2312', available_fonts, warnings, count=1)
            continue
        if isinstance(item, str):
            text = item.strip()
            para_align = None
            no_indent  = False
        else:
            text = item.get('text', '').strip()
            align_map = {'left': WD_ALIGN_PARAGRAPH.LEFT,
                         'right': WD_ALIGN_PARAGRAPH.RIGHT,
                         'center': WD_ALIGN_PARAGRAPH.CENTER}
            para_align = align_map.get(item.get('align'), None)
            no_indent  = item.get('indent') is False
        if not text:
            continue
        style = classify_paragraph(text)
        add_body_paragraph(doc, text, style,
                          available_fonts, warnings,
                          alignment=para_align, no_indent=no_indent)

def _render_attachments(doc, data, available_fonts, warnings, body_font):
    attachments = data.get('附件', [])
    if not attachments:
        return
    add_spacer(doc, '仿宋_GB2312', available_fonts, warnings, count=1)
    for i, item in enumerate(attachments, 1):
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after  = Pt(0)
        set_line_spacing_28(para)
        set_first_line_indent(para)
        text = f"附件：{i}. {item}" if i == 1 else f"      {i}. {item}"
        run = para.add_run(text)
        set_cn_font(run, body_font, SIZE_3, False)

def _render_signature(doc, data, available_fonts, warnings, body_font):
    署名 = data.get('署名', '').strip()
    日期 = data.get('日期', '').strip()
    if not (署名 or 日期):
        return
    add_spacer(doc, '仿宋_GB2312', available_fonts, warnings, count=2)
    if 署名:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after  = Pt(0)
        set_line_spacing_28(para)
        para.paragraph_format.first_line_indent = Pt(0)
        run = para.add_run(署名)
        set_cn_font(run, body_font, SIZE_3, False)
    if 日期:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after  = Pt(0)
        para.paragraph_format.right_indent = Pt(SIZE_3.pt * 2)
        set_line_spacing_28(para)
        para.paragraph_format.first_line_indent = Pt(0)
        run = para.add_run(日期)
        set_cn_font(run, body_font, SIZE_3, False)

ELEMENTS = [
    _render_title,
    _render_salutation,
    _render_body,
    _render_attachments,
    _render_signature,
]

def assemble_document(data, available_fonts):
    warnings = []
    body_font = resolve_font('仿宋_GB2312', available_fonts, warnings)
    doc = Document()
    default_style = doc.styles['Normal']
    default_style.font.size = SIZE_3
    setup_page(doc)
    add_page_number(doc)
    for render in ELEMENTS:
        render(doc, data, available_fonts, warnings, body_font)
    return doc, warnings

def main():
    parser = argparse.ArgumentParser(description='公文正文排版脚本')
    parser.add_argument('--input', required=True, help='输入 JSON 文件路径')
    parser.add_argument('--output', required=True, help='输出 .docx 文件路径')
    parser.add_argument('--font-dir', default=None, help='额外字体目录')
    args = parser.parse_args()
    with open(args.input, 'r', encoding='utf-8') as f:
        data = json.load(f)
    errors = []
    if not isinstance(data, dict):
        errors.append('输入 JSON 必须是对象')
    else:
        paragraphs = data.get('正文', [])
        if not isinstance(paragraphs, list) or len(paragraphs) == 0:
            errors.append('"正文" 必须是非空数组')
        else:
            for i, item in enumerate(paragraphs):
                if isinstance(item, str):
                    if not item.strip():
                        errors.append(f'正文第{i+1}段为空字符串')
                elif isinstance(item, dict):
                    if 'image' in item:
                        if not isinstance(item['image'], str) or not item['image'].strip():
                            errors.append(f'正文第{i+1}段图片路径无效')
                    elif 'table' in item:
                        if not isinstance(item['table'], list):
                            errors.append(f'正文第{i+1}段表格格式无效')
                    elif 'text' not in item:
                        errors.append(f'正文第{i+1}段缺少 "text" 或 "image" 字段')
                    elif not isinstance(item['text'], str) or not item['text'].strip():
                        errors.append(f'正文第{i+1}段 "text" 为空')
                else:
                    errors.append(f'正文第{i+1}段格式无效：必须是字符串或对象')
        for field in ['标题', '称谓', '署名', '日期']:
            val = data.get(field, '')
            if val and not isinstance(val, str):
                errors.append(f'"{field}" 必须是字符串')
        if '附件' in data:
            att = data['附件']
            if not isinstance(att, list):
                errors.append('"附件" 必须是数组')
            else:
                for i, item in enumerate(att):
                    if not isinstance(item, str) or not item.strip():
                        errors.append(f'附件第{i+1}项为空或格式无效')
    if errors:
        print(json.dumps({'status': 'error', 'message': '；'.join(errors)}, ensure_ascii=False))
        sys.exit(1)
    available_fonts = detect_available_fonts(args.font_dir)
    doc, warnings = assemble_document(data, available_fonts)
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    result = {'status': 'ok', 'path': str(output_path.resolve()), 'warnings': warnings}
    print(json.dumps(result, ensure_ascii=False))

if __name__ == '__main__':
    main()