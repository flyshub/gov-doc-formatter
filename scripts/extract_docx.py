"""从 .docx 提取文本和图片，含 Word 自动编号解析"""
import sys, json, re, os, io, shutil, tempfile, zipfile
from pathlib import Path
from docx import Document
from lxml import etree

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
V_NS = 'urn:schemas-microsoft-com:vml'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A_NS  = 'http://schemas.openxmlformats.org/drawingml/2006/main'

CN_NUMS = ['一','二','三','四','五','六','七','八','九','十',
           '十一','十二','十三','十四','十五','十六','十七','十八','十九','二十',
           '二十一','二十二','二十三','二十四','二十五','二十六','二十七','二十八','二十九','三十']

def cn_num(n):
    if 1 <= n <= len(CN_NUMS):
        return CN_NUMS[n-1]
    return str(n)

def parse_numbering(doc):
    numbering_part = doc.part.numbering_part
    if numbering_part is None:
        return {}
    xml = numbering_part._element
    abstract_nums = {}
    for an in xml.findall('.//{%s}abstractNum' % W_NS):
        an_id = an.get('{%s}abstractNumId' % W_NS)
        levels = {}
        for lvl in an.findall('{%s}lvl' % W_NS):
            ilvl = int(lvl.get('{%s}ilvl' % W_NS))
            numFmt = lvl.find('{%s}numFmt' % W_NS)
            lvlText = lvl.find('{%s}lvlText' % W_NS)
            start = lvl.find('{%s}start' % W_NS)
            fmt = numFmt.get('{%s}val' % W_NS) if numFmt is not None else 'decimal'
            text_tmpl = lvlText.get('{%s}val' % W_NS) if lvlText is not None else '%1'
            start_val = int(start.get('{%s}val' % W_NS)) if start is not None else 1
            levels[ilvl] = {'fmt': fmt, 'text': text_tmpl, 'start': start_val}
        abstract_nums[an_id] = levels
    num_map = {}
    for num in xml.findall('.//{%s}num' % W_NS):
        num_id = num.get('{%s}numId' % W_NS)
        an_ref = num.find('{%s}abstractNumId' % W_NS)
        if an_ref is not None:
            an_id = an_ref.get('{%s}val' % W_NS)
            if an_id in abstract_nums:
                num_map[num_id] = abstract_nums[an_id]
    return num_map

def format_number(fmt, lvl_text, num_value):
    if fmt == 'chineseCounting':
        n_str = cn_num(num_value)
    elif fmt == 'decimal':
        n_str = str(num_value)
    elif fmt == 'lowerLetter':
        n_str = chr(ord('a') + num_value - 1)
    elif fmt == 'upperLetter':
        n_str = chr(ord('A') + num_value - 1)
    elif fmt == 'lowerRoman':
        n_str = ['i','ii','iii','iv','v','vi','vii','viii','ix','x'][min(num_value-1,9)]
    elif fmt == 'upperRoman':
        n_str = ['I','II','III','IV','V','VI','VII','VIII','IX','X'][min(num_value-1,9)]
    else:
        n_str = str(num_value)
    result = lvl_text.replace('%1', n_str)
    result = re.sub(r'%\d+', '', result)
    return result

def _extract_images(para_element, rels, zf, img_dir, img_counter):
    images = []
    for shape in para_element.iterfind('.//{%s}shape' % V_NS):
        w_mm, h_mm = None, None
        style = shape.get('style', '')
        wm = re.search(r'width:([\d.]+)pt', style)
        hm = re.search(r'height:([\d.]+)pt', style)
        if wm:
            w_mm = float(wm.group(1)) * 0.3528
        if hm:
            h_mm = float(hm.group(1)) * 0.3528
        if w_mm is None:
            raw_w = shape.get('width') or shape.get('{%s}width' % V_NS)
            if raw_w:
                try:
                    w_mm = float(raw_w) * 0.3528
                except: pass
        for imdata in shape.iterfind('.//{%s}imagedata' % V_NS):
            rid = imdata.get('{%s}id' % R_NS)
            if rid and rid in rels:
                media_path = 'word/' + rels[rid]
                try:
                    data = zf.read(media_path)
                    ext = Path(rels[rid]).suffix or '.png'
                    img_counter[0] += 1
                    img_path = img_dir / f'img_{img_counter[0]:03d}{ext}'
                    img_path.write_bytes(data)
                    png_path, w, h = _convert_to_png(str(img_path), w_mm or 140, h_mm or 100)
                    images.append({'image': png_path, 'width_mm': w, 'height_mm': h})
                except Exception:
                    pass
    for drawing in para_element.iterfind('.//{%s}drawing' % W_NS):
        w_mm, h_mm = None, None
        for ext in drawing.iterfind('.//{%s}extent' % WP_NS):
            cx = int(ext.get('cx', 0))
            cy = int(ext.get('cy', 0))
            if cx:
                w_mm = cx / 360000 * 25.4
            if cy:
                h_mm = cy / 360000 * 25.4
        for blip in drawing.iterfind('.//{%s}blip' % A_NS):
            rid = blip.get('{%s}embed' % R_NS)
            if rid and rid in rels:
                media_path = 'word/' + rels[rid]
                try:
                    data = zf.read(media_path)
                    ext = Path(rels[rid]).suffix or '.png'
                    img_counter[0] += 1
                    img_path = img_dir / f'img_{img_counter[0]:03d}{ext}'
                    img_path.write_bytes(data)
                    png_path, w, h = _convert_to_png(str(img_path), w_mm or 140, h_mm or 100)
                    images.append({'image': png_path, 'width_mm': w, 'height_mm': h})
                except Exception:
                    pass
    return images

def _convert_to_png(img_path, width_mm, height_mm):
    ext = Path(img_path).suffix.lower()
    if ext not in ('.emf', '.wmf'):
        return img_path, width_mm, height_mm
    try:
        from PIL import Image
        png_path = str(Path(img_path).with_suffix('.png'))
        img = Image.open(img_path)
        dpi = 96
        act_w_mm = img.width / dpi * 25.4
        act_h_mm = img.height / dpi * 25.4
        img.save(png_path, 'PNG')
        try:
            os.remove(img_path)
        except Exception:
            pass
        return png_path, act_w_mm or width_mm, act_h_mm or height_mm
    except Exception:
        return img_path, width_mm, height_mm

def extract_with_numbering(filepath, image_dir=None):
    doc = Document(filepath)
    num_map = parse_numbering(doc)
    zf = zipfile.ZipFile(filepath)
    rels = {}
    try:
        rels_xml = zf.read('word/_rels/document.xml.rels')
        rels_tree = etree.fromstring(rels_xml)
        for rel_elem in rels_tree:
            rid = rel_elem.get('Id', '')
            target = rel_elem.get('Target', '')
            if rid and target:
                rels[rid] = target
    except Exception:
        pass
    if image_dir:
        img_dir = Path(image_dir)
    else:
        img_dir = Path(tempfile.mkdtemp(prefix='govdoc_imgs_'))
    img_dir.mkdir(parents=True, exist_ok=True)
    img_counter = [0]
    counters = {}
    items = []
    body = doc.element.body
    for child in body:
        tag = etree.QName(child).localname if child.tag != etree.Comment else ''
        if tag == 'p':
            images = _extract_images(child, rels, zf, img_dir, img_counter)
            pPr = child.find('{%s}pPr' % W_NS)
            para_rPr = pPr.find('{%s}rPr' % W_NS) if pPr is not None else None
            runs = []
            for r_elem in child.findall('.//{%s}r' % W_NS):
                rPr = r_elem.find('{%s}rPr' % W_NS)
                is_bold = False
                if rPr is not None:
                    b = rPr.find('{%s}b' % W_NS)
                    bCs = rPr.find('{%s}bCs' % W_NS)
                    if b is not None:
                        val = b.get('{%s}val' % W_NS)
                        if val is None or (val != '0' and val.lower() != 'false'):
                            is_bold = True
                    if not is_bold and bCs is not None:
                        val = bCs.get('{%s}val' % W_NS)
                        if val is not None and val != '0' and val.lower() != 'false':
                            is_bold = True
                if not is_bold and para_rPr is not None:
                    pb = para_rPr.find('{%s}b' % W_NS)
                    if pb is not None:
                        val = pb.get('{%s}val' % W_NS)
                        if val is None or (val != '0' and val.lower() != 'false'):
                            is_bold = True
                r_texts = r_elem.findall('.//{%s}t' % W_NS)
                r_text = ''.join(t.text or '' for t in r_texts)
                if r_text.strip():
                    runs.append({'text': r_text, 'bold': is_bold})
            plain_text = ''.join(r['text'] for r in runs)
            numPr = pPr.find('{%s}numPr' % W_NS) if pPr is not None else None
            number_prefix = ''
            if numPr is not None:
                ilvl_elem = numPr.find('{%s}ilvl' % W_NS)
                numId_elem = numPr.find('{%s}numId' % W_NS)
                if ilvl_elem is not None and numId_elem is not None:
                    ilvl = int(ilvl_elem.get('{%s}val' % W_NS))
                    numId = numId_elem.get('{%s}val' % W_NS)
                    if numId in num_map and ilvl in num_map[numId]:
                        level_def = num_map[numId][ilvl]
                        if numId not in counters:
                            counters[numId] = {}
                        if ilvl not in counters[numId]:
                            counters[numId][ilvl] = level_def['start']
                        num_value = counters[numId][ilvl]
                        counters[numId][ilvl] = num_value + 1
                        for lvl in list(counters[numId].keys()):
                            if lvl > ilvl:
                                del counters[numId][lvl]
                        number_prefix = format_number(level_def['fmt'], level_def['text'], num_value)
            full_text = (number_prefix + plain_text).strip() if (number_prefix or plain_text) else ''
            has_bold = any(r['bold'] for r in runs)
            if has_bold and full_text:
                output_runs = []
                if number_prefix:
                    output_runs.append({'text': number_prefix, 'bold': False})
                output_runs.extend(runs)
                items.append({'text': full_text, 'runs': output_runs})
            else:
                if images:
                    if full_text:
                        items.append(full_text)
                elif full_text:
                    items.append(full_text)
            if images:
                for img in images:
                    items.append(img)
        elif tag == 'tbl':
            tbl_xml = etree.tostring(child, encoding='unicode')
            items.append({'table_xml': tbl_xml})
    zf.close()
    return items

if __name__ == '__main__':
    filepath = sys.argv[1]
    image_dir = sys.argv[2] if len(sys.argv) > 2 else None
    paragraphs = extract_with_numbering(filepath, image_dir)
    print(json.dumps(paragraphs, ensure_ascii=False))