"""Walk a .docx body in document order and report how it is actually built.

Tables, cells, nested tables, text boxes, images with their real display size,
fills, font sizes, colours, alignment and spacing - in the order Word lays them
out. This is how you reverse-engineer a document whose look you want to
reproduce, and it is deliberately a digest: `word/document.xml` in a
design-heavy file runs well over 100 KB, which is no way to read a layout.

    python -m docxforge.inspect_docx source.docx
    python -m docxforge.inspect_docx source.docx --out-dir ./derived

Writes <name>.outline.md (the text and image sequence) and <name>.structure.txt
(the annotated tree), and prints the structure to stdout.
"""
import argparse
import collections
import os
import re
import zipfile
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
WP = '{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}'
PIC = '{http://schemas.openxmlformats.org/drawingml/2006/picture}'


def tag(e):
    return e.tag.split('}')[-1]


def para_text(p):
    return ''.join(t.text or '' for t in p.iter(W + 't'))


def para_props(p):
    pPr = p.find(W + 'pPr')
    out = {}
    if pPr is not None:
        shd = pPr.find(W + 'shd')
        if shd is not None:
            out['fill'] = shd.get(W + 'fill')
        jc = pPr.find(W + 'jc')
        if jc is not None:
            out['align'] = jc.get(W + 'val')
        sp = pPr.find(W + 'spacing')
        if sp is not None:
            out['spacing'] = {k.split('}')[-1]: v for k, v in sp.attrib.items()}
        ind = pPr.find(W + 'ind')
        if ind is not None:
            out['ind'] = {k.split('}')[-1]: v for k, v in ind.attrib.items()}
    # first run font info
    for r in p.iter(W + 'r'):
        rPr = r.find(W + 'rPr')
        if rPr is None:
            continue
        sz = rPr.find(W + 'sz')
        col = rPr.find(W + 'color')
        rf = rPr.find(W + 'rFonts')
        out['run'] = {
            'sz_half': sz.get(W + 'val') if sz is not None else None,
            'color': col.get(W + 'val') if col is not None else None,
            'font': rf.get(W + 'ascii') if rf is not None else None,
            'b': rPr.find(W + 'b') is not None,
            'i': rPr.find(W + 'i') is not None,
        }
        break
    return out


def images_in(el, rels):
    out = []
    for blip in el.iter(A + 'blip'):
        rid = blip.get(R + 'embed')
        out.append({'rid': rid, 'target': rels.get(rid)})
    for ext in el.iter('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent'):
        if out:
            out[-1]['cx'] = ext.get('cx')
            out[-1]['cy'] = ext.get('cy')
    return out


def textboxes_in(el):
    """Return list of text-box contents (each a list of paragraph texts)."""
    boxes = []
    for txbx in el.iter(W + 'txbxContent'):
        paras = []
        for p in txbx.findall(W + 'p'):
            t = para_text(p).strip()
            props = para_props(p)
            paras.append((t, props))
        boxes.append(paras)
    # alternate-content fallback (mc:Fallback) duplicates boxes; dedupe
    seen, uniq = set(), []
    for b in boxes:
        key = tuple(t for t, _ in b)
        if key in seen:
            continue
        seen.add(key)
        uniq.append(b)
    return uniq


def shape_fill(el):
    fills = []
    for solid in el.iter(A + 'solidFill'):
        clr = solid.find(A + 'srgbClr')
        if clr is not None:
            fills.append(clr.get('val'))
    return fills


def main(path, out_dir='.', echo=True):
    z = zipfile.ZipFile(path)
    raw = z.read('word/document.xml').decode('utf-8', 'ignore')
    relxml = ET.fromstring(z.read('word/_rels/document.xml.rels'))
    rels = {r.get('Id'): r.get('Target') for r in relxml}
    root = ET.fromstring(z.read('word/document.xml'))
    body = root.find(W + 'body')

    lines = []
    report = []

    def emit_para(p, depth, ctx):
        t = para_text(p).strip()
        pr = para_props(p)
        imgs = images_in(p, rels)
        boxes = textboxes_in(p)
        fills = shape_fill(p)
        pad = '  ' * depth
        if t:
            run = pr.get('run') or {}
            sz = run.get('sz_half')
            sz = f"{int(sz)/2:g}pt" if sz else '-'
            report.append(f"{pad}P {ctx} [{sz} {run.get('font') or '-'} "
                          f"{'B' if run.get('b') else ''}{'I' if run.get('i') else ''} "
                          f"c={run.get('color') or '-'} fill={pr.get('fill') or '-'} "
                          f"al={pr.get('align') or '-'}] {t[:90]}")
            lines.append(t)
        for im in imgs:
            cx = int(im.get('cx') or 0)
            cy = int(im.get('cy') or 0)
            report.append(f"{pad}IMG {ctx} {im['target']} "
                          f"{cx/914400:.2f}in x {cy/914400:.2f}in")
            lines.append(f"![{im['target']}]")
        for bi, box in enumerate(boxes):
            report.append(f"{pad}TXBOX {ctx}.{bi} fills={fills}")
            for bt, bpr in box:
                if not bt:
                    continue
                run = bpr.get('run') or {}
                sz = run.get('sz_half')
                sz = f"{int(sz)/2:g}pt" if sz else '-'
                report.append(f"{pad}    tb[{sz} {run.get('font') or '-'} "
                              f"{'B' if run.get('b') else ''} c={run.get('color') or '-'} "
                              f"al={bpr.get('align') or '-'}] {bt[:100]}")
                lines.append(f'    {bt}')

    def walk_tbl(tbl, depth, path_ctx):
        grid = tbl.find(W + 'tblGrid')
        widths = [int(gc.get(W + 'w')) for gc in grid.findall(W + 'gridCol')] if grid is not None else []
        rows = tbl.findall(W + 'tr')
        report.append('  ' * depth + f"TABLE {path_ctx} {len(rows)} rows, "
                                     f"grid={[round(w/1440,2) for w in widths]}in")
        lines.append(f"\n<!-- TABLE {path_ctx} -->")
        for ri, tr in enumerate(rows):
            for ci, tc in enumerate(tr.findall(W + 'tc')):
                tcPr = tc.find(W + 'tcPr')
                fill = span = None
                if tcPr is not None:
                    shd = tcPr.find(W + 'shd')
                    if shd is not None:
                        fill = shd.get(W + 'fill')
                    gs = tcPr.find(W + 'gridSpan')
                    if gs is not None:
                        span = gs.get(W + 'val')
                report.append('  ' * (depth + 1) +
                              f"CELL r{ri}c{ci} fill={fill or '-'} span={span or 1}")
                for child in tc:
                    if tag(child) == 'p':
                        emit_para(child, depth + 2, f"{path_ctx}.r{ri}c{ci}")
                    elif tag(child) == 'tbl':
                        walk_tbl(child, depth + 2, f"{path_ctx}.r{ri}c{ci}.t")

    ti = 0
    pi = 0
    for child in body:
        if tag(child) == 'p':
            emit_para(child, 0, f"p{pi}")
            pi += 1
        elif tag(child) == 'tbl':
            walk_tbl(child, 0, f"t{ti}")
            ti += 1
        elif tag(child) == 'sectPr':
            pgSz = child.find(W + 'pgSz')
            mar = child.find(W + 'pgMar')
            report.append(f"SECTPR pgSz={dict((k.split('}')[-1],v) for k,v in pgSz.attrib.items()) if pgSz is not None else None}")
            report.append(f"       pgMar={dict((k.split('}')[-1],v) for k,v in mar.attrib.items()) if mar is not None else None}")

    stem = os.path.splitext(os.path.basename(path))[0]
    outline_path = os.path.join(out_dir, f'{stem}.outline.md')
    struct_path = os.path.join(out_dir, f'{stem}.structure.txt')
    os.makedirs(out_dir, exist_ok=True)
    with open(outline_path, 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(lines))
    with open(struct_path, 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(report))

    if echo:
        print('\n'.join(report))
        print()
    print(f'wrote {outline_path}  ({len(lines)} lines)')
    print(f'wrote {struct_path}  ({len(report)} lines)')

    media = [n for n in z.namelist() if n.startswith('word/media/')]
    print(f'media parts: {len(media)}')
    for m in sorted(media)[:30]:
        print('  ', m, z.getinfo(m).file_size)

    # the values you actually want when deriving a brand
    fills = collections.Counter(re.findall(r'w:fill="([0-9A-Fa-f]{6})"', raw))
    colors = collections.Counter(re.findall(r'<w:color w:val="([0-9A-Fa-f]{6})"', raw))
    sizes = collections.Counter(re.findall(r'<w:sz w:val="(\d+)"', raw))
    print('\nfills      :', ', '.join(f'{k} x{v}' for k, v in fills.most_common(12)) or '-')
    print('text colours:', ', '.join(f'{k} x{v}' for k, v in colors.most_common(12)) or '-')
    print('font sizes  :', ', '.join(f'{int(k)/2:g}pt x{v}'
                                     for k, v in sizes.most_common(12)) or '-')


def main_cli():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('source', help='the .docx to inspect')
    ap.add_argument('--out-dir', default='.', help='where to write the two reports')
    ap.add_argument('--quiet', action='store_true',
                    help='write the files without printing the tree')
    a = ap.parse_args()
    if not os.path.isfile(a.source):
        raise SystemExit(f'not found: {a.source}')
    main(a.source, a.out_dir, echo=not a.quiet)


if __name__ == '__main__':
    main_cli()
