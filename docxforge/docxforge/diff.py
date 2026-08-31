"""Diff two .docx files by what they say and how they look, not by bytes.

Reduces each to a normalised outline - text runs with their size, colour and
cell fill, plus images with their real display dimensions - and compares those.
Whitespace, smart quotes, revision ids and media filenames are normalised away,
so what is left is a difference that would show on the page.

Built for one job: prove a generated document matches a reference one.

    python -m docxforge.diff reference.docx generated.docx
"""
import difflib
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def tag(e):
    return e.tag.split('}')[-1]


def norm(s):
    s = s.replace('’', "'").replace('‘', "'")
    s = s.replace('“', '"').replace('”', '"')
    s = s.replace('–', '—').replace('—', '—')
    s = re.sub(r'\s+', ' ', s)
    return s.strip()


def outline(path):
    z = zipfile.ZipFile(path)
    rels = {r.get('Id'): r.get('Target')
            for r in ET.fromstring(z.read('word/_rels/document.xml.rels'))}
    root = ET.fromstring(z.read('word/document.xml'))
    body = root.find(W + 'body')
    out = []
    img_n = [0]

    def do_para(p, fill):
        txt = norm(''.join(t.text or '' for t in p.iter(W + 't')))
        # first run size / colour, for style comparison
        sz = col = None
        for r in p.iter(W + 'r'):
            rPr = r.find(W + 'rPr')
            if rPr is None:
                continue
            e = rPr.find(W + 'sz')
            c = rPr.find(W + 'color')
            sz = e.get(W + 'val') if e is not None else None
            col = c.get(W + 'val') if c is not None else None
            break
        if txt:
            style = f'{sz or "-"}/{col or "-"}' + (f'/bg{fill}' if fill else '')
            out.append(f'TEXT [{style}] {txt}')
        for blip in p.iter(A + 'blip'):
            img_n[0] += 1
            ext = None
            for e in p.iter('{http://schemas.openxmlformats.org/drawingml/'
                            '2006/wordprocessingDrawing}extent'):
                ext = f'{int(e.get("cx"))/914400:.2f}x{int(e.get("cy"))/914400:.2f}in'
            out.append(f'IMAGE #{img_n[0]} {ext}')
        for box in p.iter(W + 'txbxContent'):
            for bp in box.findall(W + 'p'):
                bt = norm(''.join(t.text or '' for t in bp.iter(W + 't')))
                if bt:
                    out.append(f'TEXT [txbox] {bt}')
        if p.find(W + 'r') is not None:
            for br in p.iter(W + 'br'):
                if br.get(W + 'type') == 'page':
                    out.append('PAGEBREAK')

    def do_tbl(t):
        out.append('TABLE')
        for trow in t.findall(W + 'tr'):
            for tcell in trow.findall(W + 'tc'):
                tcPr = tcell.find(W + 'tcPr')
                fill = None
                if tcPr is not None:
                    shd = tcPr.find(W + 'shd')
                    if shd is not None:
                        f = shd.get(W + 'fill')
                        fill = None if f in (None, 'auto') else f
                for ch in tcell:
                    if tag(ch) == 'p':
                        do_para(ch, fill)
                    elif tag(ch) == 'tbl':
                        do_tbl(ch)

    for child in body:
        if tag(child) == 'p':
            do_para(child, None)
        elif tag(child) == 'tbl':
            do_tbl(child)
    return out


def text_only(lines):
    return [l.split('] ', 1)[1] for l in lines if l.startswith('TEXT')]


def main(a, b):
    oa, ob = outline(a), outline(b)
    ta, tb = text_only(oa), text_only(ob)

    print('=' * 78)
    print(f'A (sent)      {os.path.basename(a)}   {len(oa)} blocks, '
          f'{len(ta)} text runs, {sum(1 for l in oa if l.startswith("IMAGE"))} images')
    print(f'B (generated) {os.path.basename(b)}   {len(ob)} blocks, '
          f'{len(tb)} text runs, {sum(1 for l in ob if l.startswith("IMAGE"))} images')
    print('=' * 78)

    sm = difflib.SequenceMatcher(None, ta, tb)
    print(f'\nTEXT similarity: {sm.ratio() * 100:.1f}%')
    only_a = [t for t in ta if t not in tb]
    only_b = [t for t in tb if t not in ta]
    print(f'  in sent but not generated: {len(only_a)}')
    for t in only_a:
        print(f'    - {t[:110]}')
    print(f'  in generated but not sent: {len(only_b)}')
    for t in only_b:
        print(f'    + {t[:110]}')

    print('\nIMAGE sizes  (sent -> generated)')
    ia = [l for l in oa if l.startswith('IMAGE')]
    ib = [l for l in ob if l.startswith('IMAGE')]
    for i in range(max(len(ia), len(ib))):
        x = ia[i].split(' ', 2)[2] if i < len(ia) else '(none)'
        y = ib[i].split(' ', 2)[2] if i < len(ib) else '(none)'
        flag = '' if x == y else '   <-- differs'
        print(f'  {i + 1:>2}  {x:<16} -> {y:<16}{flag}')

    print('\nPALETTE / SIZE usage')
    def styles(lines):
        c = {}
        for l in lines:
            m = re.match(r'TEXT \[([^\]]+)\]', l)
            if m:
                c[m.group(1)] = c.get(m.group(1), 0) + 1
        return c
    sa, sb = styles(oa), styles(ob)
    keys = sorted(set(sa) | set(sb))
    for k in keys:
        x, y = sa.get(k, 0), sb.get(k, 0)
        flag = '' if x == y else '   <-- differs'
        print(f'  {k:<22} sent {x:>3}   generated {y:>3}{flag}')
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1], sys.argv[2]))
