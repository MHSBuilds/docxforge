"""Lift a brand shell out of a document somebody already designed.

The shell is everything that carries the brand and none of the content:
styles.xml, theme, fontTable, numbering, settings and the header/footer bars,
with the body emptied and the header/footer text tokenised so the assembler can
stamp a new period and product name on it.

Use this when you have a document whose look you want to keep. Use
`docxforge.blank` instead when you are starting from nothing.

Run it once per brand. Re-run only when the shell itself changes - a new bar, a
new footer line, a different palette in styles.xml.

    python -m docxforge.template "Release_March.docx" --out brands/ours.docx
"""
import argparse
import json
import os
import re
import shutil
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUT = os.path.join(HERE, '..', 'assets', 'release_template.docx')

# Body reduced to nothing but the section properties. Keeping sectPr verbatim
# preserves A4 size, the asymmetric margins and the header/footer references.
EMPTY_BODY = (
    '<w:body>'
    '<w:p/>'
    '{sectpr}'
    '</w:body>'
)

DROP_PREFIXES = ('word/media/',)

# One clean run of text, tokenised. Built fresh rather than patched onto the
# source run, which may already carry xml:space.
TOKEN_T = '<w:t xml:space="preserve">{{PERIOD}}   •   {{TENANT}}</w:t>'


def extract_sectpr(document_xml):
    m = re.search(r'<w:sectPr\b.*?</w:sectPr>', document_xml, re.S)
    if not m:
        raise SystemExit('no <w:sectPr> found in source document.xml')
    return m.group(0)


def document_shell(document_xml):
    """Keep the <w:document ...> open tag (namespaces!) and swap the body."""
    m = re.match(r'(<\?xml.*?\?>\s*)?(<w:document\b[^>]*>)', document_xml, re.S)
    if not m:
        raise SystemExit('could not parse <w:document> open tag')
    decl = m.group(1) or '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    open_tag = m.group(2)
    sectpr = extract_sectpr(document_xml)
    return decl + open_tag + EMPTY_BODY.format(sectpr=sectpr) + '</w:document>'


def strip_media_rels(rels_xml):
    return re.sub(r'<Relationship\b[^>]*Target="media/[^"]*"[^>]*/>', '', rels_xml)


def tokenise_header(xml):
    """Replace the '<Month> <Year>   •   <Tenant>' run text with tokens.

    Sources often split the line over two runs ('March ' / '2027 - Northwind').
    Collapse to a single token run so substitution is a plain string replace.
    """
    # month-year + separator + tenant, however it is split across runs
    pat = re.compile(
        r'(<w:t[^>]*>)([A-Z][a-z]+\s*)(</w:t>)(.*?)(<w:t[^>]*>)(\d{4}\s*\S\s*.*?)(</w:t>)',
        re.S)
    m = pat.search(xml)
    if m:
        xml = (xml[:m.start()]
               + TOKEN_T
               + m.group(4)
               + '<w:t/>'
               + xml[m.end():])
        return xml, True
    # already a single run
    pat2 = re.compile(r'(<w:t[^>]*>)([A-Z][a-z]+ \d{4}\s*\S\s*[^<]*)(</w:t>)')
    m2 = pat2.search(xml)
    if m2:
        return xml[:m2.start()] + TOKEN_T + xml[m2.end():], True
    return xml, False


FIRST_T = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>')


def tokenise_first_run(xml, token):
    """Swap the first text run for a token, returning the text it replaced.

    In both the header and the footer the brand line is the first run — the
    header puts its title before the tab to period/tenant, the footer puts the
    sender line before the tab to the page field. That makes "first run" a
    sound rule here, and the caller prints what was captured so a different
    document's layout is caught by eye rather than shipped silently.
    """
    m = FIRST_T.search(xml)
    if not m or not m.group(1).strip():
        return xml, None
    return (xml[:m.start()] + f'<w:t xml:space="preserve">{token}</w:t>'
            + xml[m.end():]), m.group(1)


def build(src, out):
    zin = zipfile.ZipFile(src)
    names = zin.namelist()
    if 'word/document.xml' not in names:
        raise SystemExit(f'{src} is not a .docx')

    doc = zin.read('word/document.xml').decode('utf-8')
    shell = document_shell(doc)

    out = os.path.abspath(out)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    tmp = out + '.tmp'
    header_tokenised = False
    found = {'header_title': None, 'footer_left': None}

    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            n = item.filename
            if any(n.startswith(p) for p in DROP_PREFIXES):
                continue
            if n == 'word/document.xml':
                zout.writestr(item, shell.encode('utf-8'))
                continue
            if n == 'word/_rels/document.xml.rels':
                zout.writestr(item, strip_media_rels(zin.read(n).decode('utf-8')))
                continue
            if re.fullmatch(r'word/header\d+\.xml', n):
                xml, ok = tokenise_header(zin.read(n).decode('utf-8'))
                header_tokenised = header_tokenised or ok
                xml, title = tokenise_first_run(xml, '{{HEADER_TITLE}}')
                found['header_title'] = found['header_title'] or title
                zout.writestr(item, xml.encode('utf-8'))
                continue
            if re.fullmatch(r'word/footer\d+\.xml', n):
                xml, left = tokenise_first_run(
                    zin.read(n).decode('utf-8'), '{{FOOTER_LEFT}}')
                found['footer_left'] = found['footer_left'] or left
                zout.writestr(item, xml.encode('utf-8'))
                continue
            zout.writestr(item, zin.read(n))

    shutil.move(tmp, out)
    print(f'wrote {out}')
    print(f'  period/tenant tokenised : {header_tokenised}')
    print(f'  header title captured   : {found["header_title"]!r}')
    print(f'  footer line captured    : {found["footer_left"]!r}')
    print(f'  media dropped           : '
          f'{sum(1 for n in names if n.startswith("word/media/"))}')
    print(f'  parts kept              : '
          f'{sum(1 for n in names if not n.startswith("word/media/"))}')
    if not header_tokenised:
        print('  WARNING: could not find the period/tenant line in the header; '
              'the assembler will leave that text as it stands.')
    for k, v in found.items():
        if v is None:
            print(f'  WARNING: no {k} run found; that text stays as it is in '
                  f'the template.')
    return found


def write_starter_brand(found, path):
    """Emit a brand file pre-filled with what the source document actually said.

    Only the two strings can be read off the document reliably. The palette,
    geometry and type sizes still have to be derived — validation/deep_scan.py
    prints them — so they are seeded with the shipped values and flagged.
    """
    if os.path.exists(path):
        print(f'  starter brand           : {path} exists, left alone')
        return
    shipped = os.path.join(os.path.dirname(os.path.abspath(path)), 'brand.json')
    base = {}
    if os.path.isfile(shipped):
        try:
            with open(shipped, encoding='utf-8') as fh:
                base = json.load(fh)
        except (json.JSONDecodeError, OSError):
            base = {}
    base.pop('_comment', None)
    base['name'] = 'TODO — name this brand'
    base.setdefault('strings', {})
    if found.get('header_title'):
        base['strings']['header_title'] = found['header_title']
    if found.get('footer_left'):
        base['strings']['footer_left'] = found['footer_left']
    base['_comment'] = ('Starter brand. header_title and footer_left were read '
                        'off the source document; everything else is still the '
                        'neutral default. Derive your own values with '
                        '`python -m docxforge.inspect_docx <source.docx>` and '
                        'replace them, then pass this file as --brand.')
    with open(path, 'w', encoding='utf-8') as fh:
        json.dump(base, fh, indent=2, ensure_ascii=False)
        fh.write('\n')
    print(f'  starter brand           : wrote {path}')


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('source', help='a .docx to take the brand shell from')
    ap.add_argument('--out', default=DEFAULT_OUT)
    ap.add_argument('--brand-out', default=None,
                    help='also write a starter brand file (default: '
                         'brand.starter.json beside the template)')
    a = ap.parse_args()
    info = build(a.source, a.out)
    brand_out = a.brand_out or os.path.join(
        os.path.dirname(os.path.abspath(a.out)), 'brand.starter.json')
    write_starter_brand(info, brand_out)
