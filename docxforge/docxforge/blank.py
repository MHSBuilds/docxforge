"""Generate a brand shell from scratch — no source document required.

`template.py` lifts a shell out of a document somebody already designed. This
module builds one from nothing: a valid, minimal .docx carrying page geometry,
a default font, and tokenised header and footer bars, and nothing else.

Use it to start a new document family, or as the neutral default so the
assembler works out of the box with no template to supply.

    python -m docxforge.blank --out template.docx
    python -m docxforge.blank --out t.docx --brand brands/default.json

The header/footer text is tokenised, not baked: {{HEADER_TITLE}},
{{FOOTER_LEFT}}, {{PERIOD}}, {{TENANT}}. build.py fills them from the brand.
"""
import argparse
import json
import os
import zipfile

# A4 in twips, with the margins the generated layout assumes.
PAGE = {'w': 11906, 'h': 16838,
        'top': 1180, 'right': 720, 'bottom': 760, 'left': 720,
        'header': 360, 'footer': 320}

DEFAULT_FONT = 'Calibri'
DEFAULT_SIZE = 21        # half-points, 10.5pt
DEFAULT_COLOR = '27322F'

W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/'
        'wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'mc:Ignorable=""')

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

CONTENT_TYPES = DECL + (
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package'
    '.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="png" ContentType="image/png"/>'
    '<Default Extension="jpeg" ContentType="image/jpeg"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd'
    '.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd'
    '.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/header1.xml" ContentType="application/vnd'
    '.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd'
    '.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd'
    '.openxmlformats-package.core-properties+xml"/>'
    '</Types>')

ROOT_RELS = DECL + (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
    'relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
    'officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/'
    'relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    '</Relationships>')

DOC_RELS = DECL + (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
    'relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
    'officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
    'officeDocument/2006/relationships/header" Target="header1.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/'
    'officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
    '</Relationships>')

CORE = DECL + (
    '<cp:coreProperties '
    'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/'
    'core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">'
    '<dc:title>Release notes</dc:title>'
    '<dc:creator>docxforge</dc:creator>'
    '</cp:coreProperties>')


def styles_xml(font, size, color):
    return DECL + (
        f'<w:styles {W_NS}>'
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" '
        f'w:cs="{font}"/>'
        f'<w:color w:val="{color}"/>'
        f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
        '<w:lang w:val="en-GB"/>'
        '</w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr>'
        '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
        '</w:pPr></w:pPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/><w:qFormat/></w:style>'
        '</w:styles>')


def sectpr():
    p = PAGE
    return (
        '<w:sectPr>'
        '<w:headerReference w:type="default" r:id="rId2"/>'
        '<w:footerReference w:type="default" r:id="rId3"/>'
        '<w:type w:val="continuous"/>'
        f'<w:pgSz w:w="{p["w"]}" w:h="{p["h"]}"/>'
        f'<w:pgMar w:top="{p["top"]}" w:right="{p["right"]}" '
        f'w:bottom="{p["bottom"]}" w:left="{p["left"]}" '
        f'w:header="{p["header"]}" w:footer="{p["footer"]}" w:gutter="0"/>'
        '<w:cols w:space="720"/>'
        '</w:sectPr>')


def document_xml():
    return DECL + f'<w:document {W_NS}><w:body><w:p/>{sectpr()}</w:body></w:document>'


def header_xml(bar_fill, title_color, meta_color, rule_color, content_w):
    """A full-width coloured bar: title left, period/tenant right."""
    tab_at = content_w - 566          # right tab, inside the cell padding
    return DECL + (
        f'<w:hdr {W_NS}>'
        '<w:tbl><w:tblPr>'
        f'<w:tblW w:w="{content_w}" w:type="dxa"/>'
        '<w:tblBorders>'
        + ''.join(f'<w:{e} w:val="none" w:sz="0" w:space="0" w:color="FFFFFF"/>'
                  for e in ('top', 'left', 'right', 'insideH', 'insideV'))
        + f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="{rule_color}"/>'
        '</w:tblBorders>'
        '<w:tblCellMar><w:left w:w="160" w:type="dxa"/>'
        '<w:right w:w="160" w:type="dxa"/></w:tblCellMar>'
        '<w:tblLook w:val="0000"/>'
        '</w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="{content_w}"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr>'
        f'<w:tcW w:w="{content_w}" w:type="dxa"/>'
        f'<w:shd w:val="clear" w:color="auto" w:fill="{bar_fill}"/>'
        '<w:tcMar><w:top w:w="90" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/>'
        '</w:tcMar><w:vAlign w:val="center"/>'
        '</w:tcPr>'
        '<w:p><w:pPr>'
        f'<w:tabs><w:tab w:val="right" w:pos="{tab_at}"/></w:tabs>'
        '</w:pPr>'
        '<w:r><w:rPr><w:b/><w:bCs/>'
        f'<w:color w:val="{title_color}"/><w:spacing w:val="12"/>'
        '<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
        '<w:t>{{HEADER_TITLE}}</w:t></w:r>'
        f'<w:r><w:rPr><w:color w:val="{meta_color}"/>'
        '<w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:tab/>'
        '<w:t xml:space="preserve">{{PERIOD}}   •   {{TENANT}}</w:t></w:r>'
        '</w:p></w:tc></w:tr></w:tbl>'
        '<w:p><w:pPr><w:spacing w:line="20" w:lineRule="exact"/></w:pPr></w:p>'
        '</w:hdr>')


def footer_xml(text_color, rule_color, content_w):
    return DECL + (
        f'<w:ftr {W_NS}>'
        '<w:p><w:pPr>'
        f'<w:pBdr><w:top w:val="single" w:sz="6" w:space="6" '
        f'w:color="{rule_color}"/></w:pBdr>'
        f'<w:tabs><w:tab w:val="right" w:pos="{content_w}"/></w:tabs>'
        '<w:spacing w:before="60"/></w:pPr>'
        f'<w:r><w:rPr><w:color w:val="{text_color}"/>'
        '<w:sz w:val="15"/><w:szCs w:val="15"/></w:rPr>'
        '<w:t>{{FOOTER_LEFT}}</w:t></w:r>'
        f'<w:r><w:rPr><w:color w:val="{text_color}"/>'
        '<w:sz w:val="15"/><w:szCs w:val="15"/></w:rPr>'
        '<w:tab/><w:t xml:space="preserve">Page </w:t></w:r>'
        f'<w:r><w:rPr><w:color w:val="{text_color}"/><w:sz w:val="15"/>'
        '<w:szCs w:val="15"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:rPr><w:color w:val="{text_color}"/><w:sz w:val="15"/>'
        '<w:szCs w:val="15"/></w:rPr><w:instrText>PAGE</w:instrText></w:r>'
        f'<w:r><w:rPr><w:color w:val="{text_color}"/><w:sz w:val="15"/>'
        '<w:szCs w:val="15"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r><w:rPr><w:color w:val="{text_color}"/><w:sz w:val="15"/>'
        '<w:szCs w:val="15"/></w:rPr><w:t>1</w:t></w:r>'
        f'<w:r><w:rPr><w:color w:val="{text_color}"/><w:sz w:val="15"/>'
        '<w:szCs w:val="15"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>'
        '</w:p></w:ftr>')


def content_width():
    return PAGE['w'] - PAGE['left'] - PAGE['right']


def build(out, brand=None):
    brand = brand or {}
    chrome = brand.get('chrome') or {}
    typo = brand.get('typography') or {}
    cw = content_width()

    parts = {
        '[Content_Types].xml': CONTENT_TYPES,
        '_rels/.rels': ROOT_RELS,
        'docProps/core.xml': CORE,
        'word/document.xml': document_xml(),
        'word/_rels/document.xml.rels': DOC_RELS,
        'word/styles.xml': styles_xml(
            typo.get('font', DEFAULT_FONT),
            typo.get('size', DEFAULT_SIZE),
            typo.get('color', DEFAULT_COLOR)),
        'word/header1.xml': header_xml(
            chrome.get('header_bar_fill', '1F2A37'),
            chrome.get('header_title_color', 'FFFFFF'),
            chrome.get('header_meta_color', 'C9D3DD'),
            chrome.get('header_rule', '4B5C6B'), cw),
        'word/footer1.xml': footer_xml(
            chrome.get('footer_text_color', '8A9198'),
            chrome.get('footer_rule', 'E2E2E2'), cw),
    }

    out = os.path.abspath(out)
    os.makedirs(os.path.dirname(out) or '.', exist_ok=True)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, xml in parts.items():
            z.writestr(name, xml.encode('utf-8'))
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--out', required=True)
    ap.add_argument('--brand', default=None,
                    help='take the bar colours and default font from this brand file')
    a = ap.parse_args()
    brand = {}
    if a.brand:
        with open(a.brand, encoding='utf-8') as fh:
            brand = json.load(fh)
    path = build(a.out, brand)
    print(f'wrote {path}')
    print(f'  A4, content width {content_width()} twips')
    print('  tokens: {{HEADER_TITLE}} {{FOOTER_LEFT}} {{PERIOD}} {{TENANT}}')


if __name__ == '__main__':
    main()
