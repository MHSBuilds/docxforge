"""Build a release-notes .docx from a JSON draft.

Clones a brand shell and writes a fresh word/document.xml into it. The document
is regenerated, never patched, so a later correction never costs you the images.

    python -m docxforge.build draft.json --out Release.docx
    python -m docxforge.build draft.json --shots ./shots --brand brands/slate.json

Exit codes: 0 built (possibly with gaps), 2 the draft is unusable.

Draft JSON - see DRAFT.md. Short version:

{
  "product": "Northwind",
  "period": "March 2027",
  "updates": [
    {"title": "...", "contents_title": "...", "audience": "...",
     "body": [{"kind": "para",   "text": "text with **bold** and *italic*"},
              {"kind": "bullet", "text": "..."},
              {"kind": "shot",   "slot": "update-1a.png", "note": "..."}]}
  ],
  "coming_soon": [{"title": "...", "text": "..."}],
  "gaps": ["Update 3 has no success measure recorded"]
}

Nothing here knows about any particular company: colours, fixed strings, widths
and type sizes all come from a brand file. See BRAND.md.
"""
import argparse
import json
import os
import re
import shutil
import struct
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
TEMPLATE = os.path.join(ROOT, 'brands', 'slate_template.docx')
BRAND = os.path.join(ROOT, 'brands', 'slate.json')

# ------------------------------------------------------------------ defaults
# The neutral "slate" scheme. These are only fallbacks: a brand file (--brand)
# overrides any of them, which is what makes the assembler reusable. Derive real
# values from a document you already have with `python -m docxforge.inspect_docx`
# rather than guessing them here.
#
# configure() rebinds these globals before a build. Module globals rather than an
# object threaded through the call tree on purpose: ~60 call sites read them and
# there is one build per process, so there is nothing to race.
INK = '1F2A37'          # display headings, contents header row
GREEN = '2C6E8F'        # accent - subtitles, number chips, card titles
GREEN_LT = '3E8BB0'     # lighter accent - banner title cell, masthead rule
PALE = 'EDF3F7'         # tint - mastheads, alternating contents rows
PALE2 = 'F7FAFC'        # palest tint - coming-soon cards
BODY = '39434D'         # body text
ORANGE = 'C2410C'       # Coming Soon bar
RULE = 'DDE5EB'         # contents row rule
WHITE = 'FFFFFF'
MUTED = '8A9198'        # placeholders and captions

CONTENT_W = 10466       # twips, = A4 less the 720/720 margins
NUM_W = 560             # number chip column
TITLE_W = CONTENT_W - NUM_W
CONTENTS_NUM_W = 720
CONTENTS_UPDATE_W = 5450
CS_CARD_W = 5083
CS_GUTTER_W = 300

MAX_IMG_EMU = 6543040   # 7.155in - full content width, less a hair
MAX_IMG_H_EMU = 3931920  # 4.30in. Without a height cap a portrait screenshot
                         # renders eight inches tall and swallows a page.
EMU_PER_PX = 9525       # 96 dpi

# Display type sizes, in half-points (56 = 28pt). Body copy is not here: it
# inherits whatever default the template's styles.xml sets.
TYPE = {
    'contents_title': 56, 'contents_title_line': 560, 'contents_subtitle': 24,
    'contents_blurb': 20, 'contents_header': 20, 'contents_number': 22,
    'masthead_title': 72, 'masthead_subtitle': 26, 'masthead_intro': 22,
    'banner_number': 30, 'banner_title': 25,
    'coming_soon_bar': 28, 'coming_soon_card_title': 24,
    'placeholder': 18, 'caption': 17,
}

# Every fixed string the document prints. {period} and {product} interpolate.
STRINGS = {
    'header_title': 'RELEASE NOTES',
    'footer_left': 'Release notes',
    'contents_title': 'Contents',
    'contents_subtitle': '{period} • Who each update is for',
    'contents_blurb': 'A quick guide to every update in this release and the '
                      'teams it affects.',
    'contents_col_number': '#',
    'contents_col_update': 'Update',
    'contents_col_audience': 'Audience',
    'masthead_title': '{product}',
    'masthead_subtitle': '{period} • What’s new this cycle',
    'intro': 'The following updates are now live. Here’s what changed, why '
             'it matters, and what’s coming next.',
    'coming_soon_title': 'Coming Soon',
    'coming_soon_empty': 'Nothing previewed this cycle.',
    'output_default': 'Release.docx',
    'needs_input': '[NEEDS INPUT]',
    'needs_screenshot': '[NEEDS SCREENSHOT]',
}

# Colours baked into the template's own header/footer bars rather than into
# generated body content. CHROME_SHIPPED records what the shipped template
# literally contains, so a brand can remap each one by value at build time
# without the template having to be re-made.
CHROME_SHIPPED = {
    'header_bar_fill': '1F2A37',
    'header_rule': '4B5C6B',
    'header_title_color': 'FFFFFF',
    'header_meta_color': 'C9D3DD',
    'footer_text_color': '8A9198',
    'footer_rule': 'E2E2E2',
}
CHROME = dict(CHROME_SHIPPED)
# how each one appears in the XML, so the swap is precise rather than a blanket
# hex replace that could hit unrelated attributes
CHROME_PATTERNS = {
    'header_bar_fill': 'w:fill="{}"',
    'header_rule': 'w:color="{}"',
    'header_title_color': '<w:color w:val="{}"',
    'header_meta_color': '<w:color w:val="{}"',
    'footer_text_color': '<w:color w:val="{}"',
    'footer_rule': 'w:color="{}"',
}

PALETTE_KEYS = {
    'ink': 'INK', 'green': 'GREEN', 'green_light': 'GREEN_LT',
    'pale': 'PALE', 'pale2': 'PALE2', 'body': 'BODY', 'accent': 'ORANGE',
    'rule': 'RULE', 'white': 'WHITE', 'muted': 'MUTED',
}
GEOMETRY_KEYS = {
    'content_width': 'CONTENT_W', 'number_col_width': 'NUM_W',
    'contents_num_width': 'CONTENTS_NUM_W',
    'contents_update_width': 'CONTENTS_UPDATE_W',
    'cs_card_width': 'CS_CARD_W', 'cs_gutter_width': 'CS_GUTTER_W',
    'max_image_width_emu': 'MAX_IMG_EMU', 'max_image_height_emu': 'MAX_IMG_H_EMU',
}
HEX = re.compile(r'^[0-9A-Fa-f]{6}$')

# Snapshot of the neutral defaults, taken once at import. configure() restores
# these before applying a brand, so a partial brand file falls back to the
# shipped value rather than to whatever the last configure() call happened to
# leave behind. Only the CLI's one-build-per-process makes that leak invisible;
# anything long-lived (a test run, a service, a loop over several brands) hits it.
_DEFAULTS = {name: globals()[name]
             for name in list(PALETTE_KEYS.values()) + list(GEOMETRY_KEYS.values())}
_DEFAULT_TYPE = dict(TYPE)
_DEFAULT_STRINGS = dict(STRINGS)
_DEFAULT_CHROME = dict(CHROME_SHIPPED)


def reset():
    """Restore every brand global to the shipped default."""
    globals().update(_DEFAULTS)
    globals()['TITLE_W'] = _DEFAULTS['CONTENT_W'] - _DEFAULTS['NUM_W']
    TYPE.clear()
    TYPE.update(_DEFAULT_TYPE)
    STRINGS.clear()
    STRINGS.update(_DEFAULT_STRINGS)
    CHROME.clear()
    CHROME.update(_DEFAULT_CHROME)
    globals()['DASHED'] = build_dashed()


def load_brand(path):
    """Read a brand file. Absent or unreadable falls back to the shipped
    defaults — a missing brand.json is never a reason to fail a build."""
    if not path or not os.path.isfile(path):
        return {}, None
    try:
        with open(path, encoding='utf-8') as fh:
            return json.load(fh), None
    except (json.JSONDecodeError, OSError) as e:
        return {}, f'{os.path.basename(path)}: {e}'


def configure(brand):
    """Rebind the module-level brand globals from a brand dict.

    Unknown keys are reported, not silently ignored — a typo in a brand file
    should not quietly leave the previous colour in place.

    Resets to the shipped defaults first, so calling this twice with different
    brands gives the same result as calling it once in a fresh process.
    """
    reset()
    g = globals()
    warnings = []

    for key, val in (brand.get('palette') or {}).items():
        name = PALETTE_KEYS.get(key)
        if not name:
            warnings.append(f'unknown palette key "{key}"')
        elif not HEX.match(str(val)):
            warnings.append(f'palette.{key} is not a 6-digit hex colour: {val!r}')
        else:
            g[name] = str(val).upper()

    for key, val in (brand.get('geometry') or {}).items():
        name = GEOMETRY_KEYS.get(key)
        if not name:
            warnings.append(f'unknown geometry key "{key}"')
        elif not isinstance(val, int) or val <= 0:
            warnings.append(f'geometry.{key} must be a positive integer: {val!r}')
        else:
            g[name] = val
    g['TITLE_W'] = g['CONTENT_W'] - g['NUM_W']

    for key, val in (brand.get('type') or {}).items():
        if key not in TYPE:
            warnings.append(f'unknown type key "{key}"')
        elif not isinstance(val, int) or val <= 0:
            warnings.append(f'type.{key} must be a positive integer: {val!r}')
        else:
            TYPE[key] = val

    for key, val in (brand.get('strings') or {}).items():
        if key not in STRINGS:
            warnings.append(f'unknown string key "{key}"')
        elif not isinstance(val, str):
            warnings.append(f'strings.{key} must be a string')
        else:
            STRINGS[key] = val

    for key, val in (brand.get('chrome') or {}).items():
        if key.startswith('_'):
            continue
        if key not in CHROME_SHIPPED:
            warnings.append(f'unknown chrome key "{key}"')
        elif not HEX.match(str(val)):
            warnings.append(f'chrome.{key} is not a 6-digit hex colour: {val!r}')
        else:
            CHROME[key] = str(val).upper()

    # DASHED is derived from MUTED, so it has to be rebuilt after the rebind
    g['DASHED'] = build_dashed()
    return warnings


def restyle_chrome(xml):
    """Remap the template's header/footer bar colours to the brand's."""
    for key, shipped in CHROME_SHIPPED.items():
        want = CHROME[key]
        if want.upper() == shipped.upper():
            continue
        pat = CHROME_PATTERNS[key]
        xml = xml.replace(pat.format(shipped), pat.format(want))
        xml = xml.replace(pat.format(shipped.lower()), pat.format(want))
    return xml


def s(key, period='', product=''):
    """A brand string with {period}/{product} filled in."""
    try:
        return STRINGS[key].format(period=period, product=product)
    except (KeyError, IndexError):
        return STRINGS[key]      # stray braces in a custom brand: print as-is

NS = (
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    'mc:Ignorable="w14 wp14"'
)


def esc(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace('"', '&quot;'))


# ------------------------------------------------------------------ runs/paras
INLINE = re.compile(r'(\*\*.+?\*\*|\*[^*]+?\*)', re.S)


def runs(text, color=BODY, size=None, bold=False, italic=False,
         caps=False, track=None):
    """Turn text with **bold** / *italic* markers into a run sequence.

    caps/track reproduce the contents header row, which the sent document sets
    in all-caps with 6 twips of letter-spacing.
    """
    out = []
    for part in INLINE.split(str(text)):
        if not part:
            continue
        b, i = bold, italic
        if part.startswith('**') and part.endswith('**') and len(part) > 4:
            part, b = part[2:-2], True
        elif part.startswith('*') and part.endswith('*') and len(part) > 2:
            part, i = part[1:-1], True
        rpr = ''
        if b:
            rpr += '<w:b/><w:bCs/>'
        if i:
            rpr += '<w:i/><w:iCs/>'
        if caps:
            rpr += '<w:caps/>'
        if track is not None:
            rpr += f'<w:spacing w:val="{track}"/>'
        if color:
            rpr += f'<w:color w:val="{color}"/>'
        if size:
            rpr += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
        out.append(f'<w:r><w:rPr>{rpr}</w:rPr>'
                   f'<w:t xml:space="preserve">{esc(part)}</w:t></w:r>')
    return ''.join(out) or '<w:r><w:t/></w:r>'


def para(content, *, spacing='', jc=None, keep=False, ind='', bdr='', shd=''):
    ppr = ''
    if keep:
        ppr += '<w:keepNext/>'
    ppr += bdr + shd + spacing + ind
    if jc:
        ppr += f'<w:jc w:val="{jc}"/>'
    return f'<w:p><w:pPr>{ppr}</w:pPr>{content}</w:p>'


def spacer(line=80, sz=8):
    return (f'<w:p><w:pPr><w:spacing w:line="{line}" w:lineRule="exact"/>'
            f'<w:rPr><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr></w:pPr></w:p>')


def sp(after=None, line=None, rule='auto', before=None):
    bits = []
    if before is not None:
        bits.append(f'w:before="{before}"')
    if after is not None:
        bits.append(f'w:after="{after}"')
    if line is not None:
        bits.append(f'w:line="{line}" w:lineRule="{rule}"')
    return f'<w:spacing {" ".join(bits)}/>' if bits else ''


# ----------------------------------------------------------------- table parts
def no_borders():
    return ('<w:tblBorders>' + ''.join(
        f'<w:{e} w:val="none" w:sz="0" w:space="0" w:color="{WHITE}"/>'
        for e in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')) +
        '</w:tblBorders>')


def tbl_pr(width=CONTENT_W, borders=None, cellmar=10, fixed=True):
    b = borders if borders is not None else no_borders()
    layout = '<w:tblLayout w:type="fixed"/>' if fixed else ''
    return (f'<w:tblPr><w:tblW w:w="{width}" w:type="dxa"/>{b}{layout}'
            f'<w:tblCellMar><w:left w:w="{cellmar}" w:type="dxa"/>'
            f'<w:right w:w="{cellmar}" w:type="dxa"/></w:tblCellMar>'
            f'<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" '
            f'w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr>')


def grid(*widths):
    return '<w:tblGrid>' + ''.join(f'<w:gridCol w:w="{w}"/>' for w in widths) + '</w:tblGrid>'


def tc(content, *, w, fill=None, mar=(0, 0, 0, 0), valign=None, span=None, borders=None):
    pr = f'<w:tcW w:w="{w}" w:type="dxa"/>'
    if span:
        pr += f'<w:gridSpan w:val="{span}"/>'
    if borders:
        pr += borders
    if fill:
        pr += f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
    t, l, b, r = mar
    pr += (f'<w:tcMar><w:top w:w="{t}" w:type="dxa"/><w:left w:w="{l}" w:type="dxa"/>'
           f'<w:bottom w:w="{b}" w:type="dxa"/><w:right w:w="{r}" w:type="dxa"/></w:tcMar>')
    if valign:
        pr += f'<w:vAlign w:val="{valign}"/>'
    return f'<w:tc><w:tcPr>{pr}</w:tcPr>{content}</w:tc>'


def tr(cells, *, cantsplit=False, height=None):
    pr = ''
    if cantsplit:
        pr += '<w:cantSplit/>'
    if height:
        pr += f'<w:trHeight w:val="{height}"/>'
    return f'<w:tr>{f"<w:trPr>{pr}</w:trPr>" if pr else ""}{cells}</w:tr>'


# ---------------------------------------------------------------------- images
def png_size(path):
    with open(path, 'rb') as fh:
        head = fh.read(33)
    if head[:8] != b'\x89PNG\r\n\x1a\n' or head[12:16] != b'IHDR':
        return None
    w, h = struct.unpack('>II', head[16:24])
    return w, h


def jpeg_size(path):
    with open(path, 'rb') as fh:
        if fh.read(2) != b'\xff\xd8':
            return None
        while True:
            b = fh.read(1)
            while b and b != b'\xff':
                b = fh.read(1)
            marker = fh.read(1)
            while marker == b'\xff':
                marker = fh.read(1)
            if not marker:
                return None
            if marker[0] in range(0xc0, 0xd0) and marker[0] not in (0xc4, 0xc8, 0xcc):
                fh.read(3)
                h, w = struct.unpack('>HH', fh.read(4))
                return w, h
            ln = fh.read(2)
            if len(ln) < 2:
                return None
            fh.seek(struct.unpack('>H', ln)[0] - 2, 1)


def image_size(path):
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == '.png':
            return png_size(path)
        if ext in ('.jpg', '.jpeg'):
            return jpeg_size(path)
    except OSError:
        return None
    return None


def drawing(rid, cx, cy, name, docpr_id):
    return (
        '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{docpr_id}" name="{esc(name)}"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/>'
        '</wp:cNvGraphicFramePr>'
        '<a:graphic><a:graphicData '
        'uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{docpr_id}" name="{esc(name)}"/>'
        '<pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/>'
        '</a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')


def build_dashed():
    return ('<w:pBdr>' + ''.join(
        f'<w:{e} w:val="dashSmallGap" w:sz="6" w:space="6" w:color="{MUTED}"/>'
        for e in ('top', 'left', 'bottom', 'right')) + '</w:pBdr>')


DASHED = build_dashed()


def placeholder(slot, caption):
    """A visible, labelled gap. Never a silent omission."""
    text = f'Screenshot slot — {slot}'
    if caption:
        text += f' — {caption}'
    return para(
        runs(text, color=MUTED, size=TYPE['placeholder'], italic=True) +
        runs(f"   {STRINGS['needs_screenshot']}", color=ORANGE,
             size=TYPE['placeholder'], bold=True),
        spacing=sp(before=60, after=160), jc='center', bdr=DASHED,
        shd=f'<w:shd w:val="clear" w:color="auto" w:fill="{PALE2}"/>')


# ------------------------------------------------------------------- the build
class Builder:
    def __init__(self, draft, shots_dir):
        self.d = draft
        self.shots = shots_dir
        self.media = []          # (rid, arcname, srcpath)
        self.placed = []
        self.missing = []
        self._rid = 100
        self._docpr = 1000

    def next_rid(self):
        self._rid += 1
        return f'rIdImg{self._rid}'

    # -- screenshots ------------------------------------------------------
    def shot(self, slot, caption, note, width_in=None):
        """`note` describes the shot for the shot list and the placeholder box.
        `caption` is printed under the image in the document — the August
        document uses none, so leave it unset unless asked for.
        `width_in` shrinks (or enlarges) one image where fit-to-width is wrong,
        e.g. a narrow dialog that should not be blown up to full width.
        """
        label = note or caption
        path = None
        if self.shots and slot:
            cand = os.path.join(self.shots, slot)
            if os.path.isfile(cand):
                path = cand
            else:                      # tolerate a different extension
                stem = os.path.splitext(slot)[0].lower()
                for f in sorted(os.listdir(self.shots)) if os.path.isdir(self.shots) else []:
                    if (os.path.splitext(f)[0].lower() == stem
                            and os.path.splitext(f)[1].lower() in
                            ('.png', '.jpg', '.jpeg')):
                        path = os.path.join(self.shots, f)
                        break
        if not path:
            self.missing.append({'slot': slot, 'caption': label})
            return placeholder(slot, label)

        size = image_size(path)
        if not size:
            self.missing.append({'slot': slot, 'caption': label,
                                 'why': 'unreadable image'})
            return placeholder(slot, label)
        pw, ph = size
        cx = pw * EMU_PER_PX
        cy = ph * EMU_PER_PX
        # fit inside the content box, preserving aspect
        want = int(float(width_in) * 914400) if width_in else None
        caps = [min(want or MAX_IMG_EMU, MAX_IMG_EMU) / cx, MAX_IMG_H_EMU / cy]
        if not want:
            caps.append(1.0)   # never upscale unless an explicit width asks for it
        scale = min(caps)
        cx, cy = int(cx * scale), int(cy * scale)
        rid = self.next_rid()
        ext = os.path.splitext(path)[1].lower().lstrip('.')
        ext = 'jpeg' if ext == 'jpg' else ext
        arc = f'word/media/{os.path.splitext(os.path.basename(slot))[0]}_{self._rid}.{ext}'
        self.media.append((rid, arc, path))
        self.placed.append({'slot': slot, 'file': path,
                            'width_in': round(cx / 914400, 2)})
        self._docpr += 1
        body = drawing(rid, cx, cy, slot, self._docpr)
        out = para(body, spacing=sp(after=60 if caption else 140, line=264))
        if caption:
            out += para(runs(caption, color=MUTED, size=TYPE['caption'], italic=True),
                        spacing=sp(after=140))
        return out

    # -- blocks -----------------------------------------------------------
    def masthead(self, big, sub, blurb, *, big_sz, sub_sz, blurb_sz,
                 mar, rule_bottom=False, big_line=None):
        borders = ('<w:tblBorders>'
                   f'<w:top w:val="none" w:sz="0" w:space="0" w:color="{WHITE}"/>'
                   f'<w:left w:val="none" w:sz="0" w:space="0" w:color="{WHITE}"/>'
                   + (f'<w:bottom w:val="single" w:sz="20" w:space="0" w:color="{GREEN_LT}"/>'
                      if rule_bottom else
                      f'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="{WHITE}"/>')
                   + f'<w:right w:val="none" w:sz="0" w:space="0" w:color="{WHITE}"/>'
                   f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="{WHITE}"/>'
                   f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="{WHITE}"/>'
                   '</w:tblBorders>')
        cell = (para(runs(big, color=INK, size=big_sz, bold=True),
                     spacing=sp(after=40, line=big_line, rule='exact') if big_line
                     else sp(after=40)) +
                para(runs(sub, color=GREEN, size=sub_sz, bold=True),
                     spacing=sp(after=140 if rule_bottom is False else 40)) +
                para(runs(blurb, color=BODY, size=blurb_sz), spacing=sp(after=60)))
        return ('<w:tbl>' + tbl_pr(borders=borders, cellmar=0, fixed=False)
                + grid(CONTENT_W)
                + tr(tc(cell, w=CONTENT_W, fill=PALE, mar=mar)) + '</w:tbl>')

    def contents(self, updates):
        borders = ('<w:tblBorders>'
                   + ''.join(f'<w:{e} w:val="none" w:sz="0" w:space="0" w:color="{WHITE}"/>'
                             for e in ('top', 'left', 'bottom', 'right'))
                   + f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{RULE}"/>'
                   + f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="{WHITE}"/>'
                   '</w:tblBorders>')
        W = (CONTENTS_NUM_W, CONTENTS_UPDATE_W,
             CONTENT_W - CONTENTS_NUM_W - CONTENTS_UPDATE_W)
        mar = (90, 120, 90, 120)
        hmar = (70, 120, 70, 120)

        head = tr(
            tc(para(runs(STRINGS['contents_col_number'], color=WHITE,
                         size=TYPE['contents_header'], bold=True), jc='center'),
               w=W[0], fill=INK, mar=hmar, valign='center') +
            tc(para(runs(STRINGS['contents_col_update'], color=WHITE,
                         size=TYPE['contents_header'], bold=True,
                         caps=True, track=6)),
               w=W[1], fill=INK, mar=hmar, valign='center') +
            tc(para(runs(STRINGS['contents_col_audience'], color=WHITE,
                         size=TYPE['contents_header'], bold=True,
                         caps=True, track=6)),
               w=W[2], fill=INK, mar=hmar, valign='center'),
            cantsplit=True, height=300)

        rows = [head]
        for i, u in enumerate(updates, 1):
            fill = WHITE if i % 2 else PALE
            anchor = f'sec{i}'
            num = para(
                f'<w:hyperlink w:anchor="{anchor}" w:history="1">'
                + runs(str(i), color=WHITE, size=TYPE['contents_number'],
                       bold=True) + '</w:hyperlink>',
                jc='center')
            title_txt = u.get('contents_title') or u.get('title') or ''
            link = (f'<w:hyperlink w:anchor="{anchor}" w:history="1"><w:r>'
                    f'<w:rPr><w:b/><w:bCs/><w:color w:val="{INK}"/>'
                    f'<w:u w:val="single"/></w:rPr>'
                    f'<w:t xml:space="preserve">{esc(title_txt)}</w:t></w:r></w:hyperlink>')
            rows.append(tr(
                tc(num, w=W[0], fill=GREEN, mar=mar, valign='center') +
                tc(para(link, spacing=sp(line=252)), w=W[1], fill=fill, mar=mar,
                   valign='center') +
                tc(para(runs(u.get('audience') or STRINGS['needs_input']),
                        spacing=sp(line=252)),
                   w=W[2], fill=fill, mar=mar, valign='center'),
                cantsplit=True, height=300))
        return ('<w:tbl>' + tbl_pr(borders=borders, cellmar=120, fixed=False)
                + grid(*W) + ''.join(rows) + '</w:tbl>')

    def update(self, n, u):
        banner = tr(
            tc(para(runs(str(n), color=WHITE, size=TYPE['banner_number'], bold=True),
                    jc='center', keep=True),
               w=NUM_W, fill=GREEN, mar=(60, 40, 60, 40), valign='center') +
            tc(f'<w:p><w:pPr><w:keepNext/></w:pPr>'
               f'<w:bookmarkStart w:id="{n}" w:name="sec{n}"/>'
               + runs(u.get('title') or STRINGS['needs_input'], color=WHITE,
                        size=TYPE['banner_title'], bold=True)
               + f'<w:bookmarkEnd w:id="{n}"/></w:p>',
               w=TITLE_W, fill=GREEN_LT, mar=(90, 160, 90, 120), valign='center'),
            cantsplit=True)

        blocks = u.get('body') or []
        if not blocks:
            blocks = [{'kind': 'para',
                       'text': f"{STRINGS['needs_input']} — no release copy "
                               'drafted for this update.'}]
        out, first_para_done = [], False
        for b in blocks:
            kind = (b.get('kind') or 'para').lower()
            if kind == 'shot':
                out.append(self.shot(b.get('slot') or '', b.get('caption') or '',
                                     b.get('note') or '', b.get('width_in')))
            elif kind == 'bullet':
                out.append(para(runs('•  ') + runs(b.get('text') or ''),
                                spacing=sp(after=40, line=264),
                                ind='<w:ind w:left="202"/>'))
            else:
                out.append(para(runs(b.get('text') or ''),
                                spacing=sp(after=120, line=264),
                                keep=not first_para_done))
                first_para_done = True
        body_row = tr(tc(''.join(out), w=CONTENT_W, span=2,
                         mar=(120, 0, 40, 0)))
        return ('<w:tbl>' + tbl_pr(cellmar=0) + grid(NUM_W, TITLE_W)
                + banner + body_row + '</w:tbl>')

    def coming_soon(self, items):
        bar = ('<w:tbl>' + tbl_pr(cellmar=10, fixed=False) + grid(CONTENT_W) +
               tr(tc(para(runs(STRINGS['coming_soon_title'], color=WHITE,
                               size=TYPE['coming_soon_bar'], bold=True),
                          keep=True),
                     w=CONTENT_W, fill=ORANGE, mar=(90, 200, 90, 160))) + '</w:tbl>')
        if not items:
            return bar + spacer(160, 16) + para(
                runs(STRINGS['coming_soon_empty'], color=MUTED,
                     size=TYPE['contents_blurb'], italic=True))

        card_borders = ('<w:tcBorders>'
                        f'<w:top w:val="single" w:sz="18" w:space="0" w:color="{ORANGE}"/>'
                        + ''.join(f'<w:{e} w:val="none" w:sz="0" w:space="0" '
                                  f'w:color="{WHITE}"/>'
                                  for e in ('left', 'bottom', 'right'))
                        + '</w:tcBorders>')
        blank_borders = ('<w:tcBorders>'
                         + ''.join(f'<w:{e} w:val="none" w:sz="0" w:space="0" '
                                   f'w:color="{WHITE}"/>'
                                   for e in ('top', 'left', 'bottom', 'right'))
                         + '</w:tcBorders>')

        def card(item, keep):
            """keep=True on every row but the last drags the whole grid onto one
            page, so a lone card is never stranded overleaf."""
            if item is None:
                return tc('<w:p/>', w=CS_CARD_W, mar=(0, 0, 0, 0),
                          borders=blank_borders)
            return tc(
                para(runs(item.get('title') or STRINGS['needs_input'], color=GREEN,
                          size=TYPE['coming_soon_card_title'], bold=True),
                     spacing=sp(after=60), keep=keep) +
                # card body inherits the template's default colour rather
                # than the body colour used in update paragraphs
                para(runs(item.get('text') or STRINGS['needs_input'], color=None),
                     keep=keep),
                w=CS_CARD_W, fill=PALE2, mar=(120, 160, 140, 160),
                borders=card_borders)

        def gutter(borders=blank_borders):
            return tc('<w:p/>', w=CS_GUTTER_W, mar=(0, 0, 0, 0), borders=borders)

        def gap_cell(keep):
            # the keepNext must continue through the spacer rows, or the chain
            # holding the grid on one page breaks here
            return tc('<w:p><w:pPr>' + ('<w:keepNext/>' if keep else '') +
                      '<w:spacing w:line="109" w:lineRule="exact"/>'
                      '<w:rPr><w:sz w:val="10"/></w:rPr></w:pPr></w:p>',
                      w=CS_CARD_W, mar=(0, 0, 0, 0), borders=blank_borders)

        rows = []
        pairs = [(items[i], items[i + 1] if i + 1 < len(items) else None)
                 for i in range(0, len(items), 2)]
        for idx, (left, right) in enumerate(pairs):
            if idx:
                rows.append(tr(gap_cell(True) + gutter() + gap_cell(True),
                               height=109))
            keep = idx < len(pairs) - 1
            rows.append(tr(card(left, keep) + gutter() + card(right, keep),
                           cantsplit=True))
        gridt = ('<w:tbl>' + tbl_pr(cellmar=10) + grid(CS_CARD_W, CS_GUTTER_W, CS_CARD_W)
                 + ''.join(rows) + '</w:tbl>')
        return bar + gridt

    # -- document ---------------------------------------------------------
    def body_xml(self, sectpr):
        d = self.d
        period = d.get('period') or ''
        product = d.get('product') or ''
        ups = d.get('updates') or []

        parts = [
            self.masthead(
                s('contents_title', period, product),
                s('contents_subtitle', period, product),
                d.get('contents_blurb') or s('contents_blurb', period, product),
                big_sz=TYPE['contents_title'], sub_sz=TYPE['contents_subtitle'],
                blurb_sz=TYPE['contents_blurb'], mar=(200, 240, 200, 240),
                rule_bottom=True, big_line=TYPE['contents_title_line']),
            spacer(160, 16),
            self.contents(ups),
            '<w:p><w:r><w:br w:type="page"/></w:r></w:p>',
            self.masthead(
                s('masthead_title', period, product),
                s('masthead_subtitle', period, product),
                d.get('intro') or s('intro', period, product),
                big_sz=TYPE['masthead_title'], sub_sz=TYPE['masthead_subtitle'],
                blurb_sz=TYPE['masthead_intro'], mar=(240, 260, 40, 260),
                rule_bottom=False),
            spacer(120, 12),
        ]
        for i, u in enumerate(ups, 1):
            parts.append(self.update(i, u))
            parts.append(spacer(80, 8))
        parts.append(spacer(160, 16))
        parts.append(self.coming_soon(d.get('coming_soon') or []))
        parts.append('<w:p/>')
        return ('<w:body>' + ''.join(parts) + sectpr + '</w:body>')


# ------------------------------------------------------------------- packaging
def substitute_tokens(xml, period, product):
    """Fill the header/footer tokens the template carries.

    A template from `docxforge.blank` or `docxforge.template` carries
    {{PERIOD}}/{{PRODUCT}} in the header and {{HEADER_TITLE}}/{{FOOTER_LEFT}}
    where the brand text sits. A shell missing any of them simply keeps its own
    text - the replace is a no-op, so any shell still builds.
    """
    xml = (xml.replace('{{PERIOD}}', esc(period or ''))
              .replace('{{PRODUCT}}', esc(product or ''))
              .replace('{{TENANT}}', esc(product or ''))
              .replace('{{HEADER_TITLE}}', esc(s('header_title', period, product)))
              .replace('{{FOOTER_LEFT}}', esc(s('footer_left', period, product))))
    return restyle_chrome(xml)


CT_DEFAULTS = {'png': 'image/png', 'jpeg': 'image/jpeg', 'gif': 'image/gif'}


def ensure_content_types(xml, exts):
    for ext in exts:
        ct = CT_DEFAULTS.get(ext)
        if ct and f'Extension="{ext}"' not in xml:
            xml = xml.replace('</Types>',
                              f'<Default Extension="{ext}" ContentType="{ct}"/></Types>')
    return xml


IMG_RELTYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'


def add_rels(xml, media):
    add = ''.join(
        f'<Relationship Id="{rid}" Type="{IMG_RELTYPE}" '
        f'Target="{arc[len("word/"):]}"/>'
        for rid, arc, _ in media)
    return xml.replace('</Relationships>', add + '</Relationships>')


def extract_sectpr(document_xml):
    m = re.search(r'<w:sectPr\b.*?</w:sectPr>', document_xml, re.S)
    if not m:
        raise SystemExit('template has no <w:sectPr>')
    return m.group(0)


def build(draft, template, out, shots_dir):
    if not os.path.isfile(template):
        print(f'ERROR: template not found: {template}\n'
              f'Generate one: python -m docxforge.blank --out template.docx',
              file=sys.stderr)
        return 2
    zin = zipfile.ZipFile(template)
    tpl_doc = zin.read('word/document.xml').decode('utf-8')
    sectpr = extract_sectpr(tpl_doc)

    b = Builder(draft, shots_dir)
    body = b.body_xml(sectpr)
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                f'<w:document {NS}>{body}</w:document>')

    exts = {os.path.splitext(arc)[1].lstrip('.').lower() for _, arc, _ in b.media}
    out = os.path.abspath(out)
    os.makedirs(os.path.dirname(out) or '.', exist_ok=True)
    tmp = out + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            n = item.filename
            if n == 'word/document.xml':
                zout.writestr(item, document.encode('utf-8'))
            elif n == 'word/_rels/document.xml.rels':
                zout.writestr(item, add_rels(zin.read(n).decode('utf-8'), b.media))
            elif n == '[Content_Types].xml':
                zout.writestr(item, ensure_content_types(
                    zin.read(n).decode('utf-8'), exts))
            elif re.fullmatch(r'word/(header|footer)\d+\.xml', n):
                zout.writestr(item, substitute_tokens(
                    zin.read(n).decode('utf-8'),
                    draft.get('period'), draft.get('product')).encode('utf-8'))
            else:
                zout.writestr(item, zin.read(n))
        for _, arc, src in b.media:
            zout.write(src, arc)
    shutil.move(tmp, out)
    return b, out


# -------------------------------------------------------------------- reporting
def report(b, out, draft):
    ups = draft.get('updates') or []
    cs = draft.get('coming_soon') or []
    total_slots = len(b.placed) + len(b.missing)
    print(f'Done — {os.path.basename(out)}')
    print()
    print(f'{len(ups)} update(s), {len(cs)} Coming Soon item(s), contents table built.')
    if total_slots:
        print(f'{len(b.placed)} of {total_slots} screenshot slot(s) filled'
              + (f', {len(b.missing)} still empty.' if b.missing else '.'))
    else:
        print('No screenshot slots in this draft.')

    needs = [g for g in (draft.get('gaps') or [])]
    body_text = json.dumps(draft, ensure_ascii=False)
    n_marker = body_text.count(STRINGS['needs_input'])
    if b.missing or needs or n_marker:
        print('\nNeeds your input before sending:')
        if b.missing:
            print(f'  - {len(b.missing)} screenshot(s) (shot list below)')
        for g in needs:
            print(f'  - {g}')
        if n_marker:
            print(f"  - {n_marker} {STRINGS['needs_input']} marker(s) left in "
                  'the copy')
    if b.missing:
        print('\nShot list — drop these into the shots folder and re-run')
        for m in b.missing:
            note = f"  ({m['why']})" if m.get('why') else ''
            print(f"  {m['slot']:<16} {m.get('caption', '')}{note}")
    if b.placed:
        print('\nPlaced:')
        for p in b.placed:
            print(f"  {p['slot']:<16} {p['width_in']}in wide")
    if not b.missing and not needs and not n_marker:
        print('\nNo gaps. Review and send.')


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('draft', help='draft JSON')
    ap.add_argument('--shots', default=None,
                    help='folder of images named after the slots (optional)')
    ap.add_argument('--out', default=None)
    ap.add_argument('--template', default=TEMPLATE)
    ap.add_argument('--brand', default=BRAND,
                    help='brand file: palette, chrome, strings, geometry, type '
                         'sizes (default: brands/slate.json)')
    ap.add_argument('--json-report', default=None,
                    help='also write the gap report as JSON')
    a = ap.parse_args()

    brand, err = load_brand(a.brand)
    if err:
        print(f'note: ignoring unreadable brand file — {err}\n'
              f'      building with the shipped defaults.\n', file=sys.stderr)
    for w in configure(brand):
        print(f'brand warning: {w}', file=sys.stderr)

    try:
        draft = json.load(open(a.draft, encoding='utf-8'))
    except (json.JSONDecodeError, OSError) as e:
        print(f'ERROR: cannot read draft: {e}', file=sys.stderr)
        return 2
    if not draft.get('updates') and not draft.get('coming_soon'):
        print('ERROR: draft has neither updates nor coming_soon', file=sys.stderr)
        return 2

    out = a.out or draft.get('output') or STRINGS['output_default']
    if a.shots and not os.path.isdir(a.shots):
        print(f'note: shots folder not found ({a.shots}) — building with '
              f'placeholders instead.\n')
        a.shots = None

    res = build(draft, a.template, out, a.shots)
    if res == 2:
        return 2
    b, out = res
    report(b, out, draft)
    if a.json_report:
        with open(a.json_report, 'w', encoding='utf-8') as fh:
            json.dump({'output': out, 'placed': b.placed, 'missing': b.missing,
                       'gaps': draft.get('gaps') or []}, fh, indent=2)
    return 0


if __name__ == '__main__':
    sys.exit(main())
