# The brand file

Everything about how the document looks. Pass it with `--brand`; the default is
`brands/slate.json`.

```bash
python -m docxforge.build draft.json --brand brands/ours.json --out Release.docx
```

Five blocks, all optional, and every key inside them optional too. What you
leave out keeps the shipped value, so a partial file is normal — override the
six things you care about and ignore the rest. `brands/amber.json` is a partial
file.

## `strings`

Every fixed word the document prints. `{period}` and `{product}` are filled from
the draft.

| Key | Default | Where it appears |
|---|---|---|
| `header_title` | `RELEASE NOTES` | Left of the header bar, every page |
| `footer_left` | `Release notes` | Left of the footer, every page |
| `contents_title` | `Contents` | The first masthead |
| `contents_subtitle` | `{period} • Who each update is for` | Under it |
| `contents_blurb` | *(a sentence)* | Under that. A draft's `contents_blurb` wins |
| `contents_col_number` | `#` | Contents table heading |
| `contents_col_update` | `Update` | Contents table heading |
| `contents_col_audience` | `Audience` | Contents table heading |
| `masthead_title` | `{product}` | The second masthead, large |
| `masthead_subtitle` | `{period} • What's new this cycle` | Under it |
| `intro` | *(two sentences)* | Under that. A draft's `intro` wins |
| `coming_soon_title` | `Coming Soon` | The accent bar |
| `coming_soon_empty` | `Nothing previewed this cycle.` | When there are no cards |
| `output_default` | `Release.docx` | Filename when neither `--out` nor `output` is given |
| `needs_input` | `[NEEDS INPUT]` | The missing-value marker |
| `needs_screenshot` | `[NEEDS SCREENSHOT]` | The missing-image marker |

## `palette`

The ten colours used by generated body content. Six hex digits, no `#`.

| Key | Default | Used for |
|---|---|---|
| `ink` | `1F2A37` | Display headings, contents header row fill |
| `green` | `2C6E8F` | Masthead subtitles, number chips, card titles |
| `green_light` | `3E8BB0` | Section banner fill, rule under the first masthead |
| `pale` | `EDF3F7` | Masthead fill, alternating contents rows |
| `pale2` | `F7FAFC` | Cards, placeholder fill |
| `body` | `39434D` | Body paragraphs |
| `accent` | `C2410C` | Coming Soon bar, card top rules, the needs-screenshot marker |
| `rule` | `DDE5EB` | Contents row rules |
| `white` | `FFFFFF` | Text on filled cells |
| `muted` | `8A9198` | Placeholders and captions |

`green` and `green_light` are historical names for "accent" and "lighter
accent" — they carry no colour of their own.

## `chrome`

Six colours that live in the template's own header and footer bars rather than
in generated content. They are remapped by value at build time, so you can
restyle the bars without rebuilding the shell.

`header_bar_fill`, `header_rule`, `header_title_color`, `header_meta_color`,
`footer_text_color`, `footer_rule`.

## `geometry`

Twips (1/1440 inch), except the image box which is EMU (1/914400 inch).

| Key | Default | |
|---|---|---|
| `content_width` | `10466` | Must equal page width less both margins, or tables will not line up |
| `number_col_width` | `560` | The numbered chip on a section banner |
| `contents_num_width` | `720` | Contents column 1 |
| `contents_update_width` | `5450` | Contents column 2; column 3 takes the remainder |
| `cs_card_width` | `5083` | A Coming Soon card |
| `cs_gutter_width` | `300` | Between the two card columns |
| `max_image_width_emu` | `6543040` | 7.155 in |
| `max_image_height_emu` | `3931920` | 4.30 in |

Page size and margins come from the template's `sectPr`, not from here. If you
lift a shell from a document with different margins, update `content_width` to
match or every table will overhang.

## `type`

Display type sizes in **half-points** — `56` is 28pt. Body copy is not here; it
inherits whatever default the template's `styles.xml` sets.

`contents_title` (56), `contents_title_line` (560, an exact line height),
`contents_subtitle` (24), `contents_blurb` (20), `contents_header` (20),
`contents_number` (22), `masthead_title` (72), `masthead_subtitle` (26),
`masthead_intro` (22), `banner_number` (30), `banner_title` (25),
`coming_soon_bar` (28), `coming_soon_card_title` (24), `placeholder` (18),
`caption` (17).

## `typography`

Read only by `docxforge.blank` when generating a shell from scratch:
`font` (`Calibri`), `size` (21 half-points), `color` (`27322F`).

## Deriving a brand from a document you already have

```bash
# 1. take the shell: fonts, theme, page size, header and footer bars
python -m docxforge.template "Their_Release.docx" --out brands/theirs.docx
#    it reports the header title and footer line it captured, and writes
#    brand.starter.json with those already filled in

# 2. read the design off the document
python -m docxforge.inspect_docx "Their_Release.docx"
#    the tail of the output lists every fill, text colour and font size by
#    frequency — that is your palette and type scale

# 3. build the same content and compare
python -m docxforge.build their_draft.json --brand brands/theirs.json \
       --template brands/theirs.docx --out Regenerated.docx
python -m docxforge.diff "Their_Release.docx" Regenerated.docx
```

Aim for 100% text similarity with zero runs on either side of the diff, then
work down the styling rows. Do not guess values — the point of step 2 is that
the numbers come from a document somebody already approved.

## Validation

Bad input is reported, never silently dropped:

```
brand warning: unknown palette key "inkk"
brand warning: palette.green is not a 6-digit hex colour: 'not-a-colour'
brand warning: geometry.content_width must be a positive integer: -5
brand warning: unknown type key "banner_titel"
brand warning: unknown string key "mastheadtitle"
```

A missing or unreadable brand file prints a note and builds on the defaults.
`configure()` resets to the defaults before applying a brand, so calling it
repeatedly in one process — a test run, a loop over several brands — gives the
same result as a fresh process each time.
