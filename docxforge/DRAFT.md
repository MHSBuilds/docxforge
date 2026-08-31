# The content JSON

The single input to `docxforge.build`. Write it, build, read the report, fix it,
build again — regeneration is always safe.

```jsonc
{
  "product": "Northwind",          // header bar and the second masthead
  "period":  "March 2027",         // both mastheads and the header bar
  "output":  "Release.docx",       // used when --out is not given

  "contents_blurb": "…",           // optional; overrides the brand's
  "intro": "…",                    // optional; overrides the brand's

  "updates": [                     // numbered sections, in document order
    {
      "title": "Stock Transfers — Move inventory between warehouses",
      "contents_title": "Stock Transfers",
      "audience": "Warehouse and Operations",
      "body": [
        {"kind": "para",   "text": "Moving stock between warehouses **was** a…"},
        {"kind": "bullet", "text": "**Rebalancing** — stock is being levelled…"},
        {"kind": "shot",   "slot": "update-1a.png",
                           "note": "Stock detail page with the Transfer action"}
      ]
    }
  ],

  "coming_soon": [                 // two-column card grid at the end
    {"title": "Cycle Counting", "text": "Schedule recurring partial counts…"}
  ],

  "gaps": [                        // shown in the report, not in the document
    "Update 3 does not record how ties are broken"
  ]
}
```

## Update fields

| Field | |
|---|---|
| `title` | The section banner. The full capability. |
| `contents_title` | The contents row — the **short** form, how someone would name it in conversation. Falls back to `title`. |
| `audience` | Contents column 3. Falls back to `[NEEDS INPUT]`. |
| `body` | The blocks below. An empty body renders a `[NEEDS INPUT]` paragraph rather than an empty section. |

Anything else you put on an update (a ticket id, a date) is ignored by the
builder and preserved in your file, which is useful for traceability.

## Body blocks

| `kind` | Renders as | Fields |
|---|---|---|
| `para` | Body paragraph, `after=120 line=264` | `text` |
| `bullet` | `•  ` + text, indented, tighter spacing | `text` |
| `shot` | The image, or a labelled placeholder | `slot`, `note`, `caption`, `width_in` |

The first `para` in a section carries `keepNext`, so a banner can never be
orphaned at the foot of a page from the text it introduces.

### Inline formatting

`**bold**` and `*italic*`, converted to real runs. Nothing else is interpreted —
write plain prose otherwise. Markers that do not pair are left as literal text.

### `shot` fields

| Field | |
|---|---|
| `slot` | The filename the image must have in the `--shots` folder. Convention: `update-<n><a\|b\|c>.png`, numbered by the section's position, so ordering never has to be tracked. |
| `note` | What the image should show. Appears in the shot list and inside the placeholder box. **Not** printed in the finished document. |
| `caption` | Printed under the image in italic grey. Usually leave unset. |
| `width_in` | Override fit-to-width, e.g. `4.2` for a narrow dialog that should not be blown up. Also allowed to enlarge. |

Extensions match loosely: a slot of `update-1a.png` is satisfied by
`update-1a.jpg` or `update-1a.jpeg` in the folder.

## Gaps

Anything the source did not supply goes two places, deliberately:

1. `[NEEDS INPUT]` inline, so a reviewer cannot miss it on the page.
2. A line in `gaps`, so it appears in the report even if the reviewer skims.

Write `[NEEDS INPUT]` into a `text` value wherever a number, threshold or option
list is unknown. The builder counts them and includes the total in the report.
Never fill a gap with a plausible sentence.
