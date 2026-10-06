# docx_fix

Check and fix a Microsoft Word `.docx` for **digital distribution**: clean,
consistent, safe to reflow, and free of metadata.

It edits the document's XML directly, with no Word round trip. It applies a set
of deterministic rules and repeats full passes until a pass changes nothing, so
running it twice gives the same result. Anything that needs a human decision is
reported, never changed silently.

## What it fixes

| Area | Rule |
| --- | --- |
| Page setup | A4 portrait, 2.0 cm margins on all sides, 1.25 cm header/footer, no gutter |
| Tables | Full text width (AutoFit Window); first row is a header repeated on every page; rows never split across pages |
| Lists | Each item keeps its lines together; the whole list stays together; the paragraph introducing a list stays with it |
| Lead-ins | `.` → `:` on a sentence that introduces a list or passage, when deterministic signals agree (see below) |
| Text | Spaced number ranges → en dash (`$100K - $1M` → `$100K–$1M`); leading/trailing spaces trimmed; double spaces collapsed |
| Structure | Runs of 2+ empty paragraphs collapsed to one; empty headings reported |
| Sanitize | Accept all tracked changes; remove all comments (with their parts); strip author/company metadata; remove macros, OLE embeds, hidden text, proofing marks, protection, `_GoBack` bookmarks; drop unused styles, numbering, relationships and media |
| TOC | Mark all fields dirty and set `updateFields`, so Word rebuilds the TOC and page numbers on open |

Lead-in punctuation is the only text-meaning rule. A `.`-ending sentence before
a list is changed to `:` only when the sentence names the number of items that
follow ("covers four areas." + 4 items), or uses enumerating phrasing
("consists of", "includes", …). A colon already in the final sentence blocks the
change. Everything else goes to the report as a review item.

## Requirements

- [uv](https://docs.astral.sh/uv/) (Python 3.12+, dependencies locked in `uv.lock`; the only one is `lxml`).

```bash
uv sync
```

## Usage

```bash
uv run docx_fix.py check report.docx          # dry run: what would change; exit 1 if anything
uv run docx_fix.py fix report.docx            # writes report_fixed.docx (the input is never modified)
uv run docx_fix.py fix report.docx out.docx   # choose the output path
```

Options:

- `--heading-style ID=N` — treat paragraph style `ID` as heading level `N`. Needed
  for custom heading styles that don't resolve to an outline level through their
  `basedOn` chain (repeatable).
- `--leadin ID` — accept a lead-in from the report's review list and change its
  final `.` to `:` (repeatable).
- `--strip-rsid` — also remove rsid editing fingerprints and tell Word to keep
  personal information out of future saves.

Each run also writes `<output>.report.md` (or `<input>_check.report.md` for
`check`). It contains:
- the number of changes per rule in each pass
- what was changed (converted lead-ins, tables newly given a repeating header)
- items for review (undecided lead-ins with their IDs, empty headings)
- the verification results

### Typical workflow

1. `check` the document and read the report.
2. Decide the lead-ins under review; accept the right ones with `--leadin ID`.
3. `fix` the document.
4. Open the result in Word once so the TOC and page numbers refresh, then save.

## Safety

- The input file is never overwritten; output is written through a temporary
  `.part` file.
- The output is written only if the passes converge and all checks pass:
  - every XML part parses
  - child elements are in the order the Word schema requires
  - every referenced style and numbering definition resolves
  - every relationship target exists
  - content types match the package
  - comment anchors are consistent
- Removing unused styles or numbering is checked straight away and undone if it
  would break a reference.

## Limitations

- TOC page numbers need a layout engine, so they are not computed. Word (or
  LibreOffice) rebuilds them on open.
- Header repeat is applied to every table. Check the report's list of tables that
  newly got a header row; a table that should have none needs manual fixing.
- Lead-in detection is tuned for Russian and English wording.

## Specification

[`spec.md`](spec.md) is the source of truth: every rule, its exact OOXML
mechanics, and the implementation pitfalls. `docx_fix.py` implements exactly
what it lists.
