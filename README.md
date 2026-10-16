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
| Styles | Unstyled paragraphs and headings typed as body text get their styles; captions and block titles get their own styles; dominant direct formatting is promoted into styles and the rest stripped (fonts, sizes, colours, spacing, indents); emphasis is kept |
| House style | Body justified, 6 pt before/after, auto hyphenation; theme fonts; black text and headings (colours only from a `--scheme`); equal, unindented, full-width TOC levels with 18 pt before chapters; embedded fonts removed |
| Captions | A picture and its caption, a caption and the table or picture under it, stay on one page |
| Table cells | 0.1" padding on all sides, text top-left, header rows without their own spacing |
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
- `--only N[,N...]` — run only the rules with these spec section numbers, e.g.
  `--only 8.1,8.2` accepts tracked changes and removes comments without touching
  layout or text. All verification checks still run.
- `--font NAME` — set the theme heading and body fonts to `NAME`, so the whole
  document uses it (with rule 11.6). The font is not embedded in the `.docx`; a PDF
  export embeds it.
- `--toc-levels N` — the table of contents lists heading levels 1–N; Word drops
  the deeper entries when it updates the TOC.
- `--scheme FILE.toml` — colour scheme applied after the house style, which makes
  headings black (with rule 11.6): text colours of styles and the fill of shaded table
  header rows. Keep a document's scheme next to the document; see
  [`schemes/example.toml`](schemes/example.toml) for the format.

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

### Naming convention

Working files and release files are kept apart:

- **Working files** stay next to the source document. A new iteration gets the
  next `_v##` (two digits) at the end of the name: `Report_v01.docx`,
  `Report_v02.docx`, … Earlier versions are never overwritten.
- **Release files** go to a `release/` folder and are named
  `<name>_<version>_rc##.docx`. Here `<version>` is the public version number and
  `##` is the release-candidate iteration as two digits, so the files sort
  correctly past 9: `release/Report_1.0_rc01.docx`, `release/Report_1.0_rc02.docx`,
  … Release files are not edited; each fix produces the next `rc##`.
- A document's colour scheme (`--scheme`) is kept next to the source document as
  `<name>.scheme.toml`.

## Rendering with Microsoft Word (macOS)

`docx_fix` does not paginate. To see the result and rebuild the TOC and page
numbers, let Word open the file and save a PDF:

```bash
osascript word_pdf.applescript "$PWD/report_fixed.docx" "$PWD/report_fixed.pdf"
```

The script opens the document, updates every table of contents, saves the PDF and
closes that document without saving it. Other open documents are left alone. Both
paths must be absolute.

Rule 9 (TOC refresh) sets `updateFields`, which makes Word ask on open whether to
update fields. That question blocks a scripted run, so render a file fixed without
rule 9 (`--only` without `9`). The script updates the tables of contents itself.

### File access: grant your home folder once

Word for Mac is sandboxed. It opens a file outside its own container only from a
folder the user has granted. For a new folder it shows a *Grant File Access*
dialog that offers exactly the folder of the file being opened. You can't pick a
parent folder there. If the file is in a hidden folder such as `/tmp`, the dialog
names a folder that Finder doesn't show.

To be asked once for your whole home folder, run:

```bash
sh word_grant_access.sh          # or: sh word_grant_access.sh /path/to/folder
```

It puts a throwaway document in your home folder, opens and closes it in Word,
and deletes it. When the dialog appears, click **Grant Access**. Word keeps the
grant for that folder and everything below it, so later runs in any subfolder
don't ask again.

Keep inputs and outputs under the granted folder, not in `/tmp`. The repo's
`render/` folder (gitignored) is there for that. The first `osascript` run also
asks once whether your terminal may control Microsoft Word (System Settings →
Privacy & Security → Automation).

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
