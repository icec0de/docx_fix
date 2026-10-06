# docx_fix — MS Word check & fix

## Intent

- Make a `.docx` ready for **digital distribution**: clean, consistent, reflow-safe, metadata-free.
- Replace hand edits with deterministic, idempotent rules that can run again on every new revision.
- Leave semantic decisions (§10) to a human: report them, never change text silently.

The rules synthesize the working rules and implementation lessons from applying them to a
~150-page bilingual (RU/EN) document. `docx_fix.py` implements exactly what is listed here.

## Domain facts

- Real documents often use custom heading styles. Some resolve to a heading level through
  `basedOn` (e.g. a custom style based on `Heading1`); others resolve to no outline level at all
  and need an explicit `--heading-style ID=N`.
- Bold lead-ins are often bold at run level in a plain body style, not via a bold style.
- Source documents are not kept in this repo; test against your own files.

## Usage

```
uv sync
uv run docx_fix.py check <in.docx>                 # dry run: per-rule counts; exit 1 if anything to fix
uv run docx_fix.py fix <in.docx> [<out.docx>]      # default out: <in>_fixed.docx; input never overwritten
```

Options:
- `--heading-style ID=N` — treat paragraph style `ID` as heading level `N` when its `basedOn` chain
  has no `outlineLvl` (e.g. a custom subheading style: `--heading-style Subheading=3`).
- `--leadin ID` — accept a §10a review candidate (8-hex ID from the report) and convert its `.`→`:`.
- `--strip-rsid` — also run the optional §8.10 step.

Outputs: the fixed `.docx` and `<out>.report.md` — changes per rule per pass, what was changed
(lead-ins converted, tables newly given a repeating header), findings for review (§7a empty headings,
§10a undecided lead-ins), and the verification results. The file is written only when the passes
converge and every verification passes.

## Pipeline

1. `uv run docx_fix.py check <in.docx>` — see what would change; read the report.
2. Decide the §10a review candidates; pass accepted ones as `--leadin ID`.
3. `uv run docx_fix.py fix <in.docx> --heading-style … [--leadin …]`.
4. Open the result in Word once so the TOC and page fields refresh (§9), then save.

Observed (2026-10-06) on a ~150-page RU/EN document pair and its drafts with tracked changes and
comments: converges in 2 passes, reproduces a hand-made fix of the same rules, and the output
opens in LibreOffice as A4 with a rebuilt TOC.

---

## 0. Execution model & principles

A `.docx` is a ZIP of XML parts. Edit the XML directly; do not rely on round-tripping through a
word processor except where a layout engine is genuinely required (TOC page numbers).

**Idempotent + convergent multi-pass.** Every rule must be safe to run twice (a second run makes
no change). Because one fix can create the precondition for another (e.g. `.`→`:` on a lead-in then
enables the "keep-with-next" tie), the app should:

1. Run each rule; **count the changes it made**.
2. After a full pass over all rules, if any rule changed anything, **run another full pass**.
3. Stop when a whole pass produces **0 changes** — that final no-op pass proves convergence.

This removes any need to hand-tune a global rule order for two-way dependencies.

**One-way ordering that still matters.** A few steps have a hard directional dependency and must
run in a fixed sub-order regardless of the convergence loop:

- Document sanitize: **accept-revisions → strip-comments → metadata/whitespace**
  (accepting a tracked *deletion* can remove text a comment anchored to → avoids dangling anchors).
- Structural formatting (page/margins/tables/lists) before text-dependent detection is convenient
  but not required if the loop converges.

**Verify-and-revert for destructive structural edits.** Any removal that could break references
(unused styles / numbering) must be followed by a reference-integrity check, with automatic revert
of that step if a referenced item would break.

**Cross-cutting implementation invariants** (see Appendix for detail):
- Never truth-test an lxml element — always `is None`. An empty element is falsy and silently
  breaks `elem = find(tag) or SubElement(...)`.
- Insert child elements in **schema (CT_*) order**, or Word rejects the file.
- Text edits that may span multiple runs must use an **offset map over concatenated run text**,
  not per-run regex.

---

## 1. Page size & margins

Digital distribution → symmetric margins, no binding gutter.

**Fix — for every `w:sectPr`:**
- Page size **A4**: `w:pgSz w:w="11906" w:h="16838" w:code="9"`, remove `w:orient` (portrait).
- Margins **2.0 cm** all sides: `w:pgMar` top/bottom/left/right = `1134` twips.
- Header & footer **1.25 cm**: `w:header`/`w:footer` = `709`; `w:gutter="0"`.
  (twips = cm × 566.929; 1440 twips/inch.)

**Mechanics:** remove any existing `pgSz`/`pgMar` in the section and insert single fresh ones in
`CT_SectPr` order (`…type, pgSz, pgMar, paperSrc…`). Do **not** use `find() or SubElement()` — it
duplicates the element (see Appendix A).

**Check:** exactly one `pgSz` (11906×16838) and one `pgMar` (all 1134/709) per section.

---

## 2. Table width — full page span

**Fix — every `w:tbl`:**
- `w:tblW w:type="pct" w:w="5000"` (100% of text area).
- `w:tblLayout w:type="autofit"` (columns scale to the page; replaces fixed layout).
- Reset any non-zero `w:tblInd` to 0 so the table starts at the left margin.

Equivalent to Word's **AutoFit → AutoFit Window**. Insert into `w:tblPr` in `CT_TblPr` order
(`…tblW, jc, tblCellSpacing, tblInd, tblBorders, shd, tblLayout, tblCellMar, tblLook…`).
Stored `gridCol`/`tcW` widths are left in place; Word rescales them on open, columns keep their
relative proportions while the table fills the width.

**Optional (deterministic alternative):** rewrite columns as explicit percentages summing to 100%
if you want Word-independent widths instead of relying on its recalc.

---

## 3. Page breaks — tables

**For every `w:tbl`:**
- **First row marked as header** → `w:tblPr/w:tblLook w:firstRow="1"`.
- **Header repeats on each page** → add `w:tblHeader` to the first row's `w:trPr`.
- **No row splits across pages** → add `w:cantSplit` to **every** row's `w:trPr`.

`w:trPr` order: `cantSplit` precedes `trHeight` precedes `tblHeader`.

**Caveat to surface:** applying header-repeat to *all* tables includes any that are intentionally
headerless. Either apply to all (document-wide consistency) or let the user exclude specific tables.

---

## 4. Page breaks — lists & list lead-ins

**List item** = paragraph with `w:pPr/w:numPr`. A **list** = a maximal run of consecutive sibling
list-item paragraphs.

**For every list item:**
- `w:keepLines` — keep the item's own lines together (no mid-item page split).
- `w:keepNext` — on every item **except the last of its run** (glues the list together without
  gluing it to the paragraph after the list).

**List lead-in** = a non-list paragraph **immediately followed by a list item**. Give it
`w:keepNext` + `w:keepLines` so the intro line can't be orphaned from its list. **Gate on
"immediately precedes a list", not on ending punctuation** — the colon is cosmetic; the orphan risk
applies to any lead-in, including heading-style labels. (See §10 for the semantic caveat on `.`-ending
candidates that might be summaries of the *previous* block.)

`w:pPr` order: `pStyle, keepNext, keepLines, …, numPr, …`.

---

## 5. Lead-in punctuation & ties (deterministic parts)

Two lead-in shapes get a colon and a keep-with-next tie.

### 5a. Colon lead-in before a list
A non-list paragraph ending with `:` immediately before a list → already correct punctuation; ensure
it is tied (`keepNext` + `keepLines`). Covered by §4.

### 5b. Bold sentence after a heading (introduces a passage, not a list)
**Rule — a paragraph is a lead-in when ALL hold:**
- it is a **single sentence** (no internal `. `/`! `/`? ` + word-start);
- it is **bold** (see effective-bold below);
- it comes **immediately after a heading level 1–3**;
- it **ends with `.`**;
- it is **immediately followed by a non-empty paragraph or a table** (a table is a passage too;
  e.g. a bold "We distinguish three factors." directly above a table).

**Then:** replace the trailing `.` with `:`, and tie with the next passage (`keepNext`+`keepLines`).

- **Effective-bold** resolution (any one true ⇒ bold): run `w:rPr/w:b` → run style (`rStyle`) bold →
  paragraph style (`pStyle`) bold, walking `basedOn` inheritance. A style-only or run-only check
  misses the other source.
- **Heading level 1–3** is not just stock `Heading1..3`: resolve custom styles via their `basedOn`
  chain and/or `w:outlineLvl` 0–2, plus `--heading-style` overrides for styles that resolve to
  neither.

### Shared colon guard (both 5b and §10)
Skip the `.`→`:` swap when the **final sentence** already contains a `:` (avoid a double colon);
still apply the tie. Check the *final sentence only* — a colon in an earlier sentence of a
multi-sentence lead-in is fine to convert.

---

## 6. Text sanitization

Apply per paragraph; edits that cross runs use the offset-map method (Appendix C).

- **Range dashes between numbers** (incl. currencies and units): normalize `X - X`, `X- X`, `X -X`
  → tight dash `X–X`. Use an **en-dash** (matches the standard numeric-range convention:
  `1–7 days`, `30–50%`, `L0–L5`). Fire **only** when a numeric/currency/%/unit token sits on both
  sides **and** there is ≥1 surrounding space. **Never** touch: hyphenated words (`AI-agent`,
  `backend-engineer`), em-dash punctuation (`—`), or a negative-number prefix (`−25%`). Already-tight
  ranges are left as-is.
- **Trim** leading and trailing spaces/tabs per paragraph (keep NBSP).
- **Collapse** runs of 2+ regular spaces to one (`[ ]{2,}` → single). Leave the space-thousands
  separator (single space) untouched.

---

## 7. Structural cleanup

### 7a. Empty headings
**Check:** heading-styled paragraphs (level 1–3, resolved as in §5b) whose text is empty/whitespace.
**Fix:** report; remove only if the paragraph is genuinely empty **and** not acting as deliberate
spacing before real content. Prefer flag-for-review over silent deletion (an empty heading can carry
a bookmark/anchor target — check for `bookmarkStart` / TOC `_Toc` refs before removing).

### 7b. Excessive blank lines
**Check:** runs of consecutive empty paragraphs (no text, no drawing/field).
**Fix:** collapse any run of **≥2** empty paragraphs down to **1** (keep a single intentional
spacer; never zero them out — that changes section spacing). Do not remove an empty paragraph that
holds a `w:sectPr` (section break) or a bookmark anchor.

---

## 8. Document sanitization

Run in the fixed sub-order **accept-revisions → strip-comments → metadata/whitespace**.
Skip no-op steps silently; report counts.

1. **Accept all tracked changes.** Across all story parts (document, headers, footers, foot/endnotes):
   unwrap `w:ins`/`w:moveTo` (+ range markers) to keep content; delete `w:del`/`w:moveFrom` (+ range
   markers) to drop deleted content; delete formatting-change records `w:pPrChange`/`rPrChange`/
   `tblPrChange`/`trPrChange`/`tcPrChange`/`sectPrChange`/`tblGridChange`/`numberingChange`/
   `cellIns`/`cellDel`/`cellMerge`. Remove `w:trackChanges` from `settings.xml`.
2. **Remove comments (whole family).** In-body anchors (`commentRangeStart`/`End`,
   `commentReference` runs) **and** the parts `comments.xml`, `commentsExtended.xml`,
   `commentsIds.xml`, `commentsExtensible.xml`, `people.xml`, their Content-Types Overrides, and
   their relationships.
3. **Strip metadata / PII.** `core.xml`: clear `dc:creator`, `cp:lastModifiedBy`, remove
   `cp:revision`. `app.xml`: clear `Company`/`Manager`/`Template`, `TotalTime`→0.
4. **Remove macros & embeds.** Delete `word/vbaProject.bin` (+ `vbaData.xml`, forces `.docx` not
   `.docm`) and OLE objects under `word/embeddings/*` (+ their refs/rels/content-types).
5. **Remove hidden/vanished text.** Delete runs whose `rPr` carries `w:vanish`/`w:specVanish`.
6. **Strip proofing marks.** Remove `w:proofErr`; remove `w:noProof`; remove per-run `w:lang`
   overrides (keep the document default lang in `styles.xml` docDefaults / `settings.xml`
   themeFontLang). Content-neutral; reduces noise.
7. **Remove document/write protection.** Delete `w:documentProtection` / `w:writeProtection` from
   `settings.xml`.
8. **Clean orphans (verify + auto-revert gate):**
   - **Numbering** — remove `w:num` whose `numId` is unreferenced and `w:abstractNum` not reached
     by a kept num. Keep-set follows `num`↔`abstractNum`↔`styleLink`/`numStyleLink` and numbering
     carried by referenced styles. **Verify** every referenced `numId` still resolves num→abstractNum;
     revert the step on failure.
   - **Styles** — remove `w:style` not referenced anywhere. Keep-set expands over
     `basedOn`/`link`/`next` and always keeps `w:default="1"` styles. **Verify** every referenced
     `pStyle`/`rStyle`/`tblStyle` is still present; revert on failure. (Structural defaults —
     `Normal`, `DefaultParagraphFont`, `TableNormal`, `NoList` — are `default="1"` and thus kept.)
   - **Dangling relationships** — remove rels whose local `Target` file is missing (skip `External`).
   - **Unreferenced media** — delete `word/media/*` not named by any rel Target.
   - Gather references from **every** part (settings, numbering, headers/footers), not just
     `document.xml`.
9. **Remove temp bookmarks.** `bookmarkStart`/`End` named `_GoBack` or `OLE_LINK*` (keep `_Toc*`).
10. *(Optional, out of default scope)* strip `w:rsid*` editing-fingerprint IDs from `settings.xml`
    and attributes; add `w:removePersonalInformation` to instruct Word to keep PII out on future saves.

---

## 9. Update the Table of Contents

TOC page numbers require a **pagination engine** — they cannot be computed from XML. Do not fabricate
them.

**Fix (defer to the layout engine that opens the file):**
- `settings.xml`: add `w:updateFields w:val="true"` (schema position: just before
  `hdrShapeDefaults`/`compat`/`rsids`) → Word/LibreOffice refresh **all** fields on next open.
- Mark every field-begin `w:fldChar w:dirty="true"` in `document.xml` and headers/footers (the TOC
  field, each entry's PAGEREF page number, footer page fields).

**Result:** on open in Word the TOC rebuilds text + page numbers (Word may show a one-line update
prompt). Manual equivalent: right-click TOC → Update Field → *Update entire table* (or Ctrl+A, F9).

**To materialize numbers headlessly** (no Word), drive LibreOffice `soffice --headless` with a macro
that refreshes indexes / updates all fields, then save as docx (LibreOffice pagination is close to,
but not identical to, Word's).

---

## 10. Non-deterministic / semantic cases

### 10a. Lead-in that ends with `.` but should end with `:`
A `.`-ending paragraph immediately before a list (or passage) is **either** a genuine lead-in
(introduces the next block → wants `:`) **or** a standalone/summary sentence of the *previous* block
that merely precedes the next one (leave as `.`). This is **not** deterministic — it needs a decision
about which block the sentence is semantically bound to (**next**, not previous).

**Process:**
1. Find every non-list paragraph immediately before a list, ending in `.` (not `:`).
2. Decide lead-in vs summary (semantic).
3. For genuine lead-ins: convert final `.`→`:` (with the §5 final-sentence colon guard), then apply
   the keep-with-next tie (§4). Re-run the tie after conversion so newly-colon'd lead-ins are covered.

**Deterministic signals that make most cases easy (apply before any ML):**
- **Count agreement** — the lead-in names a count that equals the number of following items
  (“four fields” → 4 items, “three metrics” → 3). Strong lead-in signal. The count may sit in an
  earlier sentence of the paragraph ("Three anti-patterns recur. Each has its own fix."), so it is
  searched in the whole paragraph; the colon guard still checks only the final sentence.
- **Enumerative phrasing** — "consists of / comprises / includes / we distinguish N / through N …".
- **Exclusions** — a sentence that reads as a conclusion/summary, or whose final sentence already
  holds a `:`.

Surface residual uncertain cases for human review rather than auto-editing.

### 10b. Can embeddings solve it?
Partly — as one signal, not the whole answer. Embedding the candidate sentence and the adjacent
blocks and comparing cosine similarity to the **next** block vs the **previous** block gives a
directional "which way does it bind" score; a lead-in should sit closer to what follows. But:
- It is **probabilistic** and needs a tuned threshold; the failure mode (wrongly flipping a summary's
  `.`→`:`) is a visible text error, so precision matters more than recall.
- The cheap deterministic signals above already resolve the overwhelming majority with **higher
  precision** and zero model cost.

**Recommended design:** deterministic signals first (count-match, enumerative verb, colon guard);
for the ambiguous remainder, use an LLM classifier or embedding-similarity with a conservative
threshold, and **route low-confidence cases to human review** (or mark them as Word comments) instead
of editing. Never let the semantic layer make a silent, unreviewable text change.

---

## Appendix — implementation gotchas

**A. lxml element truth-testing.** `find(tag) or SubElement(parent, tag)` is a trap: an element with
no children evaluates **falsy**, so the `or` fires even when the element exists, creating a duplicate
(observed: duplicate `pgSz`/`pgMar`, with the stale one first — which Word then uses). Always test
`is None`.

**B. Schema ordering.** Children of `w:pPr`, `w:rPr`, `w:trPr`, `w:tblPr`, `w:sectPr`, `w:settings`
must appear in the ECMA-376 `CT_*` sequence order. Insert new elements at the correct slot (find the
first existing child whose order-index is greater), not by append.

**C. Cross-run text edits.** A single logical string (a dash range, a double space, a trailing space)
can be split across several `w:t` runs. Edit via: concatenate run texts → find match offsets on the
full string → apply edits right-to-left, mapping each offset back to its run (offsets left of an edit
are unaffected, so original coordinates stay valid). Avoids run-splitting and preserves formatting.

**D. Comment id invariant.** After any comment edit, the id sets of `commentRangeStart` ==
`commentRangeEnd` == `commentReference` == `comments.xml` `w:comment` must be equal.

**E. Repackaging.** Store `[Content_Types].xml` first, deflate the rest, no directory entries; keep
the exact part set (minus only the parts a rule deliberately removes). Validate: every XML part parses,
ZIP integrity OK, and the reference-integrity checks in §8.8 pass.

**F. Verify each rule end-to-end**, not just that the edit was written — confirm the intended property
holds on re-read (e.g., every table row has `cantSplit`; every list lead-in has `keepNext`; every
referenced style/numId resolves).

---

## Repo

- Python ≥ 3.12, uv (`uv sync`, `uv run`), single dependency `lxml`. No inline script metadata.
- `docx_fix.py` is the only entry point; it reads one `.docx` and writes next to it (or to the
  given path) through `<out>.part` + `os.replace`.
- `README.md` is the user-facing summary; `spec.md` stays authoritative.
- Gitignored: `.venv/`, `__pycache__/`, `.claude/`, `.DS_Store`, Word lock files `~$*`, `*.part`,
  and generated `*_fixed.docx` / `*.report.md`.
