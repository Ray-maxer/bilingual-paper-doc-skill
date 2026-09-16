---
name: bilingual-paper-doc
description: Turn an academic paper PDF into a polished bilingual (English/Traditional Chinese) Word reading document — alternating EN paragraph / ZH translation, embedded figures with bilingual captions, and real editable Word equation objects. Use this whenever the user uploads a research paper (or two related PDFs, e.g. an original + their own partial translation) and asks for a "中英對照" reading doc, a bilingual translation of a paper as a Word file, or to redo/extend a bilingual paper doc they made before. Also use for requests like "整理成中英對照" or "把這篇論文的圖表跟中文翻譯做成 Word 檔". Don't use this for short passages, non-academic text, or when the user just wants a plain translation with no document/figures — this skill is specifically for the "full paper + figures + equations + Word deliverable" workflow.
---

# Bilingual paper reading document

Produces a cover-to-references .docx: title page, alternating EN/ZH paragraph
pairs, figures pulled from the source PDF with bilingual captions, numbered
equations as real editable Word equation objects, yellow "reading note"
callouts, and an English-only references list.

**Division of labor**: the mechanical parts (figure extraction, docx
assembly, XML post-processing, schema validation) are scripted in
`scripts/`. The parts that require actually reading and understanding the
paper — translating each paragraph, deciding what's a figure vs. page
furniture, transcribing equations — are NOT scriptable and are done by you
(Claude) directly, paragraph by paragraph, the same way you'd read the paper
normally. This skill turns that reading/translating pass into a highly
polished deliverable with minimal extra effort, not a magic PDF-to-docx
button.

## Translation fidelity (academic integrity)

This produces a reading document for the user's own academic work, not a
new piece of writing. Every EN/ZH pair must be a faithful translation of
what the source paper actually says — never author, add, infer, embellish,
or "improve" on content that isn't in the original. This matters more than
polish: a rougher but accurate translation is always the right trade-off
over a smooth but invented one.

- **Translate, don't compose.** The Chinese text is a translation of the
  English paragraph next to it, not a paraphrase, summary, or your own
  explanation of the topic. Match the source's claims, numbers, citations,
  and scope exactly — don't round numbers, generalize a specific claim, or
  narrow/broaden what the source says.
- **Never fabricate.** Don't invent data, results, citation details,
  example values, or filler transitions that aren't in the PDF, even to
  smooth a rough patch. If a number, unit, symbol, or citation is genuinely
  unclear in the source, transcribe it as closely as you can read it rather
  than substituting a plausible-looking value.
- **When genuinely ambiguous, ask — don't guess.** If a term or passage
  could be read multiple ways depending on the paper's subfield or the
  user's own usage, ask the user directly (this is what Step 0's
  focus/terminology question is for) rather than picking a plausible-
  sounding interpretation and presenting it as settled. Reserve an inline
  flag like `[原文此處不清楚]` for cases where asking doesn't help — e.g. a
  genuinely illegible scan — not as a substitute for asking when the user's
  answer would actually resolve it.
- **Don't silently drop or condense content.** Every paragraph, footnote,
  and equation in scope gets its own `pair`/`figure`/`equation` block.
- **Reading notes stay clearly marked as notes.** Any commentary that is
  yours rather than the paper's belongs in `note` (yellow callout) blocks
  only — never blended into a `pair` translation, where it would look like
  part of the paper itself.
- This applies equally when merging with the user's own draft translation
  (Step 0): extend their translation faithfully for the gaps, don't
  editorialize on top of what they already wrote.

## Workflow

### Step 0 — gather inputs
If the user has already written a partial translation (as in the case this
skill was built from: an original English PDF + the user's own manually
highlighted EN/ZH notes PDF), read both. Treat the user's own translation as
the authoritative source for wording/tone where it exists, and extend it
(translate) for any sections it doesn't cover — don't silently drop content
the original has that the user's draft skipped. The fidelity rules above
apply to both the user's existing text and whatever you add.

Before starting extraction/translation, ask the user what they're focusing
on: which part of the paper matters most to them (methods vs. results vs. a
specific model/section), their own field/background, and whether any
recurring technical terms have a Chinese rendering they already use. A
paper's terminology often carries different meanings across subfields, and
translating an ambiguous term on your best guess risks quietly misleading
the user in exactly the part they care about most — a short question here
is cheaper than a wrong translation later.

### Step 1 — extract figures
```bash
bash scripts/extract_figures.sh <paper.pdf> ./imgs
```
This dumps every embedded raster image as `fig-<page>-<idx>.png`, including
noise (journal logos, small icons, cover thumbnails). **View each one** to
identify the real figures and their correct number/order, then rename the
real ones to `Fig1.png`, `Fig2.png`, etc.:
```bash
cp imgs/fig-002-004.png imgs/Fig1.png
```
Check each renamed figure's pixel dimensions to plan `w`/`h` in the content
data (see sizing note in `references/data-schema.md`):
```bash
python3 -c "from PIL import Image; print(Image.open('imgs/Fig1.png').size)"
```

### Step 2 — write content-data.js
Read `references/data-schema.md` for the full field reference, and
`references/example-content.js` for a worked pattern to copy. Go through the
paper section by section and build the `sections` array: `h1`/`h2` headings,
`pair` blocks (EN paragraph + ZH translation), `figure` blocks at the point
each figure is first referenced, `table` blocks for any data/results tables
(transcribe cell values exactly as printed — see "Translation fidelity"
above), `equation` blocks for every numbered equation (using the
`SUB`/`SUP`/`SUBSUP`/`FRAC`/`INT` token builders — see "Equation tokens" in
the schema doc), and `note` blocks for any callouts the user's own draft
already had.

Translate paragraph-by-paragraph as you go — don't summarize, skip, or
invent content (see "Translation fidelity" above). Every equation in the
paper should become its own `equation` block so it renders as a real,
editable Word equation (not plain text).

**Inline scientific symbols**: `lib.js` already auto-detects common inline
variable patterns in prose (`Mmax`, `fmax`, `CD`, `CM`, `θmax`, `σmax`,
`Hpull`, `FD`, `FM`, superscript-tagged fits like `MLmax`/`MQmax`, `DBH²`,
`H²`, `a²`, `R²`, ...) and renders them with real Word subscript/superscript
formatting automatically — you don't need to do anything special in the
prose text, just type the variable name normally (e.g. `"the moment Mmax
increases with DBH"`) and it will come out with a proper subscript. If this
paper uses variable names not in that default list, add `extraInlineRules`
to your content-data.js (see schema doc) rather than editing `lib.js`.

### Step 2.5 — fidelity self-check
Before rendering, skim back through content-data.js against the source PDF:
confirm every in-scope paragraph, footnote, and equation has a matching
block, and spot-check a handful of ZH translations against their EN pair
for anything added, dropped, or softened. Catch summarization creep or
invented filler here, before it's baked into the .docx.

### Step 3 — render
```bash
node scripts/render.js content-data.js output.docx
```

### Step 4 — post-process (REQUIRED, do not skip)
```bash
python3 scripts/inject_mathpr.py output.docx output_final.docx
```
This adds the `<m:mathPr>` settings block Word expects once a document
contains equation objects. Skipping it still produces a technically valid
file, but do it anyway — it's one command.

### Step 5 — validate
```bash
python /mnt/skills/public/docx/scripts/office/validate.py output_final.docx --original output.docx
```
Must say "All validations PASSED!" before you present the file.

### Step 6 — visual sanity check
Render to images the normal way (see the docx skill) and view a handful of
pages — cover, the TOC page, a figure page, a table page (check the
bilingual header row and that captions sit above, not below), a page with
prose (to check inline subscripts), and a page with a display equation.

**Known limitation**: the equation *display* (the numbered `m:oMath` blocks
from Step 2/3, e.g. `M_max = f_max × H_pull`) will render as a **blank
paragraph** when previewed via `soffice --headless --convert-to pdf` in this
sandbox — LibreOffice's headless PDF export doesn't render OOXML equation
objects, even though they are completely valid. This is NOT a bug in the
document. Do not try to "fix" it by reverting to plain-text equations.
Instead, verify equations structurally:
```bash
pandoc -t markdown output_final.docx | grep '\$'
```
Each equation should round-trip into a readable LaTeX-ish line (e.g.
`$M_{\max}^{L}(DBH)\  = \ 166.19 \cdot DBH - 5.18$`). If it doesn't, the
token structure in content-data.js has a bug — fix there, not in lib.js.
Inline subscripts (plain TextRun-based, not m:oMath) DO render correctly in
the LibreOffice preview, so those are fine to eyeball directly.

**Same limitation applies to the Table of Contents**: the TOC page (right
after the cover) will also preview blank under headless LibreOffice, for
the identical reason — it's a Word field (`TOC \h \o "1-2"`) with no cached
text, and headless conversion doesn't evaluate fields. This is not a bug:
the document sets `updateFields` so real Word populates it automatically
the first time the file is opened. Verify structurally instead by checking
the field code and heading styles landed correctly:
```bash
python3 -c "
import zipfile, re
xml = zipfile.ZipFile('output_final.docx').read('word/document.xml').decode('utf-8')
m = re.search(r'<w:instrText[^>]*>([^<]*TOC[^<]*)</w:instrText>', xml)
print('TOC field code:', m.group(1))
print('Heading1/2 style uses:', len(re.findall(r'w:pStyle w:val=\"Heading[12]\"', xml)))
"
```

### Step 7 — deliver
Copy to `/mnt/user-data/outputs/` and use `present_files`. Mention in your
reply (briefly, once) that display equations are real editable Word
equation objects and inline symbols use native super/subscript formatting —
the user doesn't need this explained every time, just the first time or if
they ask.

## Files in this skill

- `scripts/lib.js` — the generic rendering engine (equation DSL, inline
  subscript detector, paragraph/heading/figure/note helpers, document
  assembly). Content-agnostic; should not need paper-specific edits.
- `scripts/render.js` — CLI: `node render.js content-data.js output.docx`
- `scripts/inject_mathpr.py` — required post-processing step, see Step 4
- `scripts/extract_figures.sh` — figure extraction helper, see Step 1
- `references/data-schema.md` — full content-data.js field reference +
  equation token guide
- `references/example-content.js` — trimmed worked example to copy the
  shape of

## Design notes (why things are built this way)

- **Equations are real Word equation objects (OOXML `m:oMath`), not
  Unicode/text approximations.** An earlier version of this skill used
  Unicode pseudo-superscript characters (e.g. `ᴸ`, `ᴹ`) directly in text —
  these are inconsistent across fonts and looked broken. A later version
  used real Word `subScript`/`superScript` run formatting for everything —
  correct and robust, but not truly "editable equations" in Word's sense
  (no equation editor, no fraction bars). The current version uses actual
  `m:oMath` objects for numbered display equations (matching what Word's
  Equation tool produces) and reserves the text-run-based
  subscript/superscript approach for *inline* symbols inside prose
  paragraphs (e.g. "the moment `Mmax` increases..."), where a full equation
  object would be overkill and disrupt text flow.
- **LibreOffice headless preview renders `m:oMath` blocks as blank** even
  though the XML is valid — confirmed via XSD schema validation
  (`validate.py`) and a pandoc round-trip (`pandoc -t markdown` correctly
  reconstructs LaTeX from the OOXML). Don't be misled by a blank preview
  into thinking the equations are broken; verify with pandoc instead.
- **`<m:mathPr>` needs the `m:` namespace prefix, not `w:`.** Using
  `<w:mathPr>` instead of `<m:mathPr>` produces a schema validation error
  that looks unrelated ("element not expected, expected docVars...") —
  this cost real debugging time, hence the dedicated `inject_mathpr.py`
  script with the correct namespace baked in.
- **The Table of Contents is a real Word field (`TOC \h \o "1-2"`), not a
  manually-typed heading list**, generated for free off the same real
  `HeadingLevel.HEADING_1/2` styles the headings already used. It has no
  cached entry text, so it's blank until a field-evaluating renderer
  computes it — the document sets `updateFields` so Word does this
  automatically on open. Don't be misled by a blank TOC page in a headless
  LibreOffice preview into thinking it's broken (same root cause as the
  equation-preview limitation above); verify the field code and heading
  styles structurally instead (see Step 6).
- **Table captions render above the table, figure captions render below
  theirs** — this matches standard academic convention (tables are
  captioned above, figures below) rather than using one placement for both.
- **Document metadata (title/creator/subject/description) is populated
  from `meta`**, not left blank. This is what shows up in Word's Info panel
  and what reference managers like Zotero read when someone imports the
  file, so it's worth getting `meta.titleEn`/`authors`/`citation` right
  even though they're "just" cover-page fields.
