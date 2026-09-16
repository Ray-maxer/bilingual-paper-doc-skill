# bilingual-paper-doc

A [Claude Skill](https://docs.claude.com) that turns an academic paper PDF
into a polished bilingual (English / Traditional Chinese) Word reading
document — cover page, native clickable table of contents, alternating
EN/ZH paragraph pairs, figures with bilingual captions, data tables, and
numbered equations as real, double-clickable Word equation objects (not
plain text or images).

## What it produces

- Title page with citation/DOI, pulled straight into the Word file's own
  metadata (title/author/description) so it plays nicely with reference
  managers like Zotero/EndNote
- A native, clickable table of contents (built from real Word heading
  styles, not a typed-out list)
- Alternating English paragraph / Chinese translation pairs
- Figures extracted from the source PDF with bilingual captions
- Bilingual data tables for results/comparison tables
- Numbered equations as real OOXML equation objects, editable in Word's
  own Equation Editor — plus inline scientific symbols (e.g. `Mmax`) with
  real subscript/superscript formatting in running prose
- Yellow "reading note" callouts for anything that's your own commentary,
  kept clearly separate from the translated text

## Translation fidelity

This skill is built for genuine academic use: translations are meant to be
faithful renditions of the source paper, never invented, embellished, or
summarized. See the "Translation fidelity (academic integrity)" section in
`SKILL.md` for the full rules it follows — no fabricated data or citations,
ambiguous passages get asked about rather than guessed at, and there's a
self-check step before final rendering.

## Requirements

This is a **Claude Skill**, not a standalone script — the actual reading
and translating is done by Claude, guided by `SKILL.md`. To use it you
need:

- A Claude surface that supports Skills (Claude.ai, Claude Code, or Cowork)
  with this skill added
- Anthropic's built-in `docx` skill available alongside this one (used for
  validation in Step 5)
- In the sandbox/container Claude runs in: Node.js with the [`docx`](https://www.npmjs.com/package/docx)
  npm package, Python 3 with Pillow, and `pandoc` (LibreOffice is optional,
  used only for a visual sanity-check render)

## Usage

Upload a research paper PDF (optionally alongside your own partial
translation) and ask for a "中英對照" reading doc, or a bilingual Word
translation of the paper. See `SKILL.md` for the full step-by-step
workflow, and `references/data-schema.md` + `references/example-content.js`
for the content format.

## License

MIT — see `LICENSE`.
