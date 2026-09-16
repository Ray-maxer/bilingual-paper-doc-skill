# content-data.js schema

`content-data.js` is a plain Node module: `module.exports = { ... }`. It holds
ALL paper-specific content (translated text, figure list, equations,
references). `lib.js` is 100% generic and never changes between papers.

```js
const { SUB, SUP, SUBSUP, FRAC, INT } = require("./lib.js"); // equation token helpers

module.exports = {
  // ---- optional visual style overrides ----
  style: {
    enColor: "1F1F1F", zhColor: "0B5394", headColor: "1B4F72",
    enFont: "Calibri", zhFont: "Microsoft JhengHei",
  },

  // ---- where figure PNGs live (Fig1.png, Fig2.png, ...) ----
  imageDir: "./imgs",   // defaults to content-data.js's own directory if omitted

  // ---- table of contents ----
  // Auto-generated from real h1/h2 headings as a native, clickable Word TOC
  // field on its own page right after the cover. On by default; set to
  // `false` for very short docs where a TOC would be pointless.
  toc: true,

  // ---- cover page ----
  // titleEn, authors, citation, and doiLine also populate the actual Word
  // file properties (title/creator/subject/description) — the values that
  // show up in Word's Info panel and get picked up by reference managers
  // like Zotero, so keep them accurate to the source paper, not shortened.
  meta: {
    citation: "Journal Name 199 (2025) 104749",
    titleEn: "English paper title",
    titleZh: "中文標題",
    authors: "A. Author, B. Author, ...",
    doiLine: "DOI: https://doi.org/... ｜ 中英對照閱讀版",
    headerRight: "Author et al. (2025), Journal — 中英對照閱讀版",  // running header
  },

  // ---- optional "how to read this doc" callout on the cover ----
  usageNote: {
    title: "本文件使用說明 / How to read this document",
    text: "灰底藍字段落為英文原文，藍色左框段落為對應中文翻譯...",
  },

  // ---- body: an ORDERED array of blocks, rendered top to bottom ----
  sections: [
    { type: "h1", en: "Abstract", zh: "摘要" },
    { type: "pair", en: "English paragraph text...", zh: "對應中文翻譯..." },

    { type: "h1", en: "1. Introduction", zh: "1. 前言" },
    { type: "pair", en: "...", zh: "..." },
    { type: "pair", en: "...", zh: "..." },

    { type: "h2", en: "2.1 Subsection", zh: "2.1 小節" },

    // a figure: file must exist at <imageDir>/<file>
    { type: "figure", file: "Fig1.png", w: 550, h: 307,
      capEn: "Caption in English.", capZh: "中文圖說。", num: 1 },

    // a numbered display equation — see "Equation tokens" below
    { type: "equation", label: "1",
      tokens: [SUB("M", "max"), " = ", SUB("f", "max"), " × ", SUB("H", "pull")] },

    // a data/results table — caption sits ABOVE the table (academic
    // convention; figure captions sit below theirs instead). headersZh is
    // optional — omit it if the column headers don't need translation.
    // Cell values are transcribed as-is from the source (see Translation
    // fidelity above) — don't recompute or round anything.
    { type: "table", num: 1,
      capEn: "Summary of measured values.", capZh: "量測數值摘要。",
      headers: ["Variable", "Value", "Unit"],
      headersZh: ["變數", "數值", "單位"],
      rows: [
        ["Mmax", "12.3", "N·m"],
        ["fmax", "45.6", "N"],
      ] },

    // a yellow reading-note callout box
    { type: "note", title: "重點筆記", text: "..." },

    // a bulleted list (e.g. "reading takeaways" section)
    { type: "bullets", items: ["第一點...", "第二點..."] },

    { type: "spacer", h: 200 },       // vertical gap (twips), default 120
    { type: "pageBreak" },
    { type: "en", en: "English-only paragraph (no Chinese pair)" },
    { type: "zh", zh: "純中文段落（無對應英文）" },
  ],

  // ---- references, left untranslated (English only) ----
  referencesHeading: { en: "References", zh: "參考文獻（原文，未翻譯）" }, // optional
  references: [
    "Author, A., 2024. Title of paper. Journal 1, 1–10.",
    "...",
  ],

  // ---- optional: extra inline variable subscript/superscript rules ----
  // Only needed if THIS paper uses variable names not already covered by
  // lib.js's DEFAULT_INLINE_RULES (Mmax, fmax, CD, CM, θmax, σmax, Hpull,
  // FD, FM, F̄theorymax, Mtheorymax, ML/Qmax, fL/Qmax, θL/Qmax, DBH², H², a², R²).
  // Same shape as lib.js's rule objects; matched BEFORE the defaults.
  extraInlineRules: [
    { re: /Reyn/g, base: "Re", sub: "critical" },  // example only
  ],
};
```

## Equation tokens

Import `{ SUB, SUP, SUBSUP, FRAC, INT }` from `lib.js` to build the `tokens`
array for a `type: "equation"` block. A token array is a mix of plain
strings and these builders:

- `SUB(base, sub)` → base with a subscript, e.g. `SUB("M", "max")` → M_max
- `SUP(base, sup)` → base with a superscript, e.g. `SUP("DBH", "2")` → DBH²
- `SUBSUP(base, sub, sup)` → both at once, e.g. `SUBSUP("M", "max", "L")` → M^L_max
- `FRAC(num, den)` → a fraction; `num`/`den` are themselves token arrays,
  e.g. `FRAC(["gka"], ["ω"])` → (gka)/(ω)
- `INT(base, sub, sup)` → an integral sign with limits, e.g.
  `INT(["u|u| dA"], "−h", "0")` → ∫₋ₕ⁰ u|u| dA

`base` (for SUB/SUP/SUBSUP) may itself be a token array instead of a plain
string, so you can nest — e.g. raising a fraction to a power:
`SUP(["(", FRAC(["gka"], ["ω·cosh kh"]), ")"], "2")`.

These compile to REAL Word equation objects (OOXML `m:oMath`), not styled
text — they're double-clickable and editable in Word's own Equation Editor.

## Block types reference

| `type`      | required fields                          | notes |
|-------------|-------------------------------------------|-------|
| `h1`        | `en`, `zh`                                 | section heading, underlined |
| `h2`        | `en`, `zh`                                 | subsection heading |
| `pair`      | `en`, `zh`                                 | the standard EN paragraph → ZH translation block |
| `en`        | `en`                                       | English-only paragraph |
| `zh`        | `zh`                                       | Chinese-only paragraph |
| `equation`  | `tokens`, `label`                          | numbered display equation, see above |
| `figure`    | `file`, `w`, `h`, `capEn`, `capZh`, `num`  | `w`/`h` in pixels (~96dpi); see figure sizing note below |
| `table`     | `headers`, `rows`, `num`                   | data table; `headersZh`, `capEn`, `capZh` optional; caption renders above the table |
| `note`      | `title`, `text`                            | yellow callout box (Chinese, matches original hand-annotated notes) |
| `bullets`   | `items` (array of strings)                 | bulleted list |
| `spacer`    | `h` (optional, default 120)                | vertical gap in twips |
| `pageBreak` | —                                           | forces a new page |

### Figure sizing
Pick `w` so the image fits within the ~6.2in content width (page margins are
1in each side on US-Letter): `w` up to ~550px for a full-width figure. Compute
`h = w * (originalHeightPx / originalWidthPx)` to preserve aspect ratio — get
original dimensions with `python3 -c "from PIL import Image; print(Image.open('Fig1.png').size)"`.
For very tall (portrait, multi-panel) figures, use a smaller `w` (~340–400px)
so the whole figure fits on one page without being crushed sideways.
