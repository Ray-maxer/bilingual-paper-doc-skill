/**
 * example-content.js — a TRIMMED worked example (2 sections + 1 figure +
 * 1 equation), based on the mangrove wave-conditions paper this skill was
 * originally built for. Copy the SHAPE of this file for new papers; don't
 * literally reuse the content. This excerpt doesn't happen to include a
 * table — see the `table` block worked example in
 * ../references/data-schema.md instead. See that same file for the
 * full field reference.
 */
const { SUB, SUBSUP, FRAC } = require("../scripts/lib.js");

module.exports = {
  imageDir: "./imgs",

  meta: {
    citation: "Coastal Engineering 199 (2025) 104749",
    titleEn: "Field experiment study to assess critical wave conditions leading to failure of mangrove Rhizophora stylosa",
    titleZh: "紅樹林（Rhizophora stylosa）破壞臨界波浪條件之現地實驗研究",
    authors: "Nobuhito Mori, Che-Wei Chang, Kenji Ono, et al.",
    doiLine: "DOI: https://doi.org/10.1016/j.coastaleng.2025.104749　｜　中英對照閱讀版",
    headerRight: "Mori et al. (2025), Coastal Engineering 199 — 中英對照閱讀版",
  },

  usageNote: {
    title: "本文件使用說明 / How to read this document",
    text: "本檔案為論文逐段中英對照整理，供課程閱讀使用。灰底藍字段落為英文原文，藍色左框段落為對應中文翻譯；黃底方框為重點筆記。圖表已從原始 PDF 擷取並附上雙語圖說，公式以 Word 原生方程式物件呈現並保留原文編號。",
  },

  sections: [
    { type: "h1", en: "Abstract", zh: "摘要" },
    { type: "pair",
      en: "Mangroves can attenuate tsunamis, storm surges, and waves and significantly reduce coastal hazards...",
      zh: "紅樹林可以削弱海嘯、風暴潮與一般海浪，明顯降低沿海災害的風險..." },

    { type: "h1", en: "2. Outline of field survey and measurements", zh: "2. 現地調查與量測概述" },
    { type: "pair",
      en: "A field survey was conducted in 2022 to investigate the resistance forces on mangroves and their geometrical structure...",
      zh: "2022 年，研究團隊進行了一次現地調查，用來研究紅樹林所能承受的阻力（力量）以及它們的幾何結構..." },

    { type: "figure", file: "Fig1.png", w: 550, h: 307, num: 1,
      capEn: "Location of field survey at the Urauchi River, Iriomote Island, Japan.",
      capZh: "浦內川（西表島）現地調查地點位置圖。" },

    { type: "note", title: "重點筆記",
      text: "現地調查的地點位於日本沖繩西表島浦內川感潮河段的一塊濕地（浦內橋附近），並進行了水平的極限拉力試驗以測定最大彎曲抵抗值。" },

    { type: "h1", en: "3. Results and discussions", zh: "3. 結果與討論" },
    { type: "pair",
      en: "Mangrove trees experience various failure modes, with bending failure related to moments...",
      zh: "紅樹林樹木會面臨多種破壞模式，其中彎曲破壞與力矩有關..." },

    // Mmax = fmax × Hpull    (1)
    { type: "equation", label: "1",
      tokens: [SUB("M", "max"), " = ", SUB("f", "max"), " × ", SUB("H", "pull")] },

    // MLmax(DBH) = 166.19 · DBH − 5.18   (DBH ≥ 0.032 m)    (2)
    { type: "equation", label: "2",
      tokens: [SUBSUP("M", "max", "L"), "(DBH) = 166.19 · DBH − 5.18   (DBH ≥ 0.032 m)"] },

    // a fraction example: σLmax(DBH) = (32·(...) / 10π·DBH)   (6)
    { type: "equation", label: "6",
      tokens: [
        SUBSUP("σ", "max", "L"), "(DBH) = ",
        FRAC(["32·(0.166·DBH·10⁻³ − 5.178)"], ["10π·DBH"]),
        "   (DBH ≥ 0.032 m)",
      ] },
  ],

  referencesHeading: { en: "References", zh: "參考文獻（原文，未翻譯）" },
  references: [
    "Chang, C.W., Mori, N., 2021. Green infrastructure for the reduction of coastal disasters... Coastal Engineering Journal 63(3), 370–385.",
  ],
};
