/**
 * render.js — CLI wrapper around lib.js.
 *
 * Usage:
 *   node render.js <content-data.js> <output.docx> [imageDir]
 *
 * <content-data.js> must `module.exports` a data object matching
 * ../references/data-schema.md (see ../references/example-content.js
 * for a full worked example).
 *
 * imageDir (optional) overrides data.imageDir — the folder figure files
 * (Fig1.png, Fig2.png, ...) are read from. Defaults to the content-data.js
 * file's own directory.
 */
const fs = require("fs");
const path = require("path");
const { renderDocument } = require("./lib.js");

const [, , dataPath, outPath, imageDirArg] = process.argv;
if (!dataPath || !outPath) {
  console.error("Usage: node render.js <content-data.js> <output.docx> [imageDir]");
  process.exit(1);
}

const absDataPath = path.resolve(dataPath);
const data = require(absDataPath);
if (imageDirArg) data.imageDir = path.resolve(imageDirArg);
else if (!data.imageDir) data.imageDir = path.dirname(absDataPath);

renderDocument(data).then((buf) => {
  fs.writeFileSync(outPath, buf);
  console.log("Wrote", outPath, buf.length, "bytes");
});
