#!/bin/bash
# extract_figures.sh — pull embedded raster images out of a paper PDF so they
# can be re-inserted into the bilingual docx with bilingual captions.
#
# Usage: ./extract_figures.sh <paper.pdf> <output_dir>
#
# This dumps ALL embedded images (including small logos/journal-cover
# thumbnails) as fig-<page>-<index>.png. You must still:
#   1. `view` each fig-*.png to identify which is Fig.1, Fig.2, ... and
#      which are noise (journal logo, checkmark icons, cover thumbnail).
#   2. cp/rename the real figures to Fig1.png, Fig2.png, ... for use in
#      content-data.js.
# pdfimages does not know which images are "real" figures vs page furniture,
# so this step always needs a human/Claude visual check — do not skip it.
set -e
PDF="$1"
OUTDIR="${2:-./imgs}"
mkdir -p "$OUTDIR"
pdfimages -png -p "$PDF" "$OUTDIR/fig"
echo "Extracted to $OUTDIR — now view each fig-*.png and identify Fig.1, Fig.2, ..."
ls -la "$OUTDIR"
