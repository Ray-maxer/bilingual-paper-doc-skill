#!/usr/bin/env python3
"""
inject_mathpr.py — add the <m:mathPr> settings block to a .docx produced by
docx-js when the document contains equation (m:oMath) objects.

Why this is needed:
docx-js's Math API emits valid, schema-correct <m:oMath> equation objects,
but it does NOT write the companion <m:mathPr> block into word/settings.xml
that real Word documents normally carry (default equation font, bracket
auto-sizing behavior, etc.). The document is still 100% valid OOXML without
it, but adding it makes the file's equation defaults match what Word itself
would generate, so keep this as a standard post-processing step.

CRITICAL: <m:mathPr> must sit at a specific point in the CT_Settings
element sequence — after <w:compat>/<w:docVars>/<w:rsids> (all optional)
and before <m:attachedSchema>/<w:themeFontLang>/etc. In practice, for
docx-js output (which never emits docVars/rsids), inserting it immediately
before </w:settings> is correct. Also note the element MUST use the "m:"
namespace prefix, not "w:" — using <w:mathPr> instead of <m:mathPr> is an
easy mistake that silently produces an invalid file (fails schema
validation with "element not expected" pointing at unrelated siblings).

Usage:
    python inject_mathpr.py input.docx output.docx
"""
import sys
import shutil
import zipfile
import tempfile
import os

MATHPR = (
    '<m:mathPr>'
    '<m:mathFont m:val="Cambria Math"/>'
    '<m:brkBin m:val="before"/>'
    '<m:brkBinSub m:val="--"/>'
    '<m:smallFrac m:val="0"/>'
    '<m:dispDef/>'
    '<m:lMargin m:val="0"/>'
    '<m:rMargin m:val="0"/>'
    '<m:defJc m:val="centerGroup"/>'
    '<m:wrapIndent m:val="1440"/>'
    '<m:intLim m:val="subSup"/>'
    '<m:naryLim m:val="undOvr"/>'
    '</m:mathPr>'
)


def inject(input_path: str, output_path: str) -> None:
    with tempfile.TemporaryDirectory() as tmp:
        with zipfile.ZipFile(input_path) as zf:
            zf.extractall(tmp)

        settings_path = os.path.join(tmp, "word", "settings.xml")
        with open(settings_path, encoding="utf-8") as f:
            data = f.read()

        if "<m:mathPr>" in data:
            print("mathPr already present, skipping injection")
        else:
            assert "</w:settings>" in data, "unexpected settings.xml structure"
            data = data.replace("</w:settings>", MATHPR + "</w:settings>")
            with open(settings_path, "w", encoding="utf-8") as f:
                f.write(data)

        if os.path.exists(output_path):
            os.remove(output_path)
        # Re-zip (must not use deflate-store tricks that break Word; a plain
        # zip -Xr equivalent is sufficient — shutil.make_archive handles it)
        base = output_path[:-4] if output_path.endswith(".docx") else output_path
        archive = shutil.make_archive(base, "zip", tmp)
        os.replace(archive, output_path)

    print(f"Wrote {output_path}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python inject_mathpr.py input.docx output.docx")
        sys.exit(1)
    inject(sys.argv[1], sys.argv[2])
