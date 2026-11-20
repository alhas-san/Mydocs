from pathlib import Path
import subprocess
import tempfile
import shutil
import sys

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "Formatted"
INPUT_DIRS = [ROOT, ROOT / "Incoming"]
REMOVE_GRAPHICS = True  # Set to False if diagrams/images should be preserved.
EXPORT_PDF = True


def convert_legacy_doc(source: Path, destination_dir: Path) -> Path:
    """Convert old binary .doc files to .docx using LibreOffice."""
    result = subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "docx",
         "--outdir", str(destination_dir), str(source)],
        capture_output=True, text=True
    )
    converted = destination_dir / (source.stem + ".docx")
    if result.returncode != 0 or not converted.exists():
        raise RuntimeError(
            f"Could not convert {source.name} to DOCX. "
            f"LibreOffice says: {result.stderr or result.stdout}"
        )
    return converted


def set_run_font(run, name="Times New Roman", size=12, bold=None):
    run.font.name = name
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for key in ("ascii", "hAnsi", "cs", "eastAsia"):
        rfonts.set(qn(f"w:{key}"), name)


def remove_graphics(doc):
    # Remove embedded drawings, legacy pictures and OLE objects while retaining text.
    tags = {qn("w:drawing"), qn("w:pict"), qn("w:object")}
    for paragraph in doc.paragraphs:
        for element in list(paragraph._p.iter()):
            if element.tag in tags and element.getparent() is not None:
                element.getparent().remove(element)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for element in list(paragraph._p.iter()):
                        if element.tag in tags and element.getparent() is not None:
                            element.getparent().remove(element)


def format_document(source: Path, output_docx: Path):
    doc = Document(str(source))

    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.9)
    section.right_margin = Inches(0.9)

    if REMOVE_GRAPHICS:
        remove_graphics(doc)

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        paragraph.paragraph_format.space_after = Pt(6)
        paragraph.paragraph_format.line_spacing = 1.15
        upper = text.upper()
        if (upper.startswith(("SECTION", "PART", "QUESTION", "MULTIPLE CHOICE",
                              "TRUE OR FALSE", "MATCHING", "FILL IN",
                              "SHORT ANSWER", "DIAGRAM"))):
            paragraph.paragraph_format.keep_with_next = True
            for run in paragraph.runs:
                set_run_font(run, size=12, bold=True)
        else:
            for run in paragraph.runs:
                set_run_font(run, size=12)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_after = Pt(3)
                    for run in paragraph.runs:
                        set_run_font(run, size=11)

    # Apply a consistent font to headers and footers as well.
    for section in doc.sections:
        for part in (section.header, section.footer):
            for paragraph in part.paragraphs:
                for run in paragraph.runs:
                    set_run_font(run, size=9)

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_docx))


def main():
    OUTPUT.mkdir(exist_ok=True)
    candidates = []
    for folder in INPUT_DIRS:
        if folder.exists():
            candidates.extend(
                p for p in folder.iterdir()
                if p.is_file() and p.suffix.lower() in {".doc", ".docx"}
                and not p.name.startswith("~$")
            )

    if not candidates:
        print("No Word documents found in repository root or Incoming/.")
        return

    failures = []
    for source in sorted(set(candidates)):
        try:
            with tempfile.TemporaryDirectory() as temp:
                temp_dir = Path(temp)
                if source.suffix.lower() == ".doc":
                    converted = convert_legacy_doc(source, temp_dir)
                else:
                    converted = source

                relative_stem = source.stem
                output_docx = OUTPUT / f"{relative_stem}_Formatted.docx"
                format_document(converted, output_docx)

                if EXPORT_PDF:
                    pdf_result = subprocess.run(
                        ["libreoffice", "--headless", "--convert-to", "pdf",
                         "--outdir", str(OUTPUT), str(output_docx)],
                        capture_output=True, text=True
                    )
                    expected_pdf = OUTPUT / f"{output_docx.stem}.pdf"
                    if pdf_result.returncode != 0 or not expected_pdf.exists():
                        raise RuntimeError(
                            f"PDF export failed: {pdf_result.stderr or pdf_result.stdout}"
                        )
                print(f"Formatted: {source.relative_to(ROOT)}")
        except Exception as exc:
            failures.append(f"{source.name}: {exc}")
            print(f"ERROR: {source.name}: {exc}", file=sys.stderr)

    if failures:
        raise SystemExit("\n".join(failures))


if __name__ == "__main__":
    main()
