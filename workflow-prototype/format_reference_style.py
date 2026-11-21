#!/usr/bin/env python3
"""Analytical Chemistry reference-style prototype. Isolated from the accepted formatter."""
from pathlib import Path
import re, subprocess, tempfile
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "Analytical_Chemistry_I_Formatted.docx"
OUT = ROOT / "workflow-prototype/output"
NAVY, BLUE, GOLD, WHITE, BORDER = "17365D", "EAF0F6", "B08B2B", "FFFFFF", "AAB7C4"
OPTION_RE = re.compile(r"^\\s*[A-D]\\s*\\.")
QUESTION_RE = re.compile(r"^\\s*\\d+\\s*\\.")
SECTION_RE = re.compile(r"^\\s*Question\\s*\\(\\s*\\d+\\s*\\)", re.I)
ARABIC_RE = re.compile(r"[\\u0600-\\u06ff]")

def set_shading(el, fill):
    pr = el.get_or_add_pPr() if hasattr(el, "get_or_add_pPr") else el.get_or_add_tcPr()
    shd = pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd"); pr.append(shd)
    shd.set(qn("w:fill"), fill)

def set_cell_borders(cell, color=BORDER, size="6"):
    tcpr = cell._tc.get_or_add_tcPr()
    borders = tcpr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders"); tcpr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = borders.find(qn("w:" + edge))
        if node is None:
            node = OxmlElement("w:" + edge); borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), size)
        node.set(qn("w:color"), color)

def set_run(run, size=10.5, bold=None, color=None):
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    if bold is not None: run.bold = bold
    if color: run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.rFonts
    if fonts is None:
        fonts = OxmlElement("w:rFonts"); rpr.insert(0, fonts)
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn("w:" + key), "Calibri")

def set_keep(p, next_=False, together=True):
    p.paragraph_format.keep_with_next = next_
    p.paragraph_format.keep_together = together
    p.paragraph_format.widow_control = True

def set_page_field(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = paragraph.add_run("Page ")
    set_run(r, 8, color="666666")
    fld = OxmlElement("w:fldSimple"); fld.set(qn("w:instr"), "PAGE")
    paragraph._p.append(fld)

def main():
    if not SRC.exists():
        raise SystemExit(f"Missing reference exam: {SRC}")
    OUT.mkdir(parents=True, exist_ok=True)
    doc = Document(SRC)

    # Consistent A4 exam page geometry; source content, tables, and figures remain intact.
    for sec in doc.sections:
        sec.page_width, sec.page_height = Cm(21), Cm(29.7)
        sec.top_margin, sec.bottom_margin = Cm(1.55), Cm(1.55)
        sec.left_margin, sec.right_margin = Cm(1.65), Cm(1.65)
        sec.header_distance, sec.footer_distance = Cm(.65), Cm(.65)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"; normal.font.size = Pt(10.5)
    normal.paragraph_format.space_after = Pt(3)
    normal.paragraph_format.line_spacing = 1.03

    nonempty = [p for p in doc.paragraphs if p.text.strip()]
    for idx, p in enumerate(nonempty):
        text = p.text.strip()
        pf = p.paragraph_format
        set_keep(p)

        # Institutional masthead and exam metadata.
        if idx < 2:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.space_before = Pt(2); pf.space_after = Pt(3)
            pf.keep_with_next = True
            for r in p.runs: set_run(r, 15 if idx == 0 else 10.5, True, NAVY)
            if idx == 0:
                set_shading(p._p, BLUE)
        elif re.search(r"Program:|Level:|Semester:|Date:|Time:|Name:", text, re.I):
            pf.space_before = Pt(1); pf.space_after = Pt(3)
            pf.keep_with_next = True
            for r in p.runs: set_run(r, 9.5, True, NAVY)
        elif re.search(r"Main Examination", text, re.I):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.space_before = Pt(6); pf.space_after = Pt(8)
            pf.keep_with_next = True
            for r in p.runs: set_run(r, 13, True, WHITE)
            set_shading(p._p, NAVY)
        elif SECTION_RE.match(text):
            pf.space_before = Pt(9); pf.space_after = Pt(4)
            pf.keep_with_next = True
            set_shading(p._p, BLUE)
            for r in p.runs: set_run(r, 11.5, True, NAVY)
        elif QUESTION_RE.match(text):
            pf.space_before = Pt(5); pf.space_after = Pt(2)
            pf.keep_with_next = True
            for r in p.runs: set_run(r, 10.5, True, NAVY)
        elif OPTION_RE.match(text):
            pf.left_indent = Cm(.35)
            pf.first_line_indent = Cm(-.35)
            pf.space_before = Pt(0); pf.space_after = Pt(1.5)
            pf.line_spacing = 1.0
            # Keep option lines visually grouped; no text is merged or rewritten.
            for r in p.runs: set_run(r, 9.8)
        elif re.match(r"^(Good luck|Show all your work|Choose the appropriate term)", text, re.I):
            pf.space_before = Pt(4); pf.space_after = Pt(4)
            for r in p.runs: set_run(r, 10, True, NAVY)
        else:
            pf.space_after = Pt(2.5)
            for r in p.runs: set_run(r, 10.2)

    # Tables (matching list and answer table): stable widths, compact padding, clear grid.
    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        for row_index, row in enumerate(table.rows):
            trpr = row._tr.get_or_add_trPr()
            if trpr.find(qn("w:cantSplit")) is None:
                trpr.append(OxmlElement("w:cantSplit"))
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                set_cell_borders(cell)
                tcpr = cell._tc.get_or_add_tcPr()
                margins = tcpr.find(qn("w:tcMar"))
                if margins is None:
                    margins = OxmlElement("w:tcMar"); tcpr.append(margins)
                for edge in ("top", "start", "bottom", "end"):
                    node = margins.find(qn("w:" + edge))
                    if node is None:
                        node = OxmlElement("w:" + edge); margins.append(node)
                    node.set(qn("w:w"), "70"); node.set(qn("w:type"), "dxa")
                if row_index == 0:
                    set_shading(cell._tc, NAVY)
                for p in cell.paragraphs:
                    p.paragraph_format.space_after = Pt(1)
                    p.paragraph_format.line_spacing = 1.0
                    for r in p.runs:
                        set_run(r, 8.8, row_index == 0,
                                WHITE if row_index == 0 else None)

    # Preserve all source headers/footers and existing page-number fields.
    for sec in doc.sections:
        footer = sec.footer.paragraphs[0]
        if not footer.text.strip() and not footer._p.xpath(".//w:fldSimple"):
            set_page_field(footer)

    output_docx = OUT / "Analytical_Chemistry_Reference_Style.docx"
    doc.save(output_docx)
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(
            ["libreoffice", "--headless", f"-env:UserInstallation=file://{tmp}/profile",
             "--convert-to", "pdf", "--outdir", str(OUT), str(output_docx)],
            check=True, capture_output=True, text=True
        )
    pdf = OUT / "Analytical_Chemistry_Reference_Style.pdf"
    info = subprocess.run(["pdfinfo", str(pdf)], check=True, capture_output=True, text=True).stdout
    pages = re.search(r"^Pages:\\s+(\\d+)", info, re.M)
    print(f"Created {output_docx.name} and {pdf.name}; pages={pages.group(1) if pages else 'unknown'}")

if __name__ == "__main__":
    main()
