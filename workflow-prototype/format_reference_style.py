#!/usr/bin/env python3
"""Isolated reference-style exam formatter. Does not modify source documents."""
from pathlib import Path
import argparse, re, subprocess, tempfile
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "workflow-prototype/output"
NAVY, BLUE, GOLD, WHITE, BORDER = "17365D", "EAF0F6", "B08B2B", "FFFFFF", "AAB7C4"
OPTION_RE = re.compile(r"^\s*[A-D]\s*[.)]")
QUESTION_RE = re.compile(r"^\s*\d+\s*[.)]")
SECTION_RE = re.compile(r"^(?:Question\s*\(\s*\d+\s*\)|Section\s+[A-Z0-9]+)", re.I)

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
        node.set(qn("w:val"), "single"); node.set(qn("w:sz"), size); node.set(qn("w:color"), color)

def set_run(run, size=10.5, bold=None, color=None):
    run.font.name = "Calibri"; run.font.size = Pt(size)
    if bold is not None: run.bold = bold
    if color: run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.rFonts
    if fonts is None:
        fonts = OxmlElement("w:rFonts"); rpr.insert(0, fonts)
    for key in ("ascii", "hAnsi", "eastAsia", "cs"): fonts.set(qn("w:" + key), "Calibri")

def set_keep(p, next_=False):
    p.paragraph_format.keep_with_next = next_
    p.paragraph_format.keep_together = True
    p.paragraph_format.widow_control = True

def set_page_field(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = paragraph.add_run("Page "); set_run(r, 8, color="666666")
    fld = OxmlElement("w:fldSimple"); fld.set(qn("w:instr"), "PAGE"); paragraph._p.append(fld)

def convert_legacy_doc(source, tempdir):
    profile = Path(tempdir) / "lo-profile"
    cmd = ["libreoffice", "--headless", f"-env:UserInstallation=file://{profile}",
           "--convert-to", "docx", "--outdir", tempdir, str(source)]
    subprocess.run(cmd, check=True, capture_output=True, text=True)
    converted = Path(tempdir) / (source.stem + ".docx")
    if not converted.exists(): raise RuntimeError("LibreOffice did not produce a DOCX from the legacy DOC source")
    return converted

def remove_forced_page_breaks(doc):
    """Let content paginate naturally; legacy DOC imports often carry stale hard breaks."""
    paragraphs = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                paragraphs.extend(cell.paragraphs)
    for p in paragraphs:
        p.paragraph_format.page_break_before = False
        for node in list(p._p.xpath(".//w:br[@w:type='page'] | .//w:lastRenderedPageBreak")):
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)

def format_exam(source, output_prefix):
    if not source.exists(): raise SystemExit(f"Missing source exam: {source}")
    OUT.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        working_source = convert_legacy_doc(source, tmp) if source.suffix.lower() == ".doc" else source
        doc = Document(working_source)
        remove_forced_page_breaks(doc)

        # Match the reference's clean A4 geometry while retaining source content and objects.
        for sec in doc.sections:
            # Preserve section formatting without forcing each imported section onto a new page.
            sec.start_type = WD_SECTION_START.CONTINUOUS
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
            txt = p.text.strip(); pf = p.paragraph_format
            set_keep(p)
            if idx < 2:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pf.space_before = Pt(2); pf.space_after = Pt(3); pf.keep_with_next = True
                for r in p.runs: set_run(r, 15 if idx == 0 else 10.5, True, NAVY)
                if idx == 0: set_shading(p._p, BLUE)
            elif re.search(r"Program:|Level:|Semester:|Date:|Time:|Name:|Department:|Course:", txt, re.I):
                pf.space_before = Pt(1); pf.space_after = Pt(3); pf.keep_with_next = True
                for r in p.runs: set_run(r, 9.5, True, NAVY)
            elif re.search(r"Main Examination|Final Examination|Mid.?term Examination", txt, re.I):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pf.space_before = Pt(6); pf.space_after = Pt(8); pf.keep_with_next = True
                for r in p.runs: set_run(r, 13, True, WHITE)
                set_shading(p._p, NAVY)
            elif SECTION_RE.match(txt):
                pf.space_before = Pt(9); pf.space_after = Pt(4); pf.keep_with_next = True
                set_shading(p._p, BLUE)
                for r in p.runs: set_run(r, 11.5, True, NAVY)
            elif OPTION_RE.match(txt):
                pf.left_indent = Cm(.35); pf.first_line_indent = Cm(-.35)
                pf.space_before = Pt(0); pf.space_after = Pt(1.5); pf.line_spacing = 1.0
                for r in p.runs: set_run(r, 9.8)
            elif QUESTION_RE.match(txt):
                pf.space_before = Pt(5); pf.space_after = Pt(2); pf.keep_with_next = True
                for r in p.runs: set_run(r, 10.5, True, NAVY)
            elif re.match(r"^(Good luck|Show all your work|Choose the appropriate term)", txt, re.I):
                pf.space_before = Pt(4); pf.space_after = Pt(4)
                for r in p.runs: set_run(r, 10, True, NAVY)
            else:
                pf.space_after = Pt(2.5)
                for r in p.runs: set_run(r, 10.2)

        for table in doc.tables:
            table.alignment = WD_TABLE_ALIGNMENT.CENTER; table.autofit = False
            for ri, row in enumerate(table.rows):
                trpr = row._tr.get_or_add_trPr()
                if trpr.find(qn("w:cantSplit")) is None: trpr.append(OxmlElement("w:cantSplit"))
                for cell in row.cells:
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    set_cell_borders(cell)
                    tcpr = cell._tc.get_or_add_tcPr()
                    margins = tcpr.find(qn("w:tcMar"))
                    if margins is None: margins = OxmlElement("w:tcMar"); tcpr.append(margins)
                    for edge in ("top", "start", "bottom", "end"):
                        node = margins.find(qn("w:" + edge))
                        if node is None: node = OxmlElement("w:" + edge); margins.append(node)
                        node.set(qn("w:w"), "70"); node.set(qn("w:type"), "dxa")
                    if ri == 0: set_shading(cell._tc, NAVY)
                    for p in cell.paragraphs:
                        p.paragraph_format.space_after = Pt(1); p.paragraph_format.line_spacing = 1.0
                        for r in p.runs: set_run(r, 8.8, ri == 0, WHITE if ri == 0 else None)

        for sec in doc.sections:
            footer = sec.footer.paragraphs[0]
            if not footer.text.strip() and not footer._p.xpath(".//w:fldSimple"): set_page_field(footer)

        output_docx = OUT / f"{output_prefix}.docx"
        doc.save(output_docx)
        profile = Path(tmp) / "pdf-profile"
        subprocess.run(["libreoffice", "--headless", f"-env:UserInstallation=file://{profile}",
                        "--convert-to", "pdf", "--outdir", str(OUT), str(output_docx)],
                       check=True, capture_output=True, text=True)

    pdf = OUT / f"{output_prefix}.pdf"
    info = subprocess.run(["pdfinfo", str(pdf)], check=True, capture_output=True, text=True).stdout
    pages = re.search(r"^Pages:\s+(\d+)", info, re.M)
    print(f"Created {output_docx.name} and {pdf.name}; pages={pages.group(1) if pages else 'unknown'}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="DOC or DOCX source exam path")
    parser.add_argument("--output-prefix", required=True, help="Output basename without extension")
    args = parser.parse_args()
    format_exam((ROOT / args.input).resolve(), args.output_prefix)
