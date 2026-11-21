#!/usr/bin/env python3
"""Isolated reference-style mockup formatter. Never writes over source documents."""
from __future__ import annotations
import argparse, re, subprocess
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

NAVY = "183153"
PALE = "EAF0F6"
MID = "AAB8C8"
DARK = "202A35"

def arabic(text):
    return any("\u0600" <= c <= "\u06ff" or "\u0750" <= c <= "\u077f" or "\u08a0" <= c <= "\u08ff" for c in text)

def set_rtl(p, enabled=True):
    ppr = p._p.get_or_add_pPr()
    node = ppr.find(qn("w:bidi"))
    if enabled and node is None:
        node = OxmlElement("w:bidi")
        ppr.append(node)
    elif not enabled and node is not None:
        ppr.remove(node)
    if enabled:
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

def set_font(run, size=14, bold=None, rtl=True, color=None):
    name = "Traditional Arabic" if rtl else "Calibri"
    run.font.name = name
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.rFonts
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn("w:" + key), name)
    if rtl:
        node = rpr.find(qn("w:rtl"))
        if node is None:
            node = OxmlElement("w:rtl")
            rpr.append(node)
        node.set(qn("w:val"), "1")

def set_borders(cell, color=MID, size="6"):
    tcpr = cell._tc.get_or_add_tcPr()
    borders = tcpr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcpr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = qn("w:" + edge)
        node = borders.find(tag)
        if node is None:
            node = OxmlElement("w:" + edge)
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), size)
        node.set(qn("w:color"), color)

def shade(cell, fill):
    tcpr = cell._tc.get_or_add_tcPr()
    node = tcpr.find(qn("w:shd"))
    if node is None:
        node = OxmlElement("w:shd")
        tcpr.append(node)
    node.set(qn("w:fill"), fill)

def margins(cell, value=100):
    tcpr = cell._tc.get_or_add_tcPr()
    mar = tcpr.first_child_found_in("w:tcMar")
    if mar is None:
        mar = OxmlElement("w:tcMar")
        tcpr.append(mar)
    for edge in ("top", "start", "bottom", "end"):
        node = mar.find(qn("w:" + edge))
        if node is None:
            node = OxmlElement("w:" + edge)
            mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")

def para_style(p, size=14, bold=False, after=4, centered=False):
    rtl = arabic(p.text)
    set_rtl(p, rtl)
    if centered:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(after)
    pf.line_spacing = 1.08
    pf.widow_control = True
    for run in p.runs:
        set_font(run, size=size, bold=(bold or run.bold), rtl=rtl)

def source_text(doc):
    values = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            values.extend(cell.text for cell in row.cells)
    for section in doc.sections:
        values.extend(p.text for p in section.header.paragraphs)
        values.extend(p.text for p in section.footer.paragraphs)
    return re.sub(r"\s+", " ", " ".join(values)).strip()

def add_title_frame(doc, title):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Cm(17.5)
    cell = table.cell(0, 0)
    cell.width = Cm(17.5)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    margins(cell, 190)
    set_borders(cell, color="111111", size="18")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_rtl(p, arabic(title))
    run = p.add_run(title)
    set_font(run, size=22, bold=True, rtl=arabic(title))
    # Insert the frame before all existing body content.
    body = doc._element.body
    body.remove(table._tbl)
    first = next((node for node in body if node.tag != qn("w:sectPr")), None)
    if first is not None:
        first.addprevious(table._tbl)
    else:
        body.insert(0, table._tbl)

def format_reference_style(src: Path, dest: Path):
    doc = Document(str(src))
    original_text = source_text(doc)
    title = ""
    for section in doc.sections:
        footer_text = " ".join(p.text for p in section.footer.paragraphs)
        match = re.search(r"([^|]+?)\s*\|\s*صفحة", footer_text)
        if match:
            title = match.group(1).strip()
            break
    if not title:
        title = src.stem.replace("_", " ")
    add_title_frame(doc, title)

    # Normalize legacy hard page breaks and page-start section breaks only in this prototype copy.
    for node in list(doc._element.iter()):
        if node.tag == qn("w:pageBreakBefore") and node.getparent() is not None:
            node.getparent().remove(node)
        elif node.tag == qn("w:br") and node.get(qn("w:type")) == "page" and node.getparent() is not None:
            node.getparent().remove(node)
        elif node.tag == qn("w:sectPr"):
            typ = node.find(qn("w:type"))
            if typ is not None and typ.get(qn("w:val")) in ("nextPage", "oddPage", "evenPage"):
                typ.set(qn("w:val"), "continuous")

    for section in doc.sections:
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(1.65)
        section.bottom_margin = Cm(1.7)
        section.left_margin = Cm(1.65)
        section.right_margin = Cm(1.65)
        section.header_distance = Cm(0.7)
        section.footer_distance = Cm(0.7)

    for p in doc.paragraphs:
        if not p.text.strip():
            continue
        text = p.text.strip()
        if text in ("انتهت الأسئلة", "الله الموفق"):
            para_style(p, size=16, bold=True, after=8, centered=True)
        elif re.match(r"^\d+\s*[.)]", text) or text.startswith(("السؤال الثالث", "السؤال الرابع", "السؤال الخامس")):
            para_style(p, size=15, bold=True, after=5)
        else:
            para_style(p, size=14, after=4)

    for ti, table in enumerate(doc.tables):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        if ti == 0 and len(table.rows) == 1 and len(table.columns) == 1:
            # Keep the framed title visually dominant; it is not an exam content table.
            continue
        table.autofit = True
        # Two-cell section/marks bars get a light formal band.
        is_question_band = len(table.rows) == 1 and len(table.columns) == 2
        # Only actual column-heading rows are shaded; MCQ choice grids stay white.
        first_row = " ".join(c.text for c in table.rows[0].cells) if table.rows else ""
        is_header = len(table.rows) > 1 and any(k in first_row for k in ("العبارة", "الإجابة", "العمود", "م"))
        for ri, row in enumerate(table.rows):
            trpr = row._tr.get_or_add_trPr()
            if trpr.find(qn("w:cantSplit")) is None:
                trpr.append(OxmlElement("w:cantSplit"))
            if ri == 0 and is_header:
                repeat = OxmlElement("w:tblHeader")
                repeat.set(qn("w:val"), "true")
                trpr.append(repeat)
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                margins(cell, 100 if len(table.columns) > 1 else 125)
                set_borders(cell, color=NAVY if is_question_band else MID,
                            size="10" if is_question_band else "5")
                if is_question_band or (ri == 0 and is_header):
                    shade(cell, PALE)
                for p in cell.paragraphs:
                    para_style(p, size=14 if is_question_band else 13.5,
                               bold=is_question_band or (ri == 0 and is_header),
                               after=2)
                    for run in p.runs:
                        if is_question_band or (ri == 0 and is_header):
                            run.font.color.rgb = RGBColor.from_string(DARK)

    # Keep the existing footer wording; only normalize its typeface and alignment.
    for section in doc.sections:
        for p in section.footer.paragraphs:
            para_style(p, size=10, centered=True)

    dest.parent.mkdir(parents=True, exist_ok=True)
    doc.save(dest)
    # The title frame is the only intentionally repeated text (a visual title
    # sourced from the original footer); all source text remains present.
    check = Document(str(dest))
    check_title = check.tables[0] if check.tables else None
    if check_title is not None:
        # Compare all non-title content against the original.
        check_title._element.getparent().remove(check_title._element)
    if source_text(check) != original_text:
        raise RuntimeError("Source wording changed during reference-style formatting")
    return dest

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="Repository-relative source DOCX")
    parser.add_argument("--output-dir", default="prototype-output")
    args = parser.parse_args()
    src = Path(args.input).resolve()
    if not src.is_file() or src.suffix.lower() != ".docx":
        raise SystemExit(f"Input must be an existing .docx file: {src}")
    root = Path.cwd().resolve()
    try:
        src.relative_to(root)
    except ValueError:
        raise SystemExit("Input must be inside the checked-out repository")
    dest = Path(args.output_dir).resolve() / (src.stem + "_ReferenceStyle_Mockup.docx")
    format_reference_style(src, dest)
    print(f"MOCKUP_DOCX={dest}")
    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf",
                    "--outdir", str(dest.parent), str(dest)], check=True)
    pdf = dest.with_suffix(".pdf")
    if not pdf.is_file() or pdf.stat().st_size == 0:
        raise SystemExit("PDF export failed")
    subprocess.run(["pdfinfo", str(pdf)], check=True)
    subprocess.run(["pdftoppm", "-f", "1", "-l", "3", "-png", "-r", "110",
                    str(pdf), str(dest.parent / (src.stem + "-page"))], check=True)
    print(f"MOCKUP_PDF={pdf}")

if __name__ == "__main__":
    main()
