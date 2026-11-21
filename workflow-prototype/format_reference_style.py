#!/usr/bin/env python3
"""Isolated reference-style mockup formatter. Never writes to the accepted formatter."""
from pathlib import Path
import re, subprocess, tempfile
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT=Path(__file__).resolve().parents[1]
SRC=ROOT/"workflow-prototype/input/marketing_exam.docx"
OUT=ROOT/"workflow-prototype/output"
NAVY="17365D"; BLUE="DCE6F1"; GOLD="B08B2B"; WHITE="FFFFFF"; BORDER="AAB7C4"
ARABIC_RE=re.compile(r"[\u0600-\u06ff]")
SECTION_RE=re.compile(r"^\s*(السؤال|القسم|الجزء|انتهت|تعليمات)",re.I)
META_RE=re.compile(r"^\s*(المقرر|التاريخ|الزمن|اسم الطالب|رقم الجلوس|الدرجة)",re.I)

def shade(element, fill):
    pr=element.get_or_add_pPr() if hasattr(element,"get_or_add_pPr") else element.get_or_add_tcPr()
    shd=pr.find(qn("w:shd"))
    if shd is None: shd=OxmlElement("w:shd"); pr.append(shd)
    shd.set(qn("w:fill"),fill)

def set_borders(element, color=BORDER, size="8"):
    pr=element.get_or_add_tcPr() if hasattr(element,"get_or_add_tcPr") else element.get_or_add_pPr()
    borders=pr.find(qn("w:tcBorders" if hasattr(element,"get_or_add_tcPr") else "w:pBdr"))
    if borders is None:
        borders=OxmlElement("w:tcBorders" if hasattr(element,"get_or_add_tcPr") else "w:pBdr"); pr.append(borders)
    for edge in ("top","left","bottom","right","insideH","insideV"):
        tag="w:"+edge; node=borders.find(qn(tag))
        if node is None: node=OxmlElement(tag); borders.append(node)
        node.set(qn("w:val"),"single"); node.set(qn("w:sz"),size); node.set(qn("w:color"),color)

def rtl(p, yes):
    ppr=p._p.get_or_add_pPr(); node=ppr.find(qn("w:bidi"))
    if yes:
        if node is None: node=OxmlElement("w:bidi"); ppr.append(node)
        node.set(qn("w:val"),"1"); p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    elif node is not None: ppr.remove(node)

def run_style(run, size=13, bold=None, color=None):
    is_ar=bool(ARABIC_RE.search(run.text or ""))
    run.font.name="Calibri"; run.font.size=Pt(size+(1 if is_ar else 0))
    if bold is not None: run.bold=bold
    if color: run.font.color.rgb=RGBColor.from_string(color)
    rpr=run._element.get_or_add_rPr(); fonts=rpr.rFonts
    if fonts is None: fonts=OxmlElement("w:rFonts"); rpr.insert(0,fonts)
    for k in ("ascii","hAnsi","eastAsia"): fonts.set(qn("w:"+k),"Calibri")
    fonts.set(qn("w:cs"),"Traditional Arabic" if is_ar else "Calibri")
    if is_ar:
        node=rpr.find(qn("w:rtl"))
        if node is None: node=OxmlElement("w:rtl"); rpr.append(node)
        node.set(qn("w:val"),"1")

def main():
    if not SRC.exists(): raise SystemExit(f"Missing workflow input: {SRC}")
    OUT.mkdir(parents=True,exist_ok=True)
    doc=Document(SRC)
    for sec in doc.sections:
        sec.page_width=Cm(21); sec.page_height=Cm(29.7)
        sec.top_margin=Cm(1.8); sec.bottom_margin=Cm(1.8)
        sec.left_margin=Cm(2); sec.right_margin=Cm(2)
        sec.header_distance=Cm(.8); sec.footer_distance=Cm(.8)
    normal=doc.styles["Normal"]; normal.font.name="Calibri"; normal.font.size=Pt(13)
    normal.paragraph_format.space_after=Pt(6); normal.paragraph_format.line_spacing=1.12
    nonempty=[p for p in doc.paragraphs if p.text.strip()]
    for idx,p in enumerate(nonempty):
        t=p.text.strip(); is_ar=bool(ARABIC_RE.search(t)); rtl(p,is_ar)
        pf=p.paragraph_format; pf.widow_control=True
        if idx<2:
            p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            pf.space_before=Pt(3); pf.space_after=Pt(5); pf.keep_with_next=True
            for r in p.runs: run_style(r,17 if idx==0 else 13,True,WHITE if idx==0 else NAVY)
            if idx==0:
                shade(p._p,NAVY)
                set_borders(p._p,GOLD,"16")
        elif SECTION_RE.match(t):
            pf.space_before=Pt(12); pf.space_after=Pt(6); pf.keep_with_next=True; pf.keep_together=True
            shade(p._p,BLUE); set_borders(p._p,GOLD,"8")
            for r in p.runs: run_style(r,15,True,NAVY)
        elif META_RE.match(t):
            pf.space_after=Pt(4); pf.keep_with_next=True
            for r in p.runs: run_style(r,13,True,NAVY)
        elif re.match(r"^\s*(?:السؤال\s+)?(?:[0-9٠-٩]+|[٠-٩]+)[.)：:]",t):
            pf.space_before=Pt(7); pf.space_after=Pt(4); pf.keep_with_next=True; pf.keep_together=True
            for r in p.runs: run_style(r,13,True,NAVY)
        else:
            for r in p.runs: run_style(r,13)
    for table in doc.tables:
        table.alignment=WD_TABLE_ALIGNMENT.CENTER; table.autofit=True
        for ri,row in enumerate(table.rows):
            trpr=row._tr.get_or_add_trPr(); trpr.append(OxmlElement("w:cantSplit"))
            for cell in row.cells:
                cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
                tcpr=cell._tc.get_or_add_tcPr()
                mar=tcpr.find(qn("w:tcMar"))
                if mar is None: mar=OxmlElement("w:tcMar"); tcpr.append(mar)
                for edge in ("top","start","bottom","end"):
                    node=OxmlElement("w:"+edge); node.set(qn("w:w"),"100"); node.set(qn("w:type"),"dxa"); mar.append(node)
                set_borders(cell)
                if ri==0:
                    shade(cell._tc,BLUE)
                for p in cell.paragraphs:
                    ar=bool(ARABIC_RE.search(p.text)); rtl(p,ar)
                    p.paragraph_format.space_after=Pt(2)
                    for r in p.runs: run_style(r,12,ri==0,NAVY if ri==0 else None)
    for sec in doc.sections:
        footer=sec.footer.paragraphs[0]; footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
        if not footer.text:
            r=footer.add_run()
            fld=OxmlElement("w:fldSimple"); fld.set(qn("w:instr"),"PAGE"); r._r.addnext(fld)
    dest=OUT/"Marketing_Principles_Reference_Style.docx"
    doc.save(dest)
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(["libreoffice","--headless",f"-env:UserInstallation=file://{tmp}/profile","--convert-to","pdf","--outdir",str(OUT),str(dest)],check=True,capture_output=True,text=True)
    pdf=OUT/"Marketing_Principles_Reference_Style.pdf"
    subprocess.run(["pdfinfo",str(pdf)],check=True,capture_output=True,text=True)
    print(f"Created {dest} and {pdf}")
if __name__=="__main__": main()
