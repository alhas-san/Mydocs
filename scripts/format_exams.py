#!/usr/bin/env python3
"""Professional, content-preserving exam formatter."""
from __future__ import annotations
import json, re, subprocess, sys, tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT=Path(__file__).resolve().parents[1]
OUTPUT=ROOT/"Formatted"
INPUT_DIRS=(ROOT,ROOT/"Incoming")
REMOVE_GRAPHICS=True
EXPORT_PDF=True
FONT="Times New Roman"
BODY=11
GRAPHICS={qn("w:drawing"),qn("w:pict"),qn("w:object")}
QUESTION=re.compile(r"^\s*(?:(?:question|q)\s*\d{1,3}\b|\d{1,3}\s*[\).:-])",re.I)
SUBQUESTION=re.compile(r"^\s*(?:\([a-z0-9]+\)|[a-z]\s*[\).])",re.I)
OPTION_MARKER=re.compile(r"(?<!\S)(?:\(([A-D])\)|([A-D])[.)])\s*",re.I)
SECTION=re.compile(r"^\s*(?:section|part|instructions?|multiple[\s-]+choice|true\s+(?:or|/)\s+false|short[\s-]+answer|essay|answer\s+all|answer\s+the\s+following)\b",re.I)
META=re.compile(r"^\s*(?:date|time|duration|course|subject|semester|total\s+marks?|maximum\s+marks?|student\s+name|student\s+id|registration\s+no)\s*:",re.I)

def command(args,description):
    r=subprocess.run(args,capture_output=True,text=True)
    if r.returncode: raise RuntimeError(f"{description} failed: {(r.stderr or r.stdout or 'No diagnostic output').strip()}")
    return r

def lo(*args,profile):
    return ["libreoffice","--headless",f"-env:UserInstallation={profile.resolve().as_uri()}",*args]

def convert_doc(src,tmp):
    command(lo("--convert-to","docx","--outdir",str(tmp),str(src),profile=tmp/"lo-profile"),f"DOC conversion: {src.name}")
    out=tmp/f"{src.stem}.docx"
    if not out.is_file() or not out.stat().st_size: raise RuntimeError("Legacy DOC conversion produced no DOCX")
    return out

def paragraphs(container)->Iterable:
    yield from getattr(container,"paragraphs",[])
    for table in getattr(container,"tables",[]):
        for row in table.rows:
            for cell in row.cells: yield from paragraphs(cell)

def all_paragraphs(doc):
    result=[]; seen=set()
    containers=[doc]
    for sec in doc.sections:
        containers += [getattr(sec,n) for n in ("header","footer","first_page_header","first_page_footer","even_page_header","even_page_footer")]
    for container in containers:
        for p in paragraphs(container):
            if id(p._p) not in seen: seen.add(id(p._p)); result.append(p)
    return result

def snapshot(doc):
    # Compare actual wording while allowing intentional paragraph/line-break
    # changes made by the MCQ layout pass.
    ps=list(paragraphs(doc)); values=[p.text for p in ps]
    content=re.sub(r"\s+"," "," ".join(values)).strip()
    return {"paragraphs":len(ps),"nonempty":sum(bool(x.strip()) for x in values),"text":content,"characters":len(content)}

def parse_options(text):
    """Return sequential A-D choices found at the start of an option paragraph."""
    raw=text.strip()
    matches=list(OPTION_MARKER.finditer(raw))
    if not matches or raw[:matches[0].start()].strip():
        return []
    labels=[(m.group(1) or m.group(2)).upper() for m in matches]
    sequence=list("ABCD")
    start=sequence.index(labels[0]) if labels[0] in sequence else -1
    if start<0 or labels!=sequence[start:start+len(labels)]:
        return []
    return [(labels[i], raw[m.start():matches[i+1].start()].strip() if i+1<len(matches)
             else raw[m.start():].strip()) for i,m in enumerate(matches)]

def options_fit_together(first, second):
    # Conservative line-width estimate for 11 pt Times New Roman on A4 with
    # the configured margins. The limit leaves room for the tab separator.
    return len(first)<=42 and len(second)<=42 and len(first)+len(second)+4<=86

def reflow_mcq_options(doc):
    """Place short MCQ choices in pairs; keep long choices on separate lines."""
    ps=list(doc.paragraphs)
    changed=0
    i=0
    while i<len(ps):
        choices=parse_options(ps[i].text)
        if not choices or choices[0][0]!="A":
            i+=1
            continue
        collected=list(choices)
        remove=[]
        expected=len(collected)
        j=i+1
        while expected<4 and j<len(ps):
            following=parse_options(ps[j].text)
            if len(following)!=1 or following[0][0]!=list("ABCD")[expected]:
                break
            collected.extend(following)
            remove.append(ps[j])
            expected+=1
            j+=1
        if len(collected)<2:
            i+=1
            continue
        lines=[]
        k=0
        while k<len(collected):
            if k+1<len(collected) and options_fit_together(collected[k],collected[k+1]):
                lines.append(collected[k]+"\t"+collected[k+1])
                k+=2
            else:
                lines.append(collected[k][1])
                k+=1
        ps[i].text="\n".join(lines)
        ps[i].paragraph_format.tab_stops.add_tab_stop(Cm(8.2))
        for p in remove:
            p._element.getparent().remove(p._element)
        changed+=len(collected)
        i=j
    return changed

def remove_graphics(doc):
    count=0; seen=set()
    for part in doc.part.package.parts:
        if id(part) in seen: continue
        seen.add(id(part)); root=getattr(part,"_element",None)
        if root is None: continue
        for el in list(root.iter()):
            if el.tag in GRAPHICS and el.getparent() is not None:
                el.getparent().remove(el); count+=1
    return count

def font(run,size=BODY,bold=None):
    run.font.name=FONT; run.font.size=Pt(size)
    if bold is not None: run.bold=bold
    rpr=run._element.get_or_add_rPr(); rf=rpr.rFonts
    if rf is None: rf=OxmlElement("w:rFonts"); rpr.insert(0,rf)
    for k in ("ascii","hAnsi","cs","eastAsia"): rf.set(qn(f"w:{k}"),FONT)

def classify(p,index):
    t=p.text.strip()
    if not t: return "blank"
    if index<3 and len(t)<140 and not QUESTION.match(t): return "opening"
    if SECTION.match(t) or (len(t)<=85 and t.isupper() and any(c.isalpha() for c in t)): return "section"
    if QUESTION.match(t): return "question"
    if SUBQUESTION.match(t): return "subquestion"
    if META.match(t): return "metadata"
    return "body"

def cell_margins(cell):
    tcpr=cell._tc.get_or_add_tcPr(); mar=tcpr.first_child_found_in("w:tcMar")
    if mar is None: mar=OxmlElement("w:tcMar"); tcpr.append(mar)
    for edge,value in (("top",90),("start",110),("bottom",90),("end",110)):
        node=mar.find(qn("w:"+edge))
        if node is None: node=OxmlElement("w:"+edge); mar.append(node)
        node.set(qn("w:w"),str(value)); node.set(qn("w:type"),"dxa")

def page_number(p):
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(); font(r,9)
    for tag,attrs,text in [
        ("w:fldChar",{"w:fldCharType":"begin"},None),
        ("w:instrText",{"xml:space":"preserve"}," PAGE "),
        ("w:fldChar",{"w:fldCharType":"separate"},None),
        ("w:t",{},"1"),
        ("w:fldChar",{"w:fldCharType":"end"},None)]:
        el=OxmlElement(tag)
        for k,v in attrs.items(): el.set(qn(k),v)
        if text is not None: el.text=text
        r._r.append(el)

def format_doc(src,dest):
    doc=Document(str(src)); before=snapshot(doc)
    options_reflowed=reflow_mcq_options(doc)
    removed=remove_graphics(doc) if REMOVE_GRAPHICS else 0
    normal=doc.styles["Normal"]; normal.font.name=FONT; normal.font.size=Pt(BODY)
    style_rpr=normal._element.get_or_add_rPr(); style_fonts=style_rpr.rFonts
    if style_fonts is None: style_fonts=OxmlElement("w:rFonts"); style_rpr.insert(0,style_fonts)
    style_fonts.set(qn("w:ascii"),FONT); style_fonts.set(qn("w:hAnsi"),FONT); style_fonts.set(qn("w:cs"),FONT); style_fonts.set(qn("w:eastAsia"),FONT)
    normal.paragraph_format.space_before=Pt(0); normal.paragraph_format.space_after=Pt(5); normal.paragraph_format.line_spacing=1.12
    for sec in doc.sections:
        sec.page_width=Cm(21); sec.page_height=Cm(29.7)
        sec.top_margin=Cm(2.2); sec.bottom_margin=Cm(2.0); sec.left_margin=Cm(2.3); sec.right_margin=Cm(2.0)
        sec.header_distance=Cm(1); sec.footer_distance=Cm(1)
    ps=all_paragraphs(doc); content=[p for p in ps if p.text.strip()]
    indexes={id(p._p):i for i,p in enumerate(content)}
    counts={k:0 for k in ("opening","section","question","subquestion","metadata")}
    for p in ps:
        t=p.text.strip(); kind=classify(p,indexes.get(id(p._p),999))
        pf=p.paragraph_format; pf.space_before=Pt(0); pf.space_after=Pt(5); pf.line_spacing=1.12; pf.widow_control=True
        if kind=="opening":
            p.alignment=WD_ALIGN_PARAGRAPH.CENTER; pf.space_after=Pt(3); pf.keep_with_next=True
            for r in p.runs: font(r,15 if len(t)<100 else 12,True)
        elif kind=="section":
            pf.space_before=Pt(10); pf.space_after=Pt(5); pf.keep_with_next=True; pf.keep_together=True
            for r in p.runs: font(r,12,True)
        elif kind=="question":
            pf.space_before=Pt(7); pf.space_after=Pt(4); pf.keep_with_next=True; pf.keep_together=True
            for r in p.runs: font(r,BODY,True)
        elif kind=="subquestion":
            pf.left_indent=Cm(.45); pf.first_line_indent=Cm(-.2); pf.keep_together=True
            for r in p.runs: font(r,BODY)
        elif kind=="metadata":
            pf.space_after=Pt(3)
            for r in p.runs: font(r,10,True)
        else:
            for r in p.runs: font(r,BODY)
        if kind in counts: counts[kind]+=1
    for table in doc.tables:
        table.autofit=True
        table.alignment=1
        tblpr=table._tbl.tblPr
        borders=tblpr.first_child_found_in("w:tblBorders")
        if borders is None: borders=OxmlElement("w:tblBorders"); tblpr.append(borders)
        for edge in ("top","left","bottom","right","insideH","insideV"):
            node=borders.find(qn("w:"+edge))
            if node is None: node=OxmlElement("w:"+edge); borders.append(node)
            node.set(qn("w:val"),"single"); node.set(qn("w:sz"),"6"); node.set(qn("w:space"),"0"); node.set(qn("w:color"),"AAB4C0")
        for ri,row in enumerate(table.rows):
            trpr=row._tr.get_or_add_trPr(); trpr.append(OxmlElement("w:cantSplit"))
            if ri==0 and len(table.rows)>1:
                repeat=OxmlElement("w:tblHeader"); repeat.set(qn("w:val"),"true"); trpr.append(repeat)
            for cell in row.cells:
                cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER; cell_margins(cell)
                if ri==0 and len(table.rows)>1:
                    tcpr=cell._tc.get_or_add_tcPr(); shade=OxmlElement("w:shd"); shade.set(qn("w:fill"),"E9EEF4"); tcpr.append(shade)
                for p in cell.paragraphs:
                    p.paragraph_format.space_after=Pt(2); p.paragraph_format.line_spacing=1.0
                    for r in p.runs: font(r,BODY,bold=(ri==0 and len(table.rows)>1) or r.bold)
    for sec in doc.sections:
        p=sec.footer.paragraphs[0]
        if not p.text.strip(): page_number(p)
        for r in p.runs: font(r,9)
    if before["text"]!=snapshot(doc)["text"]: raise RuntimeError("Text changed during formatting")
    dest.parent.mkdir(parents=True,exist_ok=True); doc.save(dest)
    if before["text"]!=snapshot(Document(dest))["text"]: raise RuntimeError("Saved DOCX text differs from source")
    return {"paragraph_count":before["paragraphs"],"nonempty_paragraph_count":before["nonempty"],"text_characters":before["characters"],"mcq_options_reflowed":options_reflowed,"graphics_removed":removed,"layout_counts":counts,"table_count":len(doc.tables)}

def stem(src): return src.stem+"_Incoming" if src.parent==ROOT/"Incoming" else src.stem

def validate_pdf(path):
    if not path.is_file() or not path.stat().st_size: raise RuntimeError("PDF is empty or missing")
    info=command(["pdfinfo",str(path)],"PDF metadata").stdout
    m=re.search(r"^Pages:\s+(\d+)\s*$",info,re.M)
    if not m or int(m.group(1))<1: raise RuntimeError("Invalid PDF page count")
    txt=command(["pdftotext","-layout",str(path),"-"],"PDF text extraction").stdout
    chars=len(re.sub(r"\s+","",txt))
    if chars==0: raise RuntimeError("PDF contains no extractable text")
    return int(m.group(1)),chars

def process(src):
    base=stem(src); outdoc=OUTPUT/f"{base}_Formatted.docx"; outpdf=OUTPUT/f"{base}_Formatted.pdf"
    for f in (outdoc,outpdf):
        if f.exists(): f.unlink()
    with tempfile.TemporaryDirectory(prefix="exam-format-") as tmpstr:
        tmp=Path(tmpstr); converted=convert_doc(src,tmp) if src.suffix.lower()==".doc" else src
        metrics=format_doc(converted,outdoc)
        command(lo("--convert-to","pdf","--outdir",str(OUTPUT),str(outdoc),profile=tmp/"pdf-profile"),f"PDF export: {src.name}")
        pages,pdfchars=validate_pdf(outpdf)
    return {"source":src.relative_to(ROOT).as_posix(),"status":"success","docx":outdoc.relative_to(ROOT).as_posix(),"pdf":outpdf.relative_to(ROOT).as_posix(),"pages":pages,"pdf_extracted_characters":pdfchars,**metrics}

def reports(ok,failed):
    OUTPUT.mkdir(parents=True,exist_ok=True)
    data={"generated_at_utc":datetime.now(timezone.utc).isoformat(),"format_profile":"professional_exam_v3","settings":{"font":FONT,"body_size_pt":BODY,"page":"A4","margins_cm":{"top":2.2,"bottom":2.0,"left":2.3,"right":2.0},"page_numbers":"centered_footer","remove_graphics":REMOVE_GRAPHICS},"processed":ok,"failures":failed,"summary":{"total":len(ok)+len(failed),"successful":len(ok),"failed":len(failed)}}
    (OUTPUT/"processing_report.json").write_text(json.dumps(data,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    lines=["# Exam processing report","",f"- Generated (UTC): {data['generated_at_utc']}","- Profile: Professional exam v3",f"- Successful: {len(ok)}",f"- Failed: {len(failed)}","- A4 layout, semantic question hierarchy, adaptive MCQ option pairing, table spacing and centered page numbers","","| Source | Status | Pages | Questions | Sections | Tables | Graphics removed |","|---|---:|---:|---:|---:|---:|---:|"]
    for x in ok:
        c=x["layout_counts"]; lines.append(f"| `{x['source']}` | OK | {x['pages']} | {c['question']} | {c['section']} | {x['table_count']} | {x['graphics_removed']} |")
        lines.append(f"  - DOCX: `{x['docx']}`"); lines.append(f"  - PDF: `{x['pdf']}`")
    if failed: lines+=["","## Failures",""]+[f"- `{x['source']}`: {x['error']}" for x in failed]
    lines+=["","Text is checked before and after formatting and after saving. Visual review remains necessary before distributing an exam.",""]
    (OUTPUT/"processing_report.md").write_text("\n".join(lines),encoding="utf-8")

def main():
    OUTPUT.mkdir(exist_ok=True); candidates=[]
    for folder in INPUT_DIRS:
        if folder.exists():
            candidates += [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in {".doc",".docx"} and not p.name.startswith("~$") and not p.stem.endswith("_Formatted")]
    ok=[]; failed=[]
    for src in sorted(set(candidates),key=lambda p:p.as_posix().lower()):
        try:
            x=process(src); ok.append(x); print(f"SUCCESS: {x['source']} -> {x['docx']} ({x['pages']} pages)")
        except Exception as e:
            failed.append({"source":src.relative_to(ROOT).as_posix(),"error":str(e)}); print(f"ERROR: {src.name}: {e}",file=sys.stderr)
    reports(ok,failed); print(f"REPORT: {len(ok)} successful, {len(failed)} failed")
    if failed: raise SystemExit(1)
    if not candidates: print("No Word documents found in repository root or Incoming/")
if __name__=="__main__": main()
