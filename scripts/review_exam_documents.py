#!/usr/bin/env python3
"""Automated PDF quality review for formatted exam documents."""
from __future__ import annotations
import difflib, hashlib, json, os, re, sys
from datetime import datetime, timezone
from pathlib import Path
import fitz
from docx import Document
from PIL import Image, ImageOps, ImageDraw

ROOT = Path(__file__).resolve().parents[1]
FORMATTED = ROOT / "Formatted"
OUT = FORMATTED / "review"
A4 = (595.28, 841.89)

def norm(s):
    return re.sub(r"[^a-z0-9]+", "", s.casefold())

def get_docx_text(path):
    d = Document(path)
    chunks = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            for cell in row.cells:
                chunks.extend(p.text for p in cell.paragraphs)
    return "\n".join(chunks)

def review(path):
    errors, warnings, pages, texts, images = [], [], [], [], []
    doc = fitz.open(path)
    if doc.is_encrypted:
        errors.append("PDF is encrypted.")
    if len(doc) == 0:
        errors.append("PDF has no pages.")
    render_dir = OUT / path.stem
    render_dir.mkdir(parents=True, exist_ok=True)
    hashes = {}
    for i, page in enumerate(doc):
        n = i + 1
        text = page.get_text("text", sort=True)
        compact = norm(text)
        texts.append(text)
        pe, pw = [], []
        if not compact:
            pe.append("No extractable text (blank or image-only page).")
        elif len(compact) < 35:
            pw.append(f"Low text density: {len(compact)} characters.")
        if abs(page.rect.width-A4[0]) > 8 or abs(page.rect.height-A4[1]) > 8:
            pw.append(f"Page is not A4: {page.rect.width:.1f} x {page.rect.height:.1f} pt.")
        if compact:
            digest = hashlib.sha256(compact.encode()).hexdigest()
            if digest in hashes:
                pw.append(f"Page text duplicates page {hashes[digest]}.")
            else:
                hashes[digest] = n
        minfont, edge = None, False
        for block in page.get_text("dict").get("blocks", []):
            if block.get("type") != 0:
                continue
            b = block.get("bbox", [0,0,0,0])
            if b[0] < 18 or b[1] < 18 or page.rect.width-b[2] < 18 or page.rect.height-b[3] < 18:
                edge = True
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    size = span.get("size")
                    if size and (minfont is None or size < minfont):
                        minfont = size
        if minfont is not None and minfont < 7:
            pw.append(f"Small text detected: minimum {minfont:.1f} pt.")
        if edge:
            pw.append("Text is close to a page edge; inspect for clipping.")
        imgpath = render_dir / f"page-{n:03d}.png"
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5,1.5), alpha=False)
        pix.save(imgpath)
        images.append(Image.open(imgpath).convert("RGB"))
        pages.append({"page":n,"characters":len(compact),"smallest_font_pt":round(minfont,2) if minfont else None,
                      "errors":pe,"warnings":pw,"preview":imgpath.relative_to(ROOT).as_posix()})
        errors.extend([{"page":n,"message":x} for x in pe])
        warnings.extend([{"page":n,"message":x} for x in pw])
    similarity = None
    docx = path.with_suffix(".docx")
    if docx.exists():
        a, b = norm(get_docx_text(docx)), norm("\n".join(texts))
        similarity = difflib.SequenceMatcher(None,a,b,autojunk=False).ratio()
        if similarity < .82:
            warnings.append({"message":f"DOCX/PDF text similarity is {similarity:.1%}; inspect content retention."})
    else:
        warnings.append({"message":"Matching DOCX not found; text comparison skipped."})
    if images:
        tw = 340
        thumbs = [ImageOps.expand(im.resize((tw,int(im.height*tw/im.width))),border=2,fill="white") for im in images]
        cols = 2 if len(thumbs)>1 else 1
        cw, ch = max(x.width for x in thumbs)+16, max(x.height for x in thumbs)+34
        sheet = Image.new("RGB",(cols*cw,((len(thumbs)+cols-1)//cols)*ch),"#dce1e7")
        draw = ImageDraw.Draw(sheet)
        for i,im in enumerate(thumbs):
            x=(i%cols)*cw+8; y=(i//cols)*ch+25
            draw.text((x,y-18),f"Page {i+1}",fill="#17202a"); sheet.paste(im,(x,y))
        sheet.save(OUT/f"{path.stem}-contact-sheet.jpg",quality=88)
    return {"pdf":path.relative_to(ROOT).as_posix(),"status":"FAIL" if errors else ("PASS WITH WARNINGS" if warnings else "PASS"),
            "pages":len(doc),"bytes":path.stat().st_size,"text_similarity":round(similarity,4) if similarity is not None else None,
            "page_details":pages,"errors":errors,"warnings":warnings,
            "contact_sheet":f"{path.stem}-contact-sheet.jpg"}

def main():
    OUT.mkdir(parents=True,exist_ok=True)
    pdfs=sorted(FORMATTED.glob("*_Formatted.pdf"))
    if not pdfs:
        print("ERROR: no Formatted/*_Formatted.pdf files found.",file=sys.stderr); return 2
    results=[review(p) for p in pdfs]
    summary={"files":len(results),"passed":sum(x["status"]=="PASS" for x in results),
             "warnings":sum(x["status"]=="PASS WITH WARNINGS" for x in results),
             "failed":sum(x["status"]=="FAIL" for x in results)}
    report={"generated_utc":datetime.now(timezone.utc).isoformat(),"engine":"Mydocs Exam PDF QA v1",
            "checks":["PDF opens","extractable text","blank pages","duplicate page text","A4 geometry",
                      "small fonts","text near edges","DOCX/PDF text similarity","render every page"],
            "summary":summary,"results":results}
    (OUT/"review_report.json").write_text(json.dumps(report,indent=2),encoding="utf-8")
    lines=["# Automated Exam Document Review","",f"Generated UTC: {report['generated_utc']}","",
           "| PDF | Status | Pages | Text similarity | Errors | Warnings |","|---|---|---:|---:|---:|---:|"]
    for x in results:
        sim=f"{x['text_similarity']:.1%}" if x["text_similarity"] is not None else "N/A"
        lines.append(f"| {x['pdf']} | **{x['status']}** | {x['pages']} | {sim} | {len(x['errors'])} | {len(x['warnings'])} |")
        for e in x["errors"]: lines.append(f"- ERROR page {e.get('page','')}: {e['message']}")
        for w in x["warnings"]: lines.append(f"- WARNING page {w.get('page','')}: {w['message']}")
        lines.append(f"- Contact sheet: {x['contact_sheet']}")
    lines += ["","Automated checks flag likely layout risks; visual previews are provided for human sign-off. "
              "This is not a claim of human-level semantic or typography review.",""]
    md="\n".join(lines)
    (OUT/"review_report.md").write_text(md,encoding="utf-8")
    (OUT/"index.html").write_text("<meta charset='utf-8'><title>Exam review</title><pre style='white-space:pre-wrap;font:15px system-ui'>"+
                                  __import__("html").escape(md)+"</pre>",encoding="utf-8")
    summary_path=os.environ.get("GITHUB_STEP_SUMMARY")
    if summary_path:
        with open(summary_path,"a",encoding="utf-8") as f: f.write(md)
    print(json.dumps(summary,indent=2))
    return 1 if summary["failed"] else 0

if __name__=="__main__":
    raise SystemExit(main())
