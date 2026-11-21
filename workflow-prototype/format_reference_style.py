#!/usr/bin/env python3
"""Apply typography and Word styles sampled from a real reference exam."""
from pathlib import Path
import argparse, re, subprocess, tempfile, copy
from collections import Counter
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "workflow-prototype/output"
OPTION_RE = re.compile(r"^\s*[A-D]\s*[.)]")
QUESTION_RE = re.compile(r"^\s*\d+\s*[.)]")
SECTION_RE = re.compile(r"^(?:Question\s*\(\s*\d+\s*\)|Section\s+[A-Z0-9]+)", re.I)

def role(text):
    t = text.strip()
    if not t: return "empty"
    if OPTION_RE.match(t): return "option"
    if SECTION_RE.match(t) or re.match(r"^(?:Part|Section)\s+[A-Z0-9]+", t, re.I): return "section"
    if QUESTION_RE.match(t): return "question"
    if re.search(r"Program:|Level:|Semester:|Date:|Time:|Name:|Department:|Course:", t, re.I): return "metadata"
    if re.search(r"Main Examination|Final Examination|Mid.?term Examination", t, re.I): return "exam_title"
    if len(t) < 100 and (t.isupper() or t.endswith(":")): return "heading"
    return "body"

def convert_legacy_doc(source, tempdir):
    profile = Path(tempdir) / "lo-profile"
    subprocess.run(["libreoffice", "--headless", f"-env:UserInstallation=file://{profile}",
                    "--convert-to", "docx", "--outdir", tempdir, str(source)],
                   check=True, capture_output=True, text=True)
    converted = Path(tempdir) / (source.stem + ".docx")
    if not converted.exists(): raise RuntimeError("LibreOffice did not produce DOCX from legacy DOC")
    return converted

def all_paragraphs(doc):
    out = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                out.extend(cell.paragraphs)
    return out

def font_signature(run):
    f = run.font
    color = None
    try:
        if f.color and f.color.rgb: color = str(f.color.rgb)
    except Exception: pass
    return {
        "name": f.name, "size": float(f.size.pt) if f.size else None,
        "bold": f.bold, "italic": f.italic, "underline": f.underline,
        "color": color,
    }

def build_profile(reference):
    ref = Document(reference)
    # Copy the reference's actual Word style definitions for matching style IDs,
    # including document defaults, rather than substituting a house-style palette.
    samples = {}
    for p in all_paragraphs(ref):
        r = role(p.text)
        runs = [x for x in p.runs if x.text.strip()]
        if r != "empty" and runs and r not in samples:
            samples[r] = font_signature(runs[0])
    # The reference's Normal style is the fallback for content roles without exemplars.
    normal = ref.styles["Normal"].font
    fallback = {"name": normal.name, "size": float(normal.size.pt) if normal.size else None,
                "bold": normal.bold, "italic": normal.italic, "underline": normal.underline,
                "color": None}
    return ref, samples, fallback

def apply_font(run, spec):
    f = run.font
    if spec.get("name"): f.name = spec["name"]
    if spec.get("size"): f.size = Pt(spec["size"])
    if spec.get("bold") is not None: f.bold = spec["bold"]
    if spec.get("italic") is not None: f.italic = spec["italic"]
    if spec.get("underline") is not None: f.underline = spec["underline"]
    if spec.get("color"):
        f.color.rgb = RGBColor.from_string(spec["color"])
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    if spec.get("name"):
        for key in ("ascii", "hAnsi", "eastAsia", "cs"):
            rfonts.set(qn("w:" + key), spec["name"])

def transfer_styles(reference, target):
    ref_styles = reference.styles.element
    target_styles = target.styles.element
    # Copy document defaults (default paragraph/run formatting) exactly.
    ref_defaults = ref_styles.find(qn("w:docDefaults"))
    if ref_defaults is not None:
        old = target_styles.find(qn("w:docDefaults"))
        if old is not None: target_styles.remove(old)
        target_styles.insert(0, copy.deepcopy(ref_defaults))
    # For style IDs present in both files, use the reference definition.
    ref_by_id = {s.get(qn("w:styleId")): s for s in ref_styles.findall(qn("w:style"))}
    target_by_id = {s.get(qn("w:styleId")): s for s in target_styles.findall(qn("w:style"))}
    for sid, source_style in ref_by_id.items():
        if sid in target_by_id:
            old = target_by_id[sid]
            target_styles.replace(old, copy.deepcopy(source_style))

def remove_forced_page_breaks(doc):
    for p in all_paragraphs(doc):
        p.paragraph_format.page_break_before = False
        for node in list(p._p.xpath(".//w:br[@w:type='page'] | .//w:lastRenderedPageBreak")):
            parent = node.getparent()
            if parent is not None: parent.remove(node)

def content_snapshot(doc):
    return (
        [p.text for p in doc.paragraphs],
        [[cell.text for row in table.rows for cell in row.cells] for table in doc.tables],
        len(doc.inline_shapes),
    )

def format_exam(source, reference_path, output_prefix):
    if not source.exists(): raise SystemExit(f"Missing source exam: {source}")
    if not reference_path.exists(): raise SystemExit(f"Missing reference exam: {reference_path}")
    OUT.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        working_source = convert_legacy_doc(source, tmp) if source.suffix.lower() == ".doc" else source
        doc = Document(working_source)
        before = content_snapshot(doc)
        reference, samples, fallback = build_profile(reference_path)
        transfer_styles(reference, doc)
        remove_forced_page_breaks(doc)

        # Match page geometry from the reference document, but preserve the source's
        # section structure and content. Continuous sections avoid accidental blank pages.
        ref_sec = reference.sections[0]
        for sec in doc.sections:
            sec.page_width = ref_sec.page_width
            sec.page_height = ref_sec.page_height
            sec.top_margin = ref_sec.top_margin
            sec.bottom_margin = ref_sec.bottom_margin
            sec.left_margin = ref_sec.left_margin
            sec.right_margin = ref_sec.right_margin
            sec.header_distance = ref_sec.header_distance
            sec.footer_distance = ref_sec.footer_distance
            sec.start_type = WD_SECTION_START.CONTINUOUS

        for p in all_paragraphs(doc):
            r = role(p.text)
            if r == "empty": continue
            spec = samples.get(r, samples.get("body", fallback))
            for run in p.runs:
                if run.text: apply_font(run, spec)
        # Preserve tables and their existing geometry; transfer reference table style
        # only where a matching named style exists.
        for table in doc.tables:
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            for row in table.rows:
                trpr = row._tr.get_or_add_trPr()
                if trpr.find(qn("w:cantSplit")) is None:
                    trpr.append(OxmlElement("w:cantSplit"))
                for cell in row.cells:
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        after = content_snapshot(doc)
        if before != after:
            raise RuntimeError("Content-preservation check failed: paragraph/table text or inline shape count changed")

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
    print("Reference-derived font profile:", samples)
    print("Content-preservation check: PASSED")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--reference", required=True, help="Reference DOCX whose styles and typography are applied")
    parser.add_argument("--output-prefix", required=True)
    args = parser.parse_args()
    format_exam((ROOT / args.input).resolve(), (ROOT / args.reference).resolve(), args.output_prefix)
