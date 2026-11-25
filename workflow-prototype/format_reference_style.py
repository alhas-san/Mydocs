#!/usr/bin/env python3
"""Apply typography and Word styles sampled from a real reference exam."""
from pathlib import Path
import argparse, re, subprocess, tempfile, copy
from collections import Counter
import zipfile, os
from lxml import etree
from docx import Document
from docx.text.paragraph import Paragraph
from docx.shared import Cm, Pt, Inches, RGBColor
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "workflow-prototype/output"
OPTION_RE = re.compile(r"^\s*(?:[A-D]|[أابجدهو])\s*[.)、:：-]")
QUESTION_RE = re.compile(r"^\s*[0-9٠-٩]+\s*[.)、:：-]")
ARABIC_RE = re.compile(r"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF]")
ARABIC_QUESTION_RE = re.compile(r"^\s*(?:السؤال|سؤال)\s*(?:رقم\s*)?(?:[0-9٠-٩]+|الأول|الثاني|الثالث|الرابع|الخامس|السادس|السابع|الثامن|التاسع|العاشر)?", re.I)
ARABIC_SECTION_RE = re.compile(r"^\s*(?:القسم|الجزء|أجب عن|أجيبي عن|اختر|ضع علامة)", re.I)
SECTION_RE = re.compile(r"^(?:Question\s*\(?\s*\d+\s*\)?|Section\s+[A-Z0-9]+)", re.I)

def role(text):
    t = text.strip()
    if not t: return "empty"
    if OPTION_RE.match(t): return "option"
    if ARABIC_QUESTION_RE.match(t) and ARABIC_RE.search(t): return "question"
    if ARABIC_SECTION_RE.match(t) and ARABIC_RE.search(t): return "section"
    if (SECTION_RE.match(t)
            or re.match(r"^Question\s+(?:One|Two|Three|Four|Five|Six|Seven|Eight|Nine|Ten)\b", t, re.I)
            or re.match(r"^(?:Part|Section)\s+[A-Z0-9]+", t, re.I)):
        return "section"
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

def option_chunks(paragraph):
    """Split only a clean, complete inline A-D option sequence.

    Requiring the paragraph to begin with A and contain exactly A, B, C, D
    prevents ordinary prose (or option-like text inside a question stem) from
    being mistaken for an option list and silently discarded during reflow.
    """
    text = paragraph.text
    matches = list(re.finditer(r"(?<![A-Za-z0-9])([A-D])\s*[.)]\s*", text))
    if text[:matches[0].start()].strip() if matches else True:
        return []
    if [m.group(1) for m in matches] != list("ABCD"):
        return []
    chunks = []
    for i, match in enumerate(matches):
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        label = match.group(1) + match.group(0)[len(match.group(1)):].strip()
        body = text[match.end():end].strip()
        if not body:
            return []
        chunks.append((label, body))
    return chunks

def reflow_inline_options(doc):
    """Lay out compact options in two columns; give long options a full-width row.

    Fixed tab stops replace the source's irregular run-level spacing. The paragraph
    count and the actual option wording remain unchanged.
    """
    changed = 0
    for p in all_paragraphs(doc):
        chunks = option_chunks(p)
        if not chunks:
            continue
        # Two options share a row only when each is short enough to fit in its
        # half-page column without compressing characters or colliding.
        if all(len(body) <= 22 for _, body in chunks):
            rows = [(chunks[i], chunks[i + 1]) for i in range(0, len(chunks) - 1, 2)]
            if len(chunks) % 2:
                rows.append((chunks[-1],))
        else:
            # If any option is too long for half-width, keep every option in
            # this question on its own full-width line for consistent alignment.
            rows = [(item,) for item in chunks]

        # Do not rebuild paragraphs that contain drawings, fields, hyperlinks, or
        # other non-run content: rebuilding those would silently discard objects.
        allowed_children = {qn("w:pPr"), qn("w:r")}
        if any(child.tag not in allowed_children for child in p._p):
            continue
        allowed_run_children = {qn("w:rPr"), qn("w:t"), qn("w:tab"), qn("w:br"), qn("w:cr")}
        if any(node.tag not in allowed_run_children
               for run in p._p.findall(qn("w:r")) for node in run):
            continue

        # Retain the paragraph's paragraph-level properties, but rebuild its runs
        # so inherited negative character spacing and old tabs cannot squeeze text.
        ppr = p._p.pPr
        for child in list(p._p):
            if child is not ppr:
                p._p.remove(child)
        fmt = p.paragraph_format
        fmt.left_indent = Cm(0)
        fmt.first_line_indent = Cm(0)
        fmt.right_indent = Cm(0)
        fmt.space_before = Pt(0)
        fmt.space_after = Pt(2)
        fmt.line_spacing = 1.0
        tabs = fmt.tab_stops
        tabs.clear_all()
        tabs.add_tab_stop(Cm(8.0))
        for row_index, row in enumerate(rows):
            if row_index:
                br = OxmlElement("w:r")
                brpr = OxmlElement("w:rPr")
                br.append(brpr)
                br.append(OxmlElement("w:br"))
                p._p.append(br)
            for col_index, (label, body) in enumerate(row):
                if col_index:
                    tab_run = OxmlElement("w:r")
                    tab_run.append(OxmlElement("w:tab"))
                    p._p.append(tab_run)
                run = p.add_run(label + " " + body)
                # Prevent legacy negative character-spacing values from narrowing
                # the text. Use the reference-derived option font, with normal tracking.
                rpr = run._r.get_or_add_rPr()
                spacing = rpr.find(qn("w:spacing"))
                if spacing is not None:
                    rpr.remove(spacing)
        changed += 1
    return changed

def apply_option_spacing(doc):
    """Keep question groups intact and normalize True/False answer blanks."""
    changed = 0
    paragraphs = all_paragraphs(doc)
    for index, p in enumerate(paragraphs):
        inside_table_cell = p._p.getparent() is not None and p._p.getparent().tag == qn("w:tc")
        paragraph_role = role(p.text)
        next_role = role(paragraphs[index + 1].text) if index + 1 < len(paragraphs) else "empty"
        has_inline_image = bool(p._p.xpath(".//w:drawing | .//w:pict"))
        next_has_inline_image = (
            index + 1 < len(paragraphs)
            and bool(paragraphs[index + 1]._p.xpath(".//w:drawing | .//w:pict"))
        )

        # Keep section headings, numbered stems, and each option row attached.
        # Chaining keep_with_next across A/B/C/D keeps the entire option set
        # with its question rather than allowing the last options to spill.
        if not inside_table_cell:
            # Keep section headings attached to their instruction, and question
            # stems attached to answer options. Do not chain consecutive numbered
            # questions (e.g. True/False lists) into one page-sized block.
            if paragraph_role == "section" or (p.text.strip() and next_role == "option"):
                p.paragraph_format.keep_with_next = True
            if paragraph_role == "option" and next_role == "option":
                p.paragraph_format.keep_with_next = True
        if p.text.strip() and next_has_inline_image:
            p.paragraph_format.keep_with_next = True
        if has_inline_image:
            p.paragraph_format.keep_together = True
        if paragraph_role in {"question", "option", "section"}:
            p.paragraph_format.keep_together = True

        if paragraph_role == "option":
            fmt = p.paragraph_format
            fmt.space_before = Pt(1.5)
            fmt.space_after = Pt(5)
            fmt.line_spacing = 1.12
            fmt.widow_control = True
            changed += 1

        # Replace variable-width tab stops inside True/False parentheses with
        # exactly four non-breaking spaces, so Word/LibreOffice render equal blanks.
        runs = p.runs
        for run_index, run in enumerate(runs):
            if not run.text:
                continue
            if re.search(r"\([ \t\u00a0]*\)", run.text):
                run.text = re.sub(r"\([ \t\u00a0]*\)", "(    )", run.text)
            elif run.text == "\t" and run_index > 0 and run_index + 1 < len(runs):
                if runs[run_index - 1].text.strip().endswith("(") and runs[run_index + 1].text.strip().startswith(")"):
                    run.text = "    "
    return changed



def apply_arabic_layout(doc):
    """Format Arabic text without forcing English-led mixed paragraphs into RTL layout."""
    arabic_paragraphs = 0
    arabic_questions = 0
    arabic_options = 0
    rtl_paragraphs = 0
    mixed_paragraphs = 0
    paragraphs = all_paragraphs(doc)
    for paragraph_index, p in enumerate(paragraphs):
        paragraph_text = p.text or ""
        next_role = role(paragraphs[paragraph_index + 1].text) if paragraph_index + 1 < len(paragraphs) else "empty"
        if not ARABIC_RE.search(paragraph_text):
            continue
        arabic_paragraphs += 1
        paragraph_role = role(paragraph_text)
        # Numbered statements inside table cells are table rows, not standalone
        # question headings. Treat them as body text so they do not chain rows.
        if p._p.getparent() is not None and p._p.getparent().tag == qn("w:tc") and paragraph_role == "question":
            paragraph_role = "body"
        arabic_count = len(ARABIC_RE.findall(paragraph_text))
        latin_count = len(re.findall(r"[A-Za-z]", paragraph_text))
        arabic_dominant = arabic_count >= latin_count
        fmt = p.paragraph_format

        ppr = p._p.get_or_add_pPr()
        bidi = ppr.find(qn("w:bidi"))
        if arabic_dominant:
            rtl_paragraphs += 1
            fmt.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            fmt.left_indent = Cm(0)
            fmt.right_indent = Cm(0)
            fmt.first_line_indent = Cm(0)
            fmt.line_spacing = 1.25
            fmt.widow_control = True
            if bidi is None:
                bidi = OxmlElement("w:bidi")
                ppr.append(bidi)
            bidi.set(qn("w:val"), "1")
        else:
            mixed_paragraphs += 1
            # Keep English-led bilingual lines in their original LTR layout.
            fmt.alignment = WD_ALIGN_PARAGRAPH.LEFT
            if bidi is not None:
                ppr.remove(bidi)

        if arabic_dominant:
            if paragraph_role == "question":
                fmt.space_before = Pt(9)
                fmt.space_after = Pt(5)
                fmt.keep_with_next = (next_role == "option")
                fmt.keep_together = True
                arabic_questions += 1
            elif paragraph_role == "option":
                fmt.space_before = Pt(2)
                fmt.space_after = Pt(5)
                fmt.keep_together = True
                arabic_options += 1
            elif paragraph_role == "section":
                fmt.space_before = Pt(10)
                fmt.space_after = Pt(7)
                fmt.keep_with_next = True
            else:
                fmt.space_before = Pt(3)
                fmt.space_after = Pt(6)

        for run in p.runs:
            if not run.text or not ARABIC_RE.search(run.text):
                continue
            rpr = run._r.get_or_add_rPr()
            rfonts = rpr.find(qn("w:rFonts"))
            if rfonts is None:
                rfonts = OxmlElement("w:rFonts")
                rpr.insert(0, rfonts)

            run_has_latin = bool(re.search(r"[A-Za-z]", run.text))
            if run_has_latin:
                # Preserve Latin font and size; set only the complex-script font slot.
                rfonts.set(qn("w:cs"), "Noto Naskh Arabic")
                sz_cs = rpr.find(qn("w:szCs"))
                if sz_cs is None:
                    sz_cs = OxmlElement("w:szCs")
                    rpr.append(sz_cs)
                sz_cs.set(qn("w:val"), "30")
            else:
                run.font.name = "Noto Naskh Arabic"
                run.font.size = Pt(16 if paragraph_role in {"question", "section"} else 15)
                for key in ("ascii", "hAnsi", "eastAsia", "cs"):
                    rfonts.set(qn("w:" + key), "Noto Naskh Arabic")

            rtl = rpr.find(qn("w:rtl"))
            if rtl is None:
                rtl = OxmlElement("w:rtl")
                rpr.append(rtl)
            rtl.set(qn("w:val"), "1")
            if paragraph_role == "section" and not run_has_latin:
                run.bold = True

    print(
        "Arabic layout: "
        f"{arabic_paragraphs} Arabic paragraphs; {rtl_paragraphs} RTL-aligned, "
        f"{mixed_paragraphs} English-led mixed; {arabic_questions} questions, "
        f"{arabic_options} answer options; Noto Naskh Arabic"
    )
    return arabic_paragraphs, arabic_questions, arabic_options



def content_snapshot(doc):
    normalize = lambda value: re.sub(r"\s+", " ", value).strip()
    return (
        normalize(" ".join(p.text for p in doc.paragraphs)),
        [[normalize(cell.text) for row in table.rows for cell in row.cells] for table in doc.tables],
        len(doc.inline_shapes),
    )



def _w(tag):
    return "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}" + tag

def _make_dropdown(tag, placeholder, choices, control_id, run_properties=None):
    """Create an editable Word dropdown content control."""
    sdt = etree.Element(_w("sdt"))
    props = etree.SubElement(sdt, _w("sdtPr"))
    alias = etree.SubElement(props, _w("alias")); alias.set(_w("val"), placeholder)
    tag_el = etree.SubElement(props, _w("tag")); tag_el.set(_w("val"), tag)
    id_el = etree.SubElement(props, _w("id")); id_el.set(_w("val"), str(control_id))
    dropdown = etree.SubElement(props, _w("dropDownList"))
    dropdown.set(_w("lastValue"), placeholder)
    etree.SubElement(dropdown, _w("listItem"), {_w("displayText"): placeholder, _w("value"): placeholder})
    for choice in choices:
        etree.SubElement(dropdown, _w("listItem"), {_w("displayText"): choice, _w("value"): choice})
    content = etree.SubElement(sdt, _w("sdtContent"))
    run = etree.SubElement(content, _w("r"))
    if run_properties is not None:
        run.append(copy.deepcopy(run_properties))
    text = etree.SubElement(run, _w("t"))
    text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text.text = placeholder
    return sdt

def _paragraph_text(p):
    return "".join(p.xpath(".//w:t/text()", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}))

def _replace_run_with_dropdown(p, exact_text, dropdown):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    for run in p.xpath("./w:r", namespaces=ns):
        values = run.xpath("./w:t/text()", namespaces=ns)
        if len(values) == 1 and values[0] == exact_text:
            parent = run.getparent()
            index = parent.index(run)
            parent.remove(run)
            parent.insert(index, dropdown)
            return True
    return False

def _insert_after_text_run(p, exact_text, dropdown):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    for run in p.xpath("./w:r", namespaces=ns):
        if "".join(run.xpath("./w:t/text()", namespaces=ns)) == exact_text:
            parent = run.getparent()
            index = parent.index(run)
            parent.insert(index + 1, dropdown)
            return True
    return False

# The dropdowns are editable Word content controls embedded in the first-page header.
def apply_reference_header(reference_path, output_docx, exam_title):
    """Copy the Analytical Chemistry first-page header and turn its academic fields into Word dropdowns."""
    with zipfile.ZipFile(reference_path, "r") as ref_zip:
        names = set(ref_zip.namelist())
        header_name = "word/header2.xml"
        rels_name = "word/_rels/header2.xml.rels"
        image_name = "word/media/image2.png"
        if not {header_name, rels_name, image_name}.issubset(names):
            raise RuntimeError("Reference exam is missing its first-page header template or logo asset")

        parser = etree.XMLParser(remove_blank_text=False)
        header = etree.fromstring(ref_zip.read(header_name), parser)
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        paragraphs = header.xpath(".//w:p", namespaces=ns)

        program_p = next((p for p in paragraphs if "Program:" in _paragraph_text(p)), None)
        level_semester_p = next((p for p in paragraphs if "Level:" in _paragraph_text(p) and "Semester:" in _paragraph_text(p)), None)
        title_p = next((p for p in paragraphs if "Main Examination:-" in _paragraph_text(p)), None)
        date_p = next((p for p in paragraphs if "Date:" in _paragraph_text(p) and "Time:" in _paragraph_text(p)), None)
        if not all((program_p, level_semester_p, title_p, date_p)):
            raise RuntimeError("Could not locate all expected metadata lines in the reference header")

        program_choices = [
            "Bachelor in Biomedical Engineering",
            "Diploma in Biomedical Engineering",
            "Bachelor in Information Technology",
            "Diploma in Information Technology",
            "Bachelor in Business Administration",
            "Bachelor in Nursing Science",
            "Bachelor in Medical Laboratory Science",
            "Bachelor in Pharmacy",
            "Bachelor in Medicine and Surgery",
        ]
        level_choices = [f"{n}{'st' if n == 1 else 'nd' if n == 2 else 'rd' if n == 3 else 'th'}" for n in range(1, 6)]
        semester_choices = [f"{n}{'st' if n == 1 else 'nd' if n == 2 else 'rd' if n == 3 else 'th'}" for n in range(1, 11)]

        # Reuse the reference's run formatting so the controls sit naturally in the original design.
        program_run = next((r for r in program_p.xpath("./w:r", namespaces=ns) if "(Bachelor" in "".join(r.xpath("./w:t/text()", namespaces=ns))), None)
        level_run = next((r for r in level_semester_p.xpath("./w:r", namespaces=ns) if "".join(r.xpath("./w:t/text()", namespaces=ns)).strip() == "2nd"), None)
        semester_run = next((r for r in level_semester_p.xpath("./w:r", namespaces=ns) if "".join(r.xpath("./w:t/text()", namespaces=ns)).strip() == "3rd"), None)
        if program_run is None or level_run is None or semester_run is None:
            raise RuntimeError("Could not locate program, level, and semester values in the reference header")

        # Remove the fixed program value and its surrounding parentheses, then insert a dropdown.
        program_parent = program_run.getparent()
        program_index = program_parent.index(program_run)
        for node in list(program_parent)[program_index:]:
            if node.tag == _w("r"):
                value = "".join(node.xpath("./w:t/text()", namespaces=ns))
                if value in {"(Bachelor’s in PHARMAC", "Y", ")"}:
                    program_parent.remove(node)
        program_parent.insert(program_index, _make_dropdown("academic-program", "Select Program", program_choices, 71001, program_run.find(_w("rPr"))))

        level_run_properties = level_run.find(_w("rPr"))
        _replace_run_with_dropdown(level_semester_p, "2nd", _make_dropdown("academic-level", "Select", level_choices, 71002, level_run_properties))
        semester_run_properties = semester_run.find(_w("rPr"))
        _replace_run_with_dropdown(level_semester_p, "3rd", _make_dropdown("academic-semester", "Select", semester_choices, 71003, semester_run_properties))

        # The exam title is per-document; do not carry the reference subject into another exam.
        for run in title_p.xpath("./w:r", namespaces=ns):
            for t in run.xpath("./w:t", namespaces=ns):
                if t.text and "Analytical Chemistry I" in t.text:
                    t.text = t.text.replace("Analytical Chemistry I", exam_title)

        # Remove the reference exam's date and duration to avoid publishing stale metadata.
        for run in date_p.xpath("./w:r", namespaces=ns):
            for t in run.xpath("./w:t", namespaces=ns):
                if t.text == "21":
                    t.text = "____________"
                elif t.text and t.text.startswith("/11/2026"):
                    t.text = "          Time:"
                elif t.text == "3":
                    t.text = "____ "
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        header_bytes = etree.tostring(header, xml_declaration=True, encoding="UTF-8", standalone=True)
        header_rels = ref_zip.read(rels_name).replace(b"media/image2.png", b"media/reference_header_logo.png")
        logo = ref_zip.read(image_name)

    temp_output = str(output_docx) + ".tmp"
    with zipfile.ZipFile(output_docx, "r") as src, zipfile.ZipFile(temp_output, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            if item.filename in {"word/header_reference.xml", "word/_rels/header_reference.xml.rels", "word/media/reference_header_logo.png"}:
                continue
            data = src.read(item.filename)
            if item.filename == "word/document.xml":
                root = etree.fromstring(data, parser)
                rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                for sect in root.xpath(".//w:sectPr", namespaces=ns):
                    for old in sect.xpath("./w:headerReference[@w:type='first']", namespaces=ns):
                        sect.remove(old)
                    ref = etree.Element(_w("headerReference"))
                    ref.set(_w("type"), "first")
                    ref.set("{" + rel_ns + "}id", "rIdHeaderReference")
                    sect.insert(0, ref)
                    if not sect.xpath("./w:titlePg", namespaces=ns):
                        sect.append(etree.Element(_w("titlePg")))
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            elif item.filename == "word/_rels/document.xml.rels":
                root = etree.fromstring(data, parser)
                rel_root = "http://schemas.openxmlformats.org/package/2006/relationships"
                for rel in list(root):
                    if rel.get("Id") == "rIdHeaderReference":
                        root.remove(rel)
                rel = etree.SubElement(root, "{" + rel_root + "}Relationship")
                rel.set("Id", "rIdHeaderReference")
                rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header")
                rel.set("Target", "header_reference.xml")
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            elif item.filename == "[Content_Types].xml":
                root = etree.fromstring(data, parser)
                ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
                if not root.xpath("./*[local-name()='Override' and @PartName='/word/header_reference.xml']"):
                    override = etree.SubElement(root, "{" + ct_ns + "}Override")
                    override.set("PartName", "/word/header_reference.xml")
                    override.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")
                if not root.xpath("./*[local-name()='Default' and @Extension='png']"):
                    default = etree.SubElement(root, "{" + ct_ns + "}Default")
                    default.set("Extension", "png")
                    default.set("ContentType", "image/png")
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            dst.writestr(item, data)
        dst.writestr("word/header_reference.xml", header_bytes)
        dst.writestr("word/_rels/header_reference.xml.rels", header_rels)
        dst.writestr("word/media/reference_header_logo.png", logo)
    os.replace(temp_output, output_docx)



def apply_page_number_footer(doc):
    """Add a compact centered PAGE field to every page, including first pages."""
    def write_page_field(footer):
        footer.is_linked_to_previous = False
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        for child in list(p._p):
            if child.tag != qn("w:pPr"):
                p._p.remove(child)
        p.alignment = 1
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run()
        run.font.name = "Times New Roman"
        run.font.size = Pt(8)
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        instr = OxmlElement("w:instrText")
        instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        instr.text = " PAGE "
        separate = OxmlElement("w:fldChar")
        separate.set(qn("w:fldCharType"), "separate")
        displayed = OxmlElement("w:t")
        displayed.text = "1"
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        for element in (begin, instr, separate, displayed, end):
            field_run = OxmlElement("w:r")
            rpr = OxmlElement("w:rPr")
            fonts = OxmlElement("w:rFonts")
            fonts.set(qn("w:ascii"), "Times New Roman")
            fonts.set(qn("w:hAnsi"), "Times New Roman")
            rpr.append(fonts)
            size = OxmlElement("w:sz")
            size.set(qn("w:val"), "16")
            rpr.append(size)
            field_run.append(rpr)
            field_run.append(element)
            p._p.append(field_run)

    for section in doc.sections:
        section.footer_distance = Inches(0.20)
        section.different_first_page_header_footer = True
        write_page_field(section.footer)
        write_page_field(section.first_page_footer)


def format_exam(source, reference_path, output_prefix, exam_title=None):
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
        option_paragraphs = reflow_inline_options(doc)

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
        spaced_option_paragraphs = apply_option_spacing(doc)
        arabic_layout_stats = apply_arabic_layout(doc)
        # Preserve tables and their existing geometry; transfer reference table style
        # only where a matching named style exists.
        for table in doc.tables:
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            for row_index, row in enumerate(table.rows):
                trpr = row._tr.get_or_add_trPr()
                if trpr.find(qn("w:cantSplit")) is None:
                    trpr.append(OxmlElement("w:cantSplit"))
                # Repeat the first row when a table continues on another page.
                # Do not apply keep-with-next inside this row: LibreOffice can
                # interpret that as keeping the entire table together, pushing
                # it wholesale onto the next page and leaving a large blank area.
                if row_index == 0 and trpr.find(qn("w:tblHeader")) is None:
                    header_flag = OxmlElement("w:tblHeader")
                    header_flag.set(qn("w:val"), "true")
                    trpr.append(header_flag)
                for cell in row.cells:
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        # Pagination policy:
        # 1) Never chain a whole table to the paragraph before it. LibreOffice may
        #    treat that chain as one indivisible block and push a long table to the
        #    next page, leaving most of the previous page empty.
        # 2) Keep a question heading attached to its own instruction paragraph(s),
        #    but end the chain before the table. The table can then flow naturally
        #    between pages while individual rows remain indivisible (cantSplit)
        #    and its first row repeats as a header (tblHeader).
        # 3) Keep each question/answer paragraph together; keep question stems
        #    attached to the first answer paragraph via apply_option_spacing().
        body_children = list(doc.element.body)
        for index, node in enumerate(body_children):
            if node.tag != qn("w:p"):
                continue
            paragraph = Paragraph(node, doc)
            if role(paragraph.text) != "section":
                # Numbered question paragraphs are kept with their options by
                # apply_option_spacing(); never chain one question to the next.
                continue
            next_index = index + 1
            while next_index < len(body_children):
                next_node = body_children[next_index]
                # Stop at a table: do not attach a complete table to its heading.
                if next_node.tag == qn("w:tbl"):
                    break
                if next_node.tag != qn("w:p"):
                    break
                next_paragraph = Paragraph(next_node, doc)
                if not next_paragraph.text.strip():
                    break
                next_role = role(next_paragraph.text)
                # Attach only the section's instruction block and, at most, its
                # first numbered question. Further questions must paginate freely.
                paragraph.paragraph_format.keep_with_next = True
                paragraph = next_paragraph
                next_index += 1
                if next_role == "question":
                    break
            paragraph.paragraph_format.keep_together = True
            # Keep the last instruction/first question only with what follows
            # when that next paragraph is an answer option.
            paragraph.paragraph_format.keep_with_next = (role(paragraph.text) == "question" and
                next_index < len(body_children) and body_children[next_index].tag == qn("w:p") and
                role(Paragraph(body_children[next_index], doc).text) == "option")

        # Attach a nearby question heading/instructions to the first row of a
        # following table. Ignore intervening blank spacer paragraphs, but only
        # search a short distance (max 5 non-empty paragraphs) so a later answer
        # grid is not accidentally attached to every question in a section.
        body_children = list(doc.element.body)
        for table in doc.tables:
            table_index = body_children.index(table._tbl)
            candidate_nodes = []
            nonempty_count = 0
            found_heading = False
            cursor = table_index - 1
            while cursor >= 0:
                node = body_children[cursor]
                if node.tag != qn("w:p"):
                    break
                paragraph = Paragraph(node, doc)
                candidate_nodes.append(node)
                if paragraph.text.strip():
                    nonempty_count += 1
                    if role(paragraph.text) in {"question", "section"}:
                        found_heading = True
                        break
                    if nonempty_count >= 5:
                        break
                cursor -= 1
            if found_heading:
                # Link only the short question block and the first table row.
                # Table rows themselves are not chained, so long tables still
                # split cleanly between complete rows.
                for node in reversed(candidate_nodes):
                    Paragraph(node, doc).paragraph_format.keep_with_next = True

        after = content_snapshot(doc)
        if before != after:
            raise RuntimeError("Content-preservation check failed: paragraph/table text or inline shape count changed")

        apply_page_number_footer(doc)
        output_docx = OUT / f"{output_prefix}.docx"
        doc.save(output_docx)
        inferred_title = exam_title or re.sub(r"\b(?:final|midterm|main)?\s*exam(?:ination)?\b", "", re.sub(r"[_-]+", " ", source.stem), flags=re.I).strip().title()
        apply_reference_header(reference_path, output_docx, inferred_title)
        profile = Path(tmp) / "pdf-profile"
        subprocess.run(["libreoffice", "--headless", f"-env:UserInstallation=file://{profile}",
                        "--convert-to", "pdf", "--outdir", str(OUT), str(output_docx)],
                       check=True, capture_output=True, text=True)
    pdf = OUT / f"{output_prefix}.pdf"
    info = subprocess.run(["pdfinfo", str(pdf)], check=True, capture_output=True, text=True).stdout
    pages = re.search(r"^Pages:\s+(\d+)", info, re.M)
    print(f"Created {output_docx.name} and {pdf.name}; pages={pages.group(1) if pages else 'unknown'}")
    print("Reference-derived font profile:", samples)
    print(f"Option layout: reflowed {option_paragraphs} inline option paragraphs")
    print(f"Option spacing: applied readable paragraph spacing to {spaced_option_paragraphs} option paragraphs")
    print("Content-preservation check: PASSED")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--reference", required=True, help="Reference DOCX whose styles and typography are applied")
    parser.add_argument("--output-prefix", required=True)
    parser.add_argument("--exam-title", default=None, help="Exam title printed in the first-page header")
    args = parser.parse_args()
    format_exam((ROOT / args.input).resolve(), (ROOT / args.reference).resolve(), args.output_prefix, args.exam_title)
