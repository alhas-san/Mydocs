#!/usr/bin/env python3
"""Format exam documents, preserve their text, remove embedded graphics, and validate outputs."""

from __future__ import annotations

import json
import re
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "Formatted"
INPUT_DIRS = (ROOT, ROOT / "Incoming")
REMOVE_GRAPHICS = True
EXPORT_PDF = True
FONT_NAME = "Times New Roman"
FONT_SIZE = 12
GRAPHIC_TAGS = {qn("w:drawing"), qn("w:pict"), qn("w:object")}


def run_command(command: list[str], *, description: str) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(command, capture_output=True, text=True)
    if result.returncode:
        detail = (result.stderr or result.stdout or "No diagnostic output").strip()
        raise RuntimeError(f"{description} failed: {detail}")
    return result


def libreoffice_command(*args: str, profile_dir: Path) -> list[str]:
    profile_uri = profile_dir.resolve().as_uri()
    return ["libreoffice", "--headless", f"-env:UserInstallation={profile_uri}", *args]


def convert_legacy_doc(source: Path, destination_dir: Path, profile_dir: Path) -> Path:
    run_command(
        libreoffice_command(
            "--convert-to", "docx", "--outdir", str(destination_dir), str(source),
            profile_dir=profile_dir,
        ),
        description=f"Legacy DOC conversion for {source.name}",
    )
    converted = destination_dir / f"{source.stem}.docx"
    if not converted.is_file() or converted.stat().st_size == 0:
        raise RuntimeError(f"LibreOffice did not create a valid DOCX for {source.name}")
    return converted


def iter_paragraphs(container) -> Iterable:
    """Yield paragraphs in a document/container, including paragraphs in nested tables."""
    for paragraph in getattr(container, "paragraphs", []):
        yield paragraph
    for table in getattr(container, "tables", []):
        for row in table.rows:
            for cell in row.cells:
                yield from iter_paragraphs(cell)


def all_document_paragraphs(doc) -> list:
    paragraphs = []
    seen = set()

    def add_from(container):
        for paragraph in iter_paragraphs(container):
            marker = id(paragraph._p)
            if marker not in seen:
                seen.add(marker)
                paragraphs.append(paragraph)

    add_from(doc)
    for section in doc.sections:
        for part_name in (
            "header", "footer", "first_page_header", "first_page_footer",
            "even_page_header", "even_page_footer",
        ):
            add_from(getattr(section, part_name))
    return paragraphs


def text_snapshot(doc) -> dict:
    paragraphs = all_document_paragraphs(doc)
    texts = [p.text for p in paragraphs]
    nonempty = [value for value in texts if value.strip()]
    return {
        "paragraph_count": len(paragraphs),
        "nonempty_paragraph_count": len(nonempty),
        "text": "\n".join(texts),
        "characters": sum(len(value) for value in texts),
    }


def remove_graphics_from_package(doc) -> int:
    """Remove drawing/picture/OLE XML nodes from the main document and all related parts."""
    removed = 0
    visited_parts = set()
    for part in doc.part.package.parts:
        part_key = id(part)
        if part_key in visited_parts:
            continue
        visited_parts.add(part_key)
        root = getattr(part, "_element", None)
        if root is None:
            continue
        for element in list(root.iter()):
            if element.tag in GRAPHIC_TAGS:
                parent = element.getparent()
                if parent is not None:
                    parent.remove(element)
                    removed += 1
    return removed


def set_run_font(run, size: int = FONT_SIZE, bold: bool | None = None) -> None:
    run.font.name = FONT_NAME
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for key in ("ascii", "hAnsi", "cs", "eastAsia"):
        rfonts.set(qn(f"w:{key}"), FONT_NAME)


def format_document(source: Path, output_docx: Path) -> dict:
    doc = Document(str(source))
    before = text_snapshot(doc)
    removed_graphics = remove_graphics_from_package(doc) if REMOVE_GRAPHICS else 0

    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    for paragraph in all_document_paragraphs(doc):
        text = paragraph.text.strip()
        paragraph.paragraph_format.space_after = Pt(6)
        paragraph.paragraph_format.line_spacing = 1.15
        upper = text.upper()
        is_heading = upper.startswith((
            "SECTION", "PART", "QUESTION", "MULTIPLE CHOICE", "TRUE OR FALSE",
            "MATCHING", "FILL IN", "SHORT ANSWER", "DIAGRAM",
        ))
        if is_heading:
            paragraph.paragraph_format.keep_with_next = True
        for run in paragraph.runs:
            set_run_font(run, size=FONT_SIZE, bold=True if is_heading else None)

    for section in doc.sections:
        for part_name in (
            "header", "footer", "first_page_header", "first_page_footer",
            "even_page_header", "even_page_footer",
        ):
            part = getattr(section, part_name)
            for paragraph in iter_paragraphs(part):
                for run in paragraph.runs:
                    set_run_font(run, size=9)

    after = text_snapshot(doc)
    if before["text"] != after["text"]:
        raise RuntimeError(
            "Text-preservation check failed: paragraph text changed while removing graphics."
        )

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_docx))

    reopened = Document(str(output_docx))
    saved_snapshot = text_snapshot(reopened)
    if before["text"] != saved_snapshot["text"]:
        raise RuntimeError("Saved DOCX validation failed: extracted text differs from source.")

    return {
        "paragraph_count": before["paragraph_count"],
        "nonempty_paragraph_count": before["nonempty_paragraph_count"],
        "text_characters": before["characters"],
        "graphics_removed": removed_graphics,
    }


def output_stem(source: Path) -> str:
    """Disambiguate files with the same name in repository root and Incoming/."""
    if source.parent == ROOT / "Incoming":
        return f"{source.stem}_Incoming"
    return source.stem


def validate_pdf(pdf_path: Path, expected_text_characters: int) -> int:
    if not pdf_path.is_file() or pdf_path.stat().st_size == 0:
        raise RuntimeError("PDF export did not produce a non-empty file.")
    info = run_command(["pdfinfo", str(pdf_path)], description="PDF metadata validation").stdout
    match = re.search(r"^Pages:\s+(\d+)\s*$", info, flags=re.MULTILINE)
    if not match:
        raise RuntimeError("Could not determine the PDF page count.")
    pages = int(match.group(1))
    if pages < 1:
        raise RuntimeError("PDF contains no pages.")
    extracted = run_command(
        ["pdftotext", str(pdf_path), "-"], description="PDF text extraction validation"
    ).stdout
    if expected_text_characters and not extracted.strip():
        raise RuntimeError("PDF contains no extractable text although the source contains text.")
    return pages


def process_one(source: Path) -> dict:
    stem = output_stem(source)
    output_docx = OUTPUT / f"{stem}_Formatted.docx"
    output_pdf = OUTPUT / f"{stem}_Formatted.pdf"
    for stale in (output_docx, output_pdf):
        if stale.exists():
            stale.unlink()

    with tempfile.TemporaryDirectory(prefix="exam-format-") as temp:
        temp_dir = Path(temp)
        profile_dir = temp_dir / "lo-profile"
        if source.suffix.lower() == ".doc":
            converted = convert_legacy_doc(source, temp_dir, profile_dir)
        else:
            converted = source

        metrics = format_document(converted, output_docx)
        pages = None
        if EXPORT_PDF:
            pdf_profile = temp_dir / "lo-pdf-profile"
            run_command(
                libreoffice_command(
                    "--convert-to", "pdf", "--outdir", str(OUTPUT), str(output_docx),
                    profile_dir=pdf_profile,
                ),
                description=f"PDF export for {source.name}",
            )
            pages = validate_pdf(output_pdf, metrics["text_characters"])

    return {
        "source": source.relative_to(ROOT).as_posix(),
        "status": "success",
        "docx": output_docx.relative_to(ROOT).as_posix(),
        "pdf": output_pdf.relative_to(ROOT).as_posix() if EXPORT_PDF else None,
        "pages": pages,
        **metrics,
    }


def write_reports(records: list[dict], failures: list[dict]) -> None:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    report = {
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "settings": {
            "remove_graphics": REMOVE_GRAPHICS,
            "export_pdf": EXPORT_PDF,
            "font": FONT_NAME,
            "font_size_pt": FONT_SIZE,
        },
        "processed": records,
        "failures": failures,
        "summary": {
            "total": len(records) + len(failures),
            "successful": len(records),
            "failed": len(failures),
        },
    }
    (OUTPUT / "processing_report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )

    lines = [
        "# Exam processing report",
        "",
        f"- Generated (UTC): {report['generated_at_utc']}",
        f"- Successful: {len(records)}",
        f"- Failed: {len(failures)}",
        f"- Graphics removal: {'enabled' if REMOVE_GRAPHICS else 'disabled'}",
        "",
        "## Processed files",
        "",
        "| Source | Status | Pages | Paragraphs | Characters | Graphics removed |",
        "|---|---:|---:|---:|---:|---:|",
    ]
    for item in records:
        lines.append(
            f"| `{item['source']}` | OK | {item['pages'] or '—'} | "
            f"{item['paragraph_count']} | {item['text_characters']} | {item['graphics_removed']} |"
        )
        lines.append(f"  - DOCX: `{item['docx']}`")
        if item["pdf"]:
            lines.append(f"  - PDF: `{item['pdf']}`")
    if failures:
        lines.extend(["", "## Failures", ""])
        for item in failures:
            lines.append(f"- `{item['source']}`: {item['error']}")
    lines.extend([
        "",
        "Text-preservation checks compare extracted paragraph text before and after formatting. "
        "Visual layout still requires human review.",
        "",
    ])
    (OUTPUT / "processing_report.md").write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    OUTPUT.mkdir(exist_ok=True)
    candidates = []
    for folder in INPUT_DIRS:
        if folder.exists():
            candidates.extend(
                path for path in folder.iterdir()
                if path.is_file()
                and path.suffix.lower() in {".doc", ".docx"}
                and not path.name.startswith("~$")
                and not path.stem.endswith("_Formatted")
            )

    records = []
    failures = []
    for source in sorted(set(candidates), key=lambda item: item.as_posix().lower()):
        try:
            record = process_one(source)
            records.append(record)
            print(
                f"SUCCESS: {record['source']} -> {record['docx']} "
                f"(pages={record['pages']}, graphics_removed={record['graphics_removed']})"
            )
        except Exception as exc:
            failure = {"source": source.relative_to(ROOT).as_posix(), "error": str(exc)}
            failures.append(failure)
            print(f"ERROR: {failure['source']}: {failure['error']}", file=sys.stderr)

    write_reports(records, failures)
    print(
        f"REPORT: Formatted/processing_report.md and "
        f"Formatted/processing_report.json ({len(records)} successful, {len(failures)} failed)"
    )
    if failures:
        raise SystemExit(1)
    if not candidates:
        print("No Word documents found in repository root or Incoming/.")


if __name__ == "__main__":
    main()
