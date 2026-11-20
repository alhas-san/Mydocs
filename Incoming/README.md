# Incoming exam documents

Upload Word exam files (`.doc` or `.docx`) directly into this folder.

When a matching file is pushed to GitHub, the **Format exam documents** workflow will:

1. Convert legacy `.doc` files to `.docx`.
2. Apply consistent A4 academic formatting (Times New Roman, 12 pt).
3. Remove embedded pictures, drawings and OLE objects from the document package, including headers and footers, while checking that extracted paragraph text is preserved.
4. Export a PDF copy and validate that it has at least one page and extractable text.
5. Save the formatted DOCX/PDF files in `Formatted/`.
6. Write `Formatted/processing_report.md` and `Formatted/processing_report.json` with page counts, paragraph/text counts, removed-graphics counts and any failures.

Files uploaded here receive an `_Incoming` suffix in their output names. This prevents collisions when a root-level file has the same name.

The original source file is never overwritten. Review the generated files visually before using them as final exam papers; automated text checks cannot guarantee perfect page layout.

Files placed directly in the repository root are also processed automatically. You can additionally start the workflow manually from the GitHub Actions tab using `workflow_dispatch`.
