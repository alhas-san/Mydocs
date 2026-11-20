# Incoming exam documents

Upload Word exam files (.doc or .docx) to this folder.

GitHub Actions will automatically:
- Convert legacy .doc files to .docx.
- Apply consistent academic formatting.
- Remove embedded diagrams and images (enabled by default).
- Export both DOCX and PDF copies.
- Save the generated files in the Formatted/ folder.

The original file is never overwritten.

To process files already in the repository root, open the Actions tab and run
"Format exam documents" manually using workflow_dispatch.
