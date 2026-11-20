import unittest
from tempfile import TemporaryDirectory
from pathlib import Path
from docx import Document
from scripts.format_exams import classify, format_doc, snapshot

class FormatterTests(unittest.TestCase):
    def test_semantic_classification(self):
        d=Document()
        self.assertEqual(classify(d.add_paragraph("Question 2: Explain"),9),"question")
        self.assertEqual(classify(d.add_paragraph("SECTION B"),9),"section")
        self.assertEqual(classify(d.add_paragraph("(a) Define"),9),"subquestion")
        self.assertEqual(classify(d.add_paragraph("Duration: Two hours"),9),"metadata")
    def test_text_preservation_with_table(self):
        with TemporaryDirectory() as tmp:
            src=Path(tmp)/"source.docx"; dst=Path(tmp)/"formatted.docx"
            d=Document(); d.add_paragraph("College of Science"); d.add_paragraph("Final Examination")
            d.add_paragraph("Question 1: Explain")
            t=d.add_table(rows=1,cols=2); t.cell(0,0).text="Part A"; t.cell(0,1).text="Part B"; d.save(src)
            before=snapshot(Document(src))["text"]; format_doc(src,dst)
            self.assertEqual(before,snapshot(Document(dst))["text"])
if __name__=="__main__": unittest.main()
