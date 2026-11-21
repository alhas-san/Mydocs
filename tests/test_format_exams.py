import unittest
from tempfile import TemporaryDirectory
from pathlib import Path
from docx import Document
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from scripts.format_exams import classify, format_doc, snapshot, reflow_mcq_options, remove_unwanted_page_breaks

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
            table=d.add_table(rows=1,cols=2); table.cell(0,0).text="Part A"; table.cell(0,1).text="Part B"; d.save(src)
            before=snapshot(Document(src))["text"]; format_doc(src,dst)
            self.assertEqual(before,snapshot(Document(dst))["text"])

    def test_mcq_options_each_get_uniform_paragraphs(self):
        d=Document()
        for value in ("A. Walking frame","B. Blood test","C. X-ray film","D. Antibiotic"):
            d.add_paragraph(value)
        before=snapshot(d)["text"]
        self.assertEqual(reflow_mcq_options(d),4)
        self.assertEqual(len(d.paragraphs),2)
        self.assertEqual(d.paragraphs[0].text.split("\t"),["A. Walking frame","B. Blood test"])
        self.assertEqual(d.paragraphs[1].text.split("\t"),["C. X-ray film","D. Antibiotic"])
        for p in d.paragraphs:
            self.assertAlmostEqual(p.paragraph_format.left_indent.cm,0.35,places=2)
            self.assertEqual(p.paragraph_format.first_line_indent.cm,0)
            self.assertTrue(p.paragraph_format.keep_together)
        self.assertTrue(d.paragraphs[0].paragraph_format.keep_with_next)
        self.assertFalse(d.paragraphs[1].paragraph_format.keep_with_next)
        self.assertEqual(before,snapshot(d)["text"])

    def test_inline_mcq_options_become_four_paragraphs(self):
        d=Document()
        d.add_paragraph("A. Walking frame B. Blood test C. X-ray film D. Antibiotic")
        reflow_mcq_options(d)
        self.assertEqual([p.text for p in d.paragraphs],[
            "A. Walking frame","B. Blood test","C. X-ray film","D. Antibiotic"])

    def test_long_choice_does_not_change_uniform_layout(self):
        d=Document()
        for value in (
            "A. Short",
            "B. This is the one long answer choice that wraps across several lines",
            "C. Also short",
            "D. Short too",
        ):
            d.add_paragraph(value)
        reflow_mcq_options(d)
        self.assertEqual(len(d.paragraphs),4)
        self.assertTrue(all(p.paragraph_format.keep_together for p in d.paragraphs))
        self.assertTrue(all("\t" not in p.text for p in d.paragraphs))

    def test_inline_and_separate_options_can_be_combined(self):
        d=Document()
        d.add_paragraph("A. One B. Two")
        d.add_paragraph("C. Three")
        d.add_paragraph("D. Four")
        reflow_mcq_options(d)
        self.assertEqual(len(d.paragraphs),2)
        self.assertEqual(d.paragraphs[0].text.split("\t"),["A. One","B. Two"])
        self.assertEqual(d.paragraphs[1].text.split("\t"),["C. Three","D. Four"])

    def test_removes_manual_page_breaks(self):
        d=Document()
        p=d.add_paragraph("Question")
        p.paragraph_format.page_break_before=True
        p.add_run().add_break(WD_BREAK.PAGE)
        removed=remove_unwanted_page_breaks(d)
        self.assertGreaterEqual(removed,1)
        self.assertFalse(p.paragraph_format.page_break_before)

if __name__=="__main__":
    unittest.main()
