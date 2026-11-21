import unittest
from tempfile import TemporaryDirectory
from pathlib import Path
from docx import Document
from scripts.format_exams import classify, format_doc, snapshot, reflow_mcq_options

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

    def test_short_mcq_options_are_paired(self):
        d=Document()
        for value in ("A. Walking frame","B. Blood test","C. X-ray film","D. Antibiotic"):
            d.add_paragraph(value)
        before=snapshot(d)["text"]
        self.assertEqual(reflow_mcq_options(d),4)
        self.assertEqual(len(d.paragraphs),1)
        self.assertEqual(d.paragraphs[0].text,"A. Walking frame\tB. Blood test\nC. X-ray film\tD. Antibiotic")
        self.assertEqual(before,snapshot(d)["text"])

    def test_long_mcq_options_stay_on_separate_lines(self):
        d=Document()
        for value in (
            "A. Improve quality of life and independence of people with disabilities",
            "B. Replace all clinical professionals with a complete automated engineering system",
            "C. Study only electrical circuits and laboratory measurement procedures",
            "D. Diagnose infectious diseases using clinical and laboratory investigations",
        ):
            d.add_paragraph(value)
        reflow_mcq_options(d)
        self.assertEqual(len(d.paragraphs),1)
        self.assertNotIn("\t",d.paragraphs[0].text)
        self.assertEqual(len(d.paragraphs[0].text.splitlines()),4)

    def test_inline_mcq_options_are_split_and_paired(self):
        d=Document()
        d.add_paragraph("A. Walking frame B. Blood test C. X-ray film D. Antibiotic")
        reflow_mcq_options(d)
        self.assertEqual(d.paragraphs[0].text,"A. Walking frame\tB. Blood test\nC. X-ray film\tD. Antibiotic")

    def test_one_long_choice_forces_all_choices_onto_separate_lines(self):
        d=Document()
        for value in (
            "A. Short",
            "B. This is the one long answer choice that cannot fit in the second column",
            "C. Also short",
            "D. Short too",
        ):
            d.add_paragraph(value)
        reflow_mcq_options(d)
        text=d.paragraphs[0].text
        self.assertNotIn("\\t",text)
        self.assertEqual(text.splitlines(),[
            "A. Short",
            "B. This is the one long answer choice that cannot fit in the second column",
            "C. Also short",
            "D. Short too",
        ])

    def test_short_options_use_one_fixed_column_tab_stop(self):
        d=Document()
        for value in ("A. Yes","B. No","C. Maybe","D. None"):
            d.add_paragraph(value)
        reflow_mcq_options(d)
        p=d.paragraphs[0]
        self.assertEqual(p.text,"A. Yes\\tB. No\\nC. Maybe\\tD. None")
        self.assertEqual(len(p.paragraph_format.tab_stops),1)

if __name__=="__main__":
    unittest.main()
