import unittest
from tempfile import TemporaryDirectory
from pathlib import Path
from docx import Document
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from scripts.format_exams import classify, format_doc, snapshot, reflow_mcq_options, remove_unwanted_page_breaks, parse_options

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
        self.assertEqual(len(d.paragraphs),2)
        self.assertEqual(d.paragraphs[0].text.split("\t"),["A. Walking frame","B. Blood test"])
        self.assertEqual(d.paragraphs[1].text.split("\t"),["C. X-ray film","D. Antibiotic"])

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


    def test_question_stays_with_its_option_block(self):
        d=Document()
        q=d.add_paragraph("Question 1: Choose the correct answer")
        for value in ("A. Walking frame","B. Blood test","C. X-ray film","D. Antibiotic"):
            d.add_paragraph(value)
        reflow_mcq_options(d)
        self.assertTrue(q.paragraph_format.keep_with_next)
        self.assertTrue(d.paragraphs[1].paragraph_format.keep_with_next)
        self.assertFalse(d.paragraphs[2].paragraph_format.keep_with_next)

    def test_section_break_page_start_becomes_continuous(self):
        d=Document()
        p=d.add_paragraph("Section")
        sect=OxmlElement("w:sectPr")
        typ=OxmlElement("w:type")
        typ.set(qn("w:val"),"nextPage")
        sect.append(typ)
        p._p.get_or_add_pPr().append(sect)
        removed=remove_unwanted_page_breaks(d)
        self.assertEqual(removed,1)
        self.assertEqual(typ.get(qn("w:val")),"continuous")

    def test_mixed_inline_and_separate_options_reflow(self):
        d=Document()
        d.add_paragraph("A. First B. Second")
        d.add_paragraph("C. Third D. Fourth")
        reflow_mcq_options(d)
        self.assertEqual(len(d.paragraphs),2)
        self.assertEqual(d.paragraphs[0].text.split("\t"),["A. First","B. Second"])
        self.assertEqual(d.paragraphs[1].text.split("\t"),["C. Third","D. Fourth"])


    def test_arabic_question_and_section_classification(self):
        d=Document()
        self.assertEqual(classify(d.add_paragraph("السؤال ١: عرّف المصطلح"),9),"question")
        self.assertEqual(classify(d.add_paragraph("سؤال الثاني: أجب عن الآتي"),9),"question")
        self.assertEqual(classify(d.add_paragraph("القسم الثاني"),9),"section")
        self.assertEqual(classify(d.add_paragraph("اسم الطالب: أحمد"),9),"metadata")

    def test_arabic_option_labels_are_recognized_and_reflowed_rtl(self):
        d=Document()
        d.add_paragraph("السؤال ١: اختر الإجابة الصحيحة")
        for value in ("أ. المشي","ب. الجري","ج. السباحة","د. القفز"):
            d.add_paragraph(value)
        self.assertEqual([x[0] for x in parse_options("أ. المشي ب. الجري ج. السباحة د. القفز")],list("ABCD"))
        self.assertEqual(reflow_mcq_options(d),4)
        self.assertEqual(len(d.paragraphs),3)
        self.assertEqual(d.paragraphs[1].text.split("\t"),["أ. المشي","ب. الجري"])
        self.assertEqual(d.paragraphs[2].text.split("\t"),["ج. السباحة","د. القفز"])
        for p in d.paragraphs[1:]:
            self.assertEqual(p.alignment,2)  # WD_ALIGN_PARAGRAPH.RIGHT
            self.assertIsNotNone(p._p.get_or_add_pPr().find(qn("w:bidi")))


    def test_arabic_question_with_latin_option_labels_keeps_rtl_direction(self):
        d=Document()
        d.add_paragraph("السؤال ٢: اختر الإجابة الصحيحة")
        for value in ("A- First choice","B- Second choice","C- Third choice","D- Fourth choice"):
            d.add_paragraph(value)
        self.assertEqual(reflow_mcq_options(d),4)
        self.assertEqual(len(d.paragraphs),3)
        for p in d.paragraphs[1:]:
            self.assertIsNotNone(p._p.get_or_add_pPr().find(qn("w:bidi")))


    def test_arabic_document_paragraphs_are_rtl_but_english_remains_ltr(self):
        with TemporaryDirectory() as tmp:
            src=Path(tmp)/"arabic.docx"; dst=Path(tmp)/"formatted.docx"
            d=Document()
            d.add_paragraph("كلية العلوم والتكنولوجيا")
            d.add_paragraph("Question 1: Explain the process")
            d.save(src)
            format_doc(src,dst)
            result=Document(dst)
            arabic=result.paragraphs[0]
            english=result.paragraphs[1]
            self.assertIsNotNone(arabic._p.get_or_add_pPr().find(qn("w:bidi")))
            self.assertIsNone(english._p.get_or_add_pPr().find(qn("w:bidi")))


    def test_professional_arabic_and_latin_fonts_are_applied(self):
        with TemporaryDirectory() as tmp:
            src=Path(tmp)/"fonts.docx"; dst=Path(tmp)/"formatted.docx"
            d=Document()
            d.add_paragraph("كلية العلوم والتكنولوجيا")
            d.add_paragraph("هذا نص عربي عادي لاختبار حجم المتن")
            d.add_paragraph("Computer Networks")
            d.save(src)
            format_doc(src,dst)
            result=Document(dst)
            arabic_title_run=result.paragraphs[0].runs[0]
            arabic_body_run=result.paragraphs[1].runs[0]
            english_run=result.paragraphs[2].runs[0]
            title_fonts=arabic_title_run._element.get_or_add_rPr().rFonts
            body_fonts=arabic_body_run._element.get_or_add_rPr().rFonts
            english_fonts=english_run._element.get_or_add_rPr().rFonts
            self.assertEqual(title_fonts.get(qn("w:cs")),"Traditional Arabic")
            self.assertEqual(title_fonts.get(qn("w:ascii")),"Calibri")
            self.assertEqual(arabic_title_run.font.size.pt,18)  # 15 pt title + Arabic adjustment
            self.assertEqual(body_fonts.get(qn("w:cs")),"Traditional Arabic")
            self.assertEqual(arabic_body_run.font.size.pt,14)  # 11 pt body + Arabic adjustment
            self.assertEqual(english_fonts.get(qn("w:ascii")),"Calibri")
            self.assertEqual(english_run.font.size.pt,11)



if __name__=="__main__":
    unittest.main()
