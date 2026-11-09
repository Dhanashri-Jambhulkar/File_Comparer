import io
import unittest

from docx import Document
from openpyxl import Workbook
from pptx import Presentation
from pypdf import PdfWriter

from app import (
    ContentUnit,
    Finding,
    compare_documents,
    create_powerpoint_report,
    extract_document,
)


class CompareDocumentsTests(unittest.TestCase):
    def test_exact_and_fuzzy_matches_leave_only_missing_content(self):
        left = [
            ContentUnit("Page 1", "The quarterly revenue was 100 million."),
            ContentUnit("Page 2", "A paragraph that is only in the first file."),
        ]
        right = [
            ContentUnit("Sheet A1", "The quarterly revenue was 100 million!"),
            ContentUnit("Sheet A2", "A separate note only in the second file."),
        ]

        left_only, right_only = compare_documents(left, right, threshold=88)

        self.assertEqual([item.text for item in left_only], [
            "A paragraph that is only in the first file."
        ])
        self.assertEqual([item.text for item in right_only], [
            "A separate note only in the second file."
        ])

    def test_duplicate_text_is_matched_one_to_one(self):
        left = [ContentUnit("Line 1", "same"), ContentUnit("Line 2", "same")]
        right = [ContentUnit("Line 1", "same")]

        left_only, right_only = compare_documents(left, right, threshold=90)

        self.assertEqual(len(left_only), 1)
        self.assertEqual(left_only[0].location, "Line 2")
        self.assertEqual(right_only, [])


class ExtractionTests(unittest.TestCase):
    def test_text_and_csv_extractors(self):
        text_units = extract_document("notes.txt", b"First sentence. Second sentence.\n")
        csv_units = extract_document("table.csv", b"name,value\nA,1\n")

        self.assertEqual(len(text_units), 2)
        self.assertEqual(len(csv_units), 4)
        self.assertIn("Row 2, column 2", [unit.location for unit in csv_units])

    def test_docx_extractor(self):
        document = Document()
        document.add_paragraph("Word document content.")
        data = io.BytesIO()
        document.save(data)

        units = extract_document("document.docx", data.getvalue())

        self.assertIn("Word document content.", [unit.text for unit in units])

    def test_xlsx_extractor(self):
        workbook = Workbook()
        workbook.active["A1"] = "Spreadsheet content"
        data = io.BytesIO()
        workbook.save(data)

        units = extract_document("book.xlsx", data.getvalue())

        self.assertEqual(units[0].location, "Sheet Sheet, cell A1")
        self.assertEqual(units[0].text, "Spreadsheet content")

    def test_pptx_extractor(self):
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        slide.shapes.add_textbox(0, 0, 3000000, 500000).text = "Slide content"
        data = io.BytesIO()
        presentation.save(data)

        units = extract_document("deck.pptx", data.getvalue())

        self.assertEqual(units[0].text, "Slide content")
        self.assertTrue(units[0].location.startswith("Slide 1"))

    def test_pdf_extractor_handles_page_without_text(self):
        writer = PdfWriter()
        writer.add_blank_page(width=200, height=200)
        data = io.BytesIO()
        writer.write(data)

        self.assertEqual(extract_document("blank.pdf", data.getvalue()), [])


class ReportTests(unittest.TestCase):
    def test_powerpoint_report_is_generated(self):
        report = create_powerpoint_report(
            "first.txt",
            "second.txt",
            [Finding("Line 1", "Missing from second")],
            [],
            1,
            1,
            88,
        )

        presentation = Presentation(io.BytesIO(report))
        self.assertEqual(len(presentation.slides), 2)
        self.assertIn(
            "Document comparison summary",
            presentation.slides[0].shapes[0].text,
        )


if __name__ == "__main__":
    unittest.main()
