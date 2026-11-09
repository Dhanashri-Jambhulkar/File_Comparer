from __future__ import annotations

import csv
import io
import re
from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO

import streamlit as st
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from pypdf import PdfReader
from rapidfuzz import fuzz, process


SUPPORTED_EXTENSIONS = {
    ".pdf",
    ".docx",
    ".xlsx",
    ".pptx",
    ".txt",
    ".md",
    ".csv",
    ".tsv",
}
MAX_FILE_SIZE = 25 * 1024 * 1024
MAX_UNITS = 3000
SENTENCE_BOUNDARY = re.compile(r"(?<=[.!?])\s+")
NORMALIZE_PATTERN = re.compile(r"[^\w]+", flags=re.UNICODE)


@dataclass(frozen=True)
class ContentUnit:
    location: str
    text: str


@dataclass(frozen=True)
class Finding:
    location: str
    text: str


def _add_text_units(units: list[ContentUnit], location: str, text: str) -> None:
    for part in SENTENCE_BOUNDARY.split(text):
        cleaned = " ".join(part.split())
        if cleaned:
            units.append(ContentUnit(location=location, text=cleaned))


def _extract_pdf(data: BinaryIO) -> list[ContentUnit]:
    units: list[ContentUnit] = []
    reader = PdfReader(data)
    for page_number, page in enumerate(reader.pages, start=1):
        _add_text_units(units, f"Page {page_number}", page.extract_text() or "")
    return units


def _extract_docx(data: BinaryIO) -> list[ContentUnit]:
    units: list[ContentUnit] = []
    document = Document(data)
    for paragraph_number, paragraph in enumerate(document.paragraphs, start=1):
        _add_text_units(units, f"Paragraph {paragraph_number}", paragraph.text)

    for table_number, table in enumerate(document.tables, start=1):
        for row_number, row in enumerate(table.rows, start=1):
            for column_number, cell in enumerate(row.cells, start=1):
                _add_text_units(
                    units,
                    f"Table {table_number}, row {row_number}, column {column_number}",
                    cell.text,
                )

    for section_number, section in enumerate(document.sections, start=1):
        for label, part in (("Header", section.header), ("Footer", section.footer)):
            for paragraph_number, paragraph in enumerate(part.paragraphs, start=1):
                _add_text_units(
                    units,
                    f"Section {section_number} {label}, paragraph {paragraph_number}",
                    paragraph.text,
                )
    return units


def _extract_xlsx(data: BinaryIO) -> list[ContentUnit]:
    units: list[ContentUnit] = []
    workbook = load_workbook(data, read_only=True, data_only=False)
    try:
        for worksheet in workbook.worksheets:
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        text = str(cell.value).strip()
                        if text:
                            units.append(
                                ContentUnit(
                                    location=f"Sheet {worksheet.title}, cell {cell.coordinate}",
                                    text=text,
                                )
                            )
    finally:
        workbook.close()
    return units


def _extract_pptx(data: BinaryIO) -> list[ContentUnit]:
    units: list[ContentUnit] = []
    presentation = Presentation(data)
    for slide_number, slide in enumerate(presentation.slides, start=1):
        for shape_number, shape in enumerate(slide.shapes, start=1):
            if shape.has_text_frame:
                _add_text_units(
                    units,
                    f"Slide {slide_number}, shape {shape_number}",
                    shape.text,
                )
            if shape.has_table:
                for row_number, row in enumerate(shape.table.rows, start=1):
                    for column_number, cell in enumerate(row.cells, start=1):
                        _add_text_units(
                            units,
                            (
                                f"Slide {slide_number}, table shape {shape_number}, "
                                f"row {row_number}, column {column_number}"
                            ),
                            cell.text,
                        )
    return units


def _decode_text(data: bytes) -> str:
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("cp1252")


def _extract_delimited(data: BinaryIO, extension: str) -> list[ContentUnit]:
    content = _decode_text(data.read())
    if extension == ".tsv":
        dialect = csv.excel_tab
    else:
        try:
            dialect = csv.Sniffer().sniff(content[:8192], delimiters=",;\t|")
        except csv.Error:
            dialect = csv.excel

    units: list[ContentUnit] = []
    for row_number, row in enumerate(csv.reader(io.StringIO(content), dialect), start=1):
        for column_number, value in enumerate(row, start=1):
            if value.strip():
                _add_text_units(
                    units,
                    f"Row {row_number}, column {column_number}",
                    value,
                )
    return units


def extract_document(filename: str, data: bytes) -> list[ContentUnit]:
    """Extract comparable text units from a supported document."""
    extension = Path(filename).suffix.lower()
    if extension not in SUPPORTED_EXTENSIONS:
        supported = ", ".join(sorted(SUPPORTED_EXTENSIONS))
        raise ValueError(f"Unsupported file type {extension or '(no extension)'}. Supported: {supported}")
    if len(data) > MAX_FILE_SIZE:
        raise ValueError("Each file must be 25 MB or smaller.")

    buffer = io.BytesIO(data)
    if extension == ".pdf":
        units = _extract_pdf(buffer)
    elif extension == ".docx":
        units = _extract_docx(buffer)
    elif extension == ".xlsx":
        units = _extract_xlsx(buffer)
    elif extension == ".pptx":
        units = _extract_pptx(buffer)
    elif extension in {".csv", ".tsv"}:
        units = _extract_delimited(buffer, extension)
    else:
        content = _decode_text(data)
        units = []
        for line_number, line in enumerate(content.splitlines(), start=1):
            _add_text_units(units, f"Line {line_number}", line)

    if len(units) > MAX_UNITS:
        raise ValueError(
            f"{filename} contains more than {MAX_UNITS:,} text items. "
            "Split the file into smaller parts and compare again."
        )
    return units


def _normalized(text: str) -> str:
    return NORMALIZE_PATTERN.sub(" ", text.casefold()).strip()


def compare_documents(
    left: list[ContentUnit],
    right: list[ContentUnit],
    threshold: int,
) -> tuple[list[Finding], list[Finding]]:
    """Return content unique to the left and right documents, respectively."""
    right_by_text: dict[str, list[int]] = {}
    for index, unit in enumerate(right):
        key = _normalized(unit.text)
        if key:
            right_by_text.setdefault(key, []).append(index)

    unmatched_left: list[int] = []
    available_right = set(range(len(right)))
    left_only: list[Finding] = []

    for left_index, unit in enumerate(left):
        key = _normalized(unit.text)
        exact_matches = right_by_text.get(key, [])
        match_index = next(
            (index for index in exact_matches if index in available_right),
            None,
        )
        if match_index is None:
            unmatched_left.append(left_index)
        else:
            available_right.remove(match_index)

    right_choices = {
        index: _normalized(right[index].text)
        for index in available_right
    }
    fuzzy_candidates: list[tuple[float, int, int]] = []
    for left_index in unmatched_left:
        candidates = process.extract(
            _normalized(left[left_index].text),
            right_choices,
            scorer=fuzz.ratio,
            limit=5,
        )
        for _, score, right_index in candidates:
            if score >= threshold:
                fuzzy_candidates.append((score, left_index, right_index))

    matched_left: set[int] = set()
    for score, left_index, right_index in sorted(fuzzy_candidates, reverse=True):
        if left_index in matched_left or right_index not in available_right:
            continue
        matched_left.add(left_index)
        available_right.remove(right_index)

    for left_index in unmatched_left:
        if left_index not in matched_left:
            unit = left[left_index]
            left_only.append(Finding(unit.location, unit.text))

    right_only = [
        Finding(right[index].location, right[index].text)
        for index in sorted(available_right)
    ]
    return left_only, right_only


def _add_report_slide(presentation: Presentation, title: str, lines: list[str]) -> None:
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    title_box = slide.shapes.add_textbox(Inches(0.55), Inches(0.35), Inches(12.2), Inches(0.65))
    title_paragraph = title_box.text_frame.paragraphs[0]
    title_paragraph.text = title
    title_paragraph.font.size = Pt(24)
    title_paragraph.font.bold = True
    title_paragraph.font.color.rgb = RGBColor(31, 78, 121)

    body = slide.shapes.add_textbox(Inches(0.7), Inches(1.2), Inches(12), Inches(5.8))
    body.text_frame.word_wrap = True
    for index, line in enumerate(lines):
        paragraph = body.text_frame.paragraphs[0] if index == 0 else body.text_frame.add_paragraph()
        paragraph.text = line
        paragraph.font.size = Pt(13)
        paragraph.space_after = Pt(10)


def create_powerpoint_report(
    left_name: str,
    right_name: str,
    left_only: list[Finding],
    right_only: list[Finding],
    left_count: int,
    right_count: int,
    threshold: int,
) -> bytes:
    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)

    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.7), Inches(1))
    title = title_box.text_frame.paragraphs[0]
    title.text = "Document comparison summary"
    title.font.size = Pt(30)
    title.font.bold = True
    title.font.color.rgb = RGBColor(31, 78, 121)

    subtitle_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.7), Inches(1))
    subtitle = subtitle_box.text_frame.paragraphs[0]
    subtitle.text = f"{left_name}\ncompared with\n{right_name}"
    subtitle.font.size = Pt(18)

    summary = [
        f"Text items checked: {left_count:,} in the first file; {right_count:,} in the second file.",
        f"Possible content missing from the second file: {len(left_only):,} item(s).",
        f"Possible content missing from the first file: {len(right_only):,} item(s).",
        f"Fuzzy match threshold: {threshold}%.",
    ]
    summary_box = slide.shapes.add_textbox(Inches(0.8), Inches(3.2), Inches(11.7), Inches(2.5))
    for index, line in enumerate(summary):
        paragraph = summary_box.text_frame.paragraphs[0] if index == 0 else summary_box.text_frame.add_paragraph()
        paragraph.text = line
        paragraph.font.size = Pt(18)
        paragraph.space_after = Pt(12)

    if not left_only and not right_only:
        _add_report_slide(presentation, "No missing items detected", [
            "All extracted text items matched at the selected threshold."
        ])

    for title, findings in (
        (f"Found in {left_name}, not in {right_name}", left_only),
        (f"Found in {right_name}, not in {left_name}", right_only),
    ):
        for offset in range(0, len(findings), 8):
            lines = [
                f"{item.location}: {item.text[:500]}"
                for item in findings[offset : offset + 8]
            ]
            _add_report_slide(
                presentation,
                f"{title} ({offset + 1}-{min(offset + 8, len(findings))})",
                lines,
            )

    output = io.BytesIO()
    presentation.save(output)
    return output.getvalue()


def create_csv_report(
    left_name: str,
    right_name: str,
    left_only: list[Finding],
    right_only: list[Finding],
) -> bytes:
    output = io.StringIO(newline="")
    writer = csv.writer(output)
    writer.writerow(["Missing from file", "Source file", "Location", "Content"])
    for missing_from, source_name, findings in (
        (right_name, left_name, left_only),
        (left_name, right_name, right_only),
    ):
        for finding in findings:
            writer.writerow([missing_from, source_name, finding.location, finding.text])
    return output.getvalue().encode("utf-8-sig")


def main() -> None:
    st.set_page_config(page_title="File comparison", page_icon="📄", layout="wide")
    st.title("Compare two files")
    st.write(
        "Find text present in one file but not the other, even when the documents use "
        "different supported formats. Download a PowerPoint summary and a full CSV list."
    )
    st.caption(
        "Supported: PDF, DOCX, XLSX, PPTX, TXT, Markdown, CSV, and TSV. "
        "Scanned PDFs and image-only documents need OCR, which is not included."
    )

    upload_column1, upload_column2 = st.columns(2)
    with upload_column1:
        first_file = st.file_uploader(
            "First file",
            type=[item[1:] for item in SUPPORTED_EXTENSIONS],
            key="first",
        )
        if first_file is not None:
            st.caption(f"Selected: **{first_file.name}** · {first_file.size:,} bytes")
    with upload_column2:
        second_file = st.file_uploader(
            "Second file",
            type=[item[1:] for item in SUPPORTED_EXTENSIONS],
            key="second",
        )
        if second_file is not None:
            st.caption(f"Selected: **{second_file.name}** · {second_file.size:,} bytes")

    threshold = st.slider(
        "Fuzzy match threshold",
        min_value=70,
        max_value=100,
        value=88,
        help="Higher values require closer text matches. Lower values may treat different text as a match.",
    )

    if first_file is None or second_file is None:
        st.info("Upload both files above to compare them and see the results here.")
        return

    try:
        with st.spinner("Reading and comparing files..."):
            first_units = extract_document(first_file.name, first_file.getvalue())
            second_units = extract_document(second_file.name, second_file.getvalue())
            first_only, second_only = compare_documents(
                first_units,
                second_units,
                threshold,
            )
    except Exception as error:
        st.error(f"Could not read the selected files: {error}")
        return

    if not first_units and not second_units:
        st.warning("No extractable text was found in either file.")
        return

    st.success(
        f"Comparison complete: {len(first_units):,} text items checked in "
        f"{first_file.name} and {len(second_units):,} in {second_file.name}."
    )
    left_column, right_column = st.columns(2)
    left_column.metric(f"Found only in {first_file.name}", len(first_only))
    right_column.metric(f"Found only in {second_file.name}", len(second_only))

    first_tab, second_tab = st.tabs(
        [f"Only in {first_file.name}", f"Only in {second_file.name}"]
    )
    with first_tab:
        if first_only:
            for item in first_only:
                preview = item.text[:100].replace("\n", " ")
                if len(item.text) > 100:
                    preview += "..."
                with st.expander(f"{item.location}: {preview}"):
                    st.text(item.text)
        else:
            st.info("No unique text items found.")
    with second_tab:
        if second_only:
            for item in second_only:
                preview = item.text[:100].replace("\n", " ")
                if len(item.text) > 100:
                    preview += "..."
                with st.expander(f"{item.location}: {preview}"):
                    st.text(item.text)
        else:
            st.info("No unique text items found.")

    powerpoint = create_powerpoint_report(
        first_file.name,
        second_file.name,
        first_only,
        second_only,
        len(first_units),
        len(second_units),
        threshold,
    )
    csv_report = create_csv_report(
        first_file.name,
        second_file.name,
        first_only,
        second_only,
    )
    download_column, csv_column = st.columns(2)
    download_column.download_button(
        "Download PowerPoint summary",
        data=powerpoint,
        file_name="file_comparison_summary.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
    csv_column.download_button(
        "Download full results (CSV)",
        data=csv_report,
        file_name="file_comparison_results.csv",
        mime="text/csv",
    )


if __name__ == "__main__":
    main()
