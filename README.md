# File comparison web app

A local Python/Streamlit app that compares the extractable text in two files,
shows text items found in only one file, and creates a PowerPoint summary plus
a full CSV report.

## Supported formats

- PDF (text-based PDFs; scanned/image-only pages are not OCR'd)
- Word: `.docx`
- Excel: `.xlsx`
- PowerPoint: `.pptx`
- Plain text: `.txt`, `.md`
- Delimited text: `.csv`, `.tsv`

Files can use different supported formats. The comparison is based on extracted
text, not visual layout, formatting, formulas' calculated results, or images.
Excel formulas are compared as formula text. Each uploaded file is limited to
25 MB and 3,000 extracted text items.

## Run locally

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
py -m streamlit run app.py
```

Open the local URL printed by Streamlit, select two files, adjust the fuzzy
match threshold if needed, and click **Compare files**.

## Notes

Fuzzy text matching can mistake similar passages for matches or flag reworded
content as missing. Review the findings before relying on them. Scanned PDFs,
legacy `.doc`/`.xls` files, images, archives, and other unsupported formats are
not parsed by this app.
