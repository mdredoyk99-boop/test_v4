# TRF Generator

    pip install -r requirements.txt
    streamlit run app.py

**Poppler is required** (pdf2image + pdftotext use it).
* Windows: download Poppler for Windows, unzip, add its `Library\bin` folder to PATH.
* macOS: `brew install poppler`   * Linux: `apt install poppler-utils`

Layout: `app.py` (UI) · `trf/excel_reader.py` · `trf/techpack.py` · `trf/docx_filler.py` ·
`trf/prepare_template.py` (one-off tagger) · `trf/config.py` (dropdowns, aliases, file-name pattern) ·
`templates/TRF_TEMPLATE.docx` (your demo TRF with {{tags}}).

Re-tag a new/edited original template:
    python -m trf.prepare_template "original TRF.docx" templates/TRF_TEMPLATE.docx

## Online deployment
Streamlit Community Cloud / Hugging Face Spaces: `packages.txt` (poppler-utils) is already included.
Main file: `app.py`.
