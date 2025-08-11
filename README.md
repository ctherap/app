# PDF to Excel Converter Streamlit App

This app extracts tabular data from a PDF file and converts it into an Excel spreadsheet.

## How to Run

1. Install dependencies:

```bash
pip install -r requirements.txt
```

2. Run the app:

```bash
streamlit run app.py
```

3. Upload your PDF with a table containing columns:

- #
- Full Name
- Hotel
- Room Type
- Date In
- Date Out
- Comment

4. Preview the extracted data and download the Excel file.

---

## Notes

- The app uses `pdfplumber` to extract tables from PDFs.
- Make sure your PDF contains clearly defined tables for best results.
