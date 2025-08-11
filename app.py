import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.title("PDF to Excel Converter")

uploaded_file = st.file_uploader("Upload PDF with table", type=["pdf"])

def extract_table_from_pdf(pdf_file):
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if any(cell and cell.strip() != '' for cell in row):
                        rows.append(row)
    return rows

def rows_to_df(rows):
    expected_headers = ['#', 'Full Name', 'Hotel', 'Room Type', 'Date In', 'Date Out', 'Comment']
    if not rows:
        return pd.DataFrame(columns=expected_headers)

    df = pd.DataFrame(rows[1:], columns=rows[0])

    # Normalize headers for matching
    df.columns = [col.strip().lower().replace(' ', '') for col in df.columns]
    mapping = {}
    for expected in expected_headers:
        key = expected.lower().replace(' ', '')
        matched_cols = [c for c in df.columns if key in c or c in key]
        mapping[expected] = matched_cols[0] if matched_cols else None

    filtered_data = {}
    for col in expected_headers:
        if mapping[col]:
            filtered_data[col] = df[mapping[col]]
        else:
            filtered_data[col] = [''] * len(df)

    new_df = pd.DataFrame(filtered_data)

    # Debug print raw '#' column values
    print("Raw # column values:", new_df['#'].tolist())

    # Remove rows where '#' is empty or header label
    new_df = new_df[(new_df['#'].str.strip() != '') & (new_df['#'].str.strip() != '#')]

    # Filter rows where '#' is numeric
    new_df = new_df[new_df['#'].apply(lambda x: str(x).strip().isdigit())]

    # Convert '#' to int and sort ascending
    new_df['#'] = new_df['#'].astype(int)
    new_df = new_df.sort_values('#').reset_index(drop=True)

    # Drop duplicates if any
    new_df = new_df.drop_duplicates()

    return new_df

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Accommodation Report')
    return output.getvalue()

if uploaded_file:
    try:
        rows = extract_table_from_pdf(uploaded_file)
        if rows:
            df = rows_to_df(rows)
            st.subheader("Preview of extracted data")
            st.dataframe(df)
            excel_data = to_excel(df)
            st.download_button(label="Download Excel file",
                               data=excel_data,
                               file_name="Accommodation_Report.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.error("No table data found in the PDF.")
    except Exception as e:
        st.error(f"Error processing PDF: {e}")
