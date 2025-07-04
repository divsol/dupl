import streamlit as st
import pandas as pd
import tempfile
import os
import subprocess
from zipfile import ZipFile
import io

# 🛠️ Helper functions
def extract_table_names(access_path):
    result = subprocess.run(['mdb-tables', '-1', access_path], capture_output=True, text=True)
    tables = result.stdout.strip().split('\n')
    return [tbl for tbl in tables if tbl]

def export_table_to_csv(access_path, table_name, output_path):
    with open(output_path, 'w') as f:
        subprocess.run(['mdb-export', access_path, table_name], stdout=f)

def generate_keys(df):
    df['key1'] = df['Invoice Date'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key2'] = df['Invoice Number'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key3'] = df['Invoice Number'].astype(str) + '_' + df['Invoice Date'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key4'] = df['Invoice Date'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str) + '_' + df['Invoice Number'].astype(str)
    return df

def check_match(row, key_sets):
    if row['key4'] in key_sets['key4']:
        return pd.Series(['Yes', 'Date+Amount+Supplier+Number'])
    elif row['key1'] in key_sets['key1']:
        return pd.Series(['Yes', 'Date+Amount+Supplier'])
    elif row['key2'] in key_sets['key2']:
        return pd.Series(['Yes', 'Number+Amount+Supplier'])
    elif row['key3'] in key_sets['key3']:
        return pd.Series(['Yes', 'Number+Date+Supplier'])
    else:
        return pd.Series(['No', 'UNIQUE'])

# 🧠 App layout
st.set_page_config(page_title="Access + Excel Invoice Deduplicator", layout="centered")
st.title("📁 Access to CSV + Invoice Deduplication (mdbtools)")

access_file = st.file_uploader("Upload your MS Access file (.mdb or .accdb)", type=["mdb", "accdb"])
if access_file:
    with st.spinner("Processing Access database..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".mdb") as tmp:
            tmp.write(access_file.getbuffer())
            access_path = tmp.name

        # 🔍 Extract table names
        tables = extract_table_names(access_path)
        if not tables:
            st.error("No tables found in Access file.")
        else:
            first_table = tables[0]

            # 🔄 Convert all tables to CSV and ZIP
            with tempfile.TemporaryDirectory() as tempdir:
                zip_path = os.path.join(tempdir, "access_tables.zip")
                with ZipFile(zip_path, 'w') as zipf:
                    table_frames = {}
                    for table in tables:
                        csv_file = os.path.join(tempdir, f"{table}.csv")
                        export_table_to_csv(access_path, table, csv_file)
                        zipf.write(csv_file, arcname=f"{table}.csv")

                        # Read first table for comparison
                        if table == first_table:
                            df = pd.read_csv(csv_file)
                            df = df[['Invoice Number', 'Invoice Date', 'Gross Amount', 'Supplier Number']].dropna()
                            df['Invoice Date'] = pd.to_datetime(df['Invoice Date'], errors='coerce')
                            db_df = generate_keys(df)

                # 📥 ZIP download
                with open(zip_path, "rb") as f:
                    st.download_button(
                        label="📦 Download All Tables as ZIP",
                        data=f.read(),
                        file_name="access_tables.zip",
                        mime="application/zip"
                    )

            os.remove(access_path)
            st.success("✅ Access DB processed!")

            # 🧮 Excel logic
            st.subheader("📊 Upload Excel Invoice File for Comparison")
            excel_file = st.file_uploader("Upload Excel file (.xlsx or .xls)", type=["xlsx", "xls"])
            if excel_file:
                try:
                    excel_df = pd.read_excel(excel_file)
                    excel_df.columns = excel_df.columns.str.strip()
                    raw = excel_df[['Invoice Number', 'Invoice Date', 'Gross Amount', 'Supplier Number']].dropna()
                    raw['Invoice Date'] = pd.to_datetime(raw['Invoice Date'], errors='coerce')
                    raw = generate_keys(raw)

                    # 🔎 Compare logic
                    key_sets = {key: set(db_df[key]) for key in ['key1', 'key2', 'key3', 'key4']}
                    raw[['Duplicate', 'Match Logic']] = raw.apply(lambda row: check_match(row, key_sets), axis=1)
                    st.success("✅ Comparison complete.")
                    st.dataframe(raw)

                    # 📤 Excel download
                    excel_buf = io.BytesIO()
                    result = excel_df.copy()
                    result['Duplicate'] = raw['Duplicate']
                    result['Match Logic'] = raw['Match Logic']
                    result.to_excel(excel_buf, index=False)

                    st.download_button("📥 Download Comparison Report", excel_buf.getvalue(), "invoice_duplicates.xlsx")

                except Exception as e:
                    st.error(f"❌ Error processing Excel file: {e}")
