import streamlit as st
import pandas as pd
import tempfile
import os
import urllib
from sqlalchemy import create_engine
from zipfile import ZipFile
import io

# 💡 Create Access DB engine using sqlalchemy-access
def connect_access_db(access_db):
    cnnstr = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        f'DBQ={access_db};'
    )
    cnnurl = f"access+pyodbc:///?odbc_connect={urllib.parse.quote_plus(cnnstr)}"
    return create_engine(cnnurl)

# 🗝️ Key generator for comparison
def generate_keys(df):
    df['key1'] = df['Invoice Date'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key2'] = df['Invoice Number'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key3'] = df['Invoice Number'].astype(str) + '_' + df['Invoice Date'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key4'] = df['Invoice Date'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str) + '_' + df['Invoice Number'].astype(str)
    return df

# 🧠 Matching logic
def check_match(row, key_sets):
    if row["key4"] in key_sets["key4"]:
        return pd.Series(["Yes", "Date+Amount+Supplier+Number"])
    elif row["key1"] in key_sets["key1"]:
        return pd.Series(["Yes", "Date+Amount+Supplier"])
    elif row["key2"] in key_sets["key2"]:
        return pd.Series(["Yes", "Number+Amount+Supplier"])
    elif row["key3"] in key_sets["key3"]:
        return pd.Series(["Yes", "Number+Date+Supplier"])
    else:
        return pd.Series(["No", "UNIQUE"])

# 🧭 App UI starts
st.set_page_config(page_title="Access to CSV & Deduplication", layout="centered")
st.title("📁 Access to CSV + Excel Duplicate Checker")

uploaded_file = st.file_uploader("Upload your .mdb or .accdb file", type=["mdb", "accdb"])

if uploaded_file is not None:
    with tempfile.NamedTemporaryFile(delete=False, suffix='.' + uploaded_file.name.split('.')[-1]) as tmp:
        tmp.write(uploaded_file.getbuffer())
        access_path = tmp.name

    try:
        with st.spinner("Connecting to database..."):
            engine = connect_access_db(access_path)
            with engine.connect() as conn:
                # Fetch all table names
                tables = conn.engine.table_names()
                if not tables:
                    st.error("No tables found in the database.")
                else:
                    # Convert all tables to CSV and zip them
                    with tempfile.TemporaryDirectory() as tempdir:
                        zip_path = os.path.join(tempdir, "tables_csv.zip")
                        with ZipFile(zip_path, 'w') as zipf:
                            for table in tables:
                                df = pd.read_sql(f"SELECT * FROM [{table}]", conn)
                                csv_filename = f"{table}.csv"
                                csv_path = os.path.join(tempdir, csv_filename)
                                df.to_csv(csv_path, index=False)
                                zipf.write(csv_path, arcname=csv_filename)
                        
                        with open(zip_path, "rb") as f:
                            st.download_button(
                                label="📥 Download All Tables as ZIP",
                                data=f,
                                file_name="tables_csv.zip",
                                mime="application/zip"
                            )
                    
                    # Load first table for duplicate matching
                    db_df = pd.read_sql(f"SELECT [Invoice Number], [Invoice Date], [Gross Amount], [Supplier Number] FROM [{tables[0]}]", conn)
                    db_df['Invoice Date'] = pd.to_datetime(db_df['Invoice Date'], errors='coerce')
                    db_df = generate_keys(db_df)
                    key_sets = {key: set(db_df[key]) for key in ['key1', 'key2', 'key3', 'key4']}
            st.success("✅ Conversion complete and reference data loaded!")
    except Exception as e:
        st.error(f"❌ Error processing the database: {e}")
    finally:
        os.remove(access_path)

    # Upload Excel and compare
    st.subheader("📊 Upload Excel Invoice File for Comparison")
    excel_file = st.file_uploader("Upload Excel file (.xlsx or .xls)", type=["xlsx", "xls"])
    if excel_file:
        try:
            excel_df = pd.read_excel(excel_file)
            excel_df.columns = excel_df.columns.str.strip()
            compare_df = excel_df[['Invoice Number', 'Invoice Date', 'Gross Amount', 'Supplier Number']].dropna()
            compare_df['Invoice Date'] = pd.to_datetime(compare_df['Invoice Date'], errors='coerce')
            compare_df = generate_keys(compare_df)

            with st.spinner("Checking for duplicates..."):
                compare_df[['Duplicate', 'Match Logic']] = compare_df.apply(lambda row: check_match(row, key_sets), axis=1)

            st.success("✅ Duplicate check complete.")
            st.dataframe(compare_df)

            # Download results as Excel
            buffer = io.BytesIO()
            output_df = excel_df.copy()
            output_df['Duplicate'] = compare_df['Duplicate']
            output_df['Match Logic'] = compare_df['Match Logic']
            output_df.to_excel(buffer, index=False)

            st.download_button("📥 Download Comparison Report", data=buffer.getvalue(), file_name="invoice_duplicates.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"❌ Error processing Excel file: {e}")
