import streamlit as st
import pandas as pd
import tempfile
import os
import urllib
from sqlalchemy import create_engine
from zipfile import ZipFile
import io

# ⚙️ Connect to Access DB using sqlalchemy-access
def connect_access_db(access_db):
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        f'DBQ={access_db};'
    )
    engine_url = f"access+pyodbc:///?odbc_connect={urllib.parse.quote_plus(conn_str)}"
    return create_engine(engine_url)

# 🗝️ Generate keys for duplicate detection
def generate_keys(df):
    df['key1'] = df['Invoice Date'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key2'] = df['Invoice Number'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key3'] = df['Invoice Number'].astype(str) + '_' + df['Invoice Date'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key4'] = df['Invoice Date'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str) + '_' + df['Invoice Number'].astype(str)
    return df

# 🧠 Match each invoice against database entries
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

# 🎬 UI Begins
st.set_page_config(page_title="Access Converter & Duplicate Checker")
st.title("📁 MS Access to CSV & Invoice Deduplication")

uploaded_file = st.file_uploader("Upload MS Access file (.mdb or .accdb)", type=["mdb", "accdb"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix='.' + uploaded_file.name.split('.')[-1]) as tmp:
        tmp.write(uploaded_file.getbuffer())
        access_path = tmp.name

    try:
        with st.spinner("Connecting to Access database..."):
            engine = connect_access_db(access_path)
            with engine.connect() as conn:
                tables = conn.engine.table_names()
                if not tables:
                    st.error("No tables found.")
                else:
                    # 🔁 Convert to ZIP
                    with tempfile.TemporaryDirectory() as tempdir:
                        zip_path = os.path.join(tempdir, "tables_csv.zip")
                        with ZipFile(zip_path, 'w') as zipf:
                            for table in tables:
                                df = pd.read_sql(f"SELECT * FROM [{table}]", conn)
                                csv_path = os.path.join(tempdir, f"{table}.csv")
                                df.to_csv(csv_path, index=False)
                                zipf.write(csv_path, arcname=os.path.basename(csv_path))
                        with open(zip_path, "rb") as f:
                            st.download_button("📦 Download All Tables as ZIP", f.read(), "tables_csv.zip", "application/zip")

                    # 💾 Load first table for comparison
                    base_query = f"SELECT [Invoice Number], [Invoice Date], [Gross Amount], [Supplier Number] FROM [{tables[0]}]"
                    db_df = pd.read_sql(base_query, conn)
                    db_df['Invoice Date'] = pd.to_datetime(db_df['Invoice Date'], errors='coerce')
                    db_df = generate_keys(db_df)
                    key_sets = {key: set(db_df[key]) for key in ['key1', 'key2', 'key3', 'key4']}
                    st.success("✅ Database loaded and converted.")
    except Exception as e:
        st.error(f"❌ Error accessing database: {e}")
    finally:
        os.remove(access_path)

    # 📊 Upload Excel for duplicate checking
    st.subheader("📄 Upload Excel Invoice File")
    excel_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])
    if excel_file:
        try:
            raw_df = pd.read_excel(excel_file)
            raw_df.columns = raw_df.columns.str.strip()
            df = raw_df[['Invoice Number', 'Invoice Date', 'Gross Amount', 'Supplier Number']].copy()
            df['Invoice Date'] = pd.to_datetime(df['Invoice Date'], errors='coerce')
            df = df.dropna(subset=['Invoice Date'])
            df = generate_keys(df)
            with st.spinner("Checking for duplicates..."):
                df[['Duplicate', 'Match Logic']] = df.apply(lambda row: check_match(row, key_sets), axis=1)
            st.success("✅ Duplicate check complete.")
            st.dataframe(df)

            # 📥 Download Excel
            excel_buf = io.BytesIO()
            result_df = raw_df.copy()
            result_df['Duplicate'] = df['Duplicate']
            result_df['Match Logic'] = df['Match Logic']
            result_df.to_excel(excel_buf, index=False)

            st.download_button("📥 Download Duplicate Report", excel_buf.getvalue(), "duplicates_report.xlsx")

        except Exception as e:
            st.error(f"❌ Error processing Excel file: {e}")
