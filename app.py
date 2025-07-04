import streamlit as st
import pandas as pd
import pyodbc
import os
import json
import io
from datetime import datetime
from zipfile import ZipFile
import tempfile

CONFIG_FILE = "config.json"

# 🔒 Save and load last DB path
def save_config(db_path):
    with open(CONFIG_FILE, 'w') as f:
        json.dump({"db_path": db_path}, f)

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f).get("db_path", "")
    return ""

# 📅 File metadata
def get_db_modified_time(path):
    return datetime.fromtimestamp(os.path.getmtime(path)).strftime('%Y-%m-%d %H:%M:%S')

# 🗝️ Key generator for matching
def generate_keys(df):
    df['key1'] = df['Invoice Date'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key2'] = df['Invoice Number'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key3'] = df['Invoice Number'].astype(str) + '_' + df['Invoice Date'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key4'] = df['Invoice Date'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str) + '_' + df['Invoice Number'].astype(str)
    return df

# 🔌 Access DB connector
def connect_access_db(path):
    return pyodbc.connect(rf'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={path};')

# 🧠 Matching logic
def check_match(row):
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

# 🌟 Streamlit UI
st.set_page_config(page_title="MS Access Invoice Checker", layout="centered")
st.title("📁 MS Access Invoice Analyzer")

# 📂 Upload Access DB
uploaded_db = st.file_uploader("Upload MS Access DB (.mdb or .accdb)", type=["mdb", "accdb"])
if not uploaded_db:
    st.stop()

# 🚀 Process uploaded Access DB
with st.spinner("Processing Access database..."):
    db_path = os.path.join(os.getcwd(), "uploaded_db.accdb")
    with open(db_path, "wb") as f:
        f.write(uploaded_db.getbuffer())
    save_config(db_path)

    try:
        conn = connect_access_db(db_path)
        cursor = conn.cursor()
        tables = [t.table_name for t in cursor.tables(tableType='TABLE') if not t.table_name.startswith("MSys")]

        if not tables:
            st.error("No valid tables found in Access database.")
            st.stop()

        table = tables[0]  # use first valid table

        # ⏱️ Convert all tables to CSV ZIP
        with tempfile.TemporaryDirectory() as tempdir:
            zip_path = os.path.join(tempdir, "access_tables.zip")
            with ZipFile(zip_path, 'w') as zipf:
                for tbl in tables:
                    df_tbl = pd.read_sql(f"SELECT * FROM [{tbl}]", conn)
                    csv_path = os.path.join(tempdir, f"{tbl}.csv")
                    df_tbl.to_csv(csv_path, index=False)
                    zipf.write(csv_path, arcname=f"{tbl}.csv")
            with open(zip_path, "rb") as f:
                st.download_button("📦 Download All Tables as CSV ZIP", data=f.read(), file_name="access_tables.zip", mime="application/zip")

        # Load relevant data for duplicate check
        query = f"SELECT [Invoice Number], [Invoice Date], [Gross Amount], [Supplier Number] FROM [{table}]"
        master_df = pd.read_sql(query, conn)
        master_df['Invoice Date'] = pd.to_datetime(master_df['Invoice Date'], errors='coerce')
        master_df = generate_keys(master_df)

    except Exception as e:
        st.error(f"❌ Error accessing DB: {e}")
        st.stop()

# 📥 Upload Excel invoice file
excel_file = st.file_uploader("Upload Excel Invoice File", type=["xlsx", "xls"])
if not excel_file:
    st.stop()

try:
    raw_df = pd.read_excel(excel_file)
    raw_df.columns = raw_df.columns.str.strip()
    df = raw_df[['Invoice Number', 'Invoice Date', 'Gross Amount', 'Supplier Number']].copy()
    df['Invoice Date'] = pd.to_datetime(df['Invoice Date'], errors='coerce')
    df = df.dropna(subset=['Invoice Date'])
    df = generate_keys(df)
except Exception as e:
    st.error(f"❌ Error reading Excel file: {e}")
    st.stop()

# 🚦 Matching
key_sets = {
    "key1": set(master_df["key1"]),
    "key2": set(master_df["key2"]),
    "key3": set(master_df["key3"]),
    "key4": set(master_df["key4"]),
}

with st.spinner("🔍 Checking for duplicates..."):
    df[["Duplicate", "Match Logic"]] = df.apply(check_match, axis=1)
    st.success("✅ Duplicate check completed.")
    st.dataframe(df)

# 📤 Excel export
excel_buffer = io.BytesIO()
final_df = raw_df.copy()
final_df['Duplicate'] = df['Duplicate']
final_df['Match Logic'] = df['Match Logic']
final_df.to_excel(excel_buffer, index=False)

st.download_button("📥 Download Report", data=excel_buffer.getvalue(), file_name="duplicates_report.xlsx")

# ➕ Optional merging of unique rows
if "No" in df["Duplicate"].values:
    if st.button("🔄 Merge UNIQUE rows into Access DB"):
        unique_df = df[df["Duplicate"] == "No"][["Invoice Number", "Invoice Date", "Gross Amount", "Supplier Number"]]
        try:
            with st.spinner("Adding unique rows to Access DB..."):
                for _, row in unique_df.iterrows():
                    cursor.execute(
                        f"""INSERT INTO [{table}] ([Invoice Number], [Invoice Date], [Gross Amount], [Supplier Number])
                            VALUES (?, ?, ?, ?)""",
                        row['Invoice Number'], row['Invoice Date'], row['Gross Amount'], row['Supplier Number']
                    )
                conn.commit()
                st.success("🎉 Unique rows added to Access database.")
        except Exception as e:
            st.error(f"❌ Error merging rows: {e}")
else:
    st.info("ℹ️ No unique records to merge.")

conn.close()
