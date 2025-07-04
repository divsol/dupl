import streamlit as st
import pandas as pd
import io

def generate_keys(df):
    df['key1'] = df['Invoice Date'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key2'] = df['Invoice Number'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key3'] = df['Invoice Number'].astype(str) + '_' + df['Invoice Date'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key4'] = df['Invoice Date'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str) + '_' + df['Invoice Number'].astype(str)
    return df

def check_match(row, master_df):
    if row['key4'] in master_df['key4'].values:
        return pd.Series(['Yes', 'Date+Amount+Supplier+Number'])
    elif row['key1'] in master_df['key1'].values:
        return pd.Series(['Yes', 'Date+Amount+Supplier'])
    elif row['key2'] in master_df['key2'].values:
        return pd.Series(['Yes', 'Number+Amount+Supplier'])
    elif row['key3'] in master_df['key3'].values:
        return pd.Series(['Yes', 'Number+Date+Supplier'])
    else:
        return pd.Series(['No', 'UNIQUE'])

st.set_page_config(page_title="Invoice Duplicate Checker", layout="centered")
st.title("Invoice Duplicate Checker")

# Load the master CSV from Access
try:
    master_df = pd.read_csv("data/access_table.csv")
    master_df['Invoice Date'] = pd.to_datetime(master_df['Invoice Date'], errors='coerce')
    master_df = master_df.dropna(subset=['Invoice Date'])
    master_df = master_df[['Invoice Number', 'Invoice Date', 'Gross Amount', 'Supplier Number']]
    master_df = generate_keys(master_df)
    st.success("Master data loaded successfully from CSV.")
except Exception as e:
    st.error(f"Failed to load master CSV file: {e}")
    st.stop()

# Upload invoice Excel file
invoice_file = st.file_uploader("Upload Invoice Excel File", type=["xlsx", "xls"])
if not invoice_file:
    st.info("Please upload an invoice file to check for duplicates.")
    st.stop()

# Process uploaded invoice file
try:
    raw_df = pd.read_excel(invoice_file)
    raw_df.columns = raw_df.columns.str.strip()
    df = raw_df[['Invoice Number', 'Invoice Date', 'Gross Amount', 'Supplier Number']].copy()
    df['Invoice Date'] = pd.to_datetime(df['Invoice Date'], errors='coerce')
    df = df.dropna(subset=['Invoice Date'])
    df = generate_keys(df)
    df[['Duplicate', 'Match Logic']] = df.apply(lambda row: check_match(row, master_df), axis=1)
except Exception as e:
    st.error(f"Failed to process invoice file: {e}")
    st.stop()

# Display results
st.success("Duplicate check complete.")
st.dataframe(df)

# Download results as Excel
excel_buffer = io.BytesIO()
final_df = raw_df.copy()
final_df['Duplicate'] = df['Duplicate']
final_df['Match Logic'] = df['Match Logic']
final_df.to_excel(excel_buffer, index=False)
st.download_button("📥 Download Results as Excel", data=excel_buffer.getvalue(), file_name="duplicates_report.xlsx")

# Show unique records
if "No" in df['Duplicate'].values:
    if st.button("🧩 Show UNIQUE Records"):
        unique_df = df[df['Duplicate'] == "No"][['Invoice Number', 'Invoice Date', 'Gross Amount', 'Supplier Number']]
        st.write("Records not found in the master data:")
        st.dataframe(unique_df)
else:
    st.warning("No unique records found.")
