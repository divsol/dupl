import streamlit as st
import pandas as pd
import tempfile
import os
import pyodbc
from zipfile import ZipFile

st.title("MS Access to CSV Converter")

uploaded_file = st.file_uploader("Upload your .mdb or .accdb file", type=["mdb", "accdb"])

if uploaded_file is not None:
    with tempfile.NamedTemporaryFile(delete=False, suffix='.' + uploaded_file.name.split('.')[-1]) as tmp:
        tmp.write(uploaded_file.getbuffer())
        access_path = tmp.name

    try:
        # Connection string for Access DB
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={access_path};'
        )
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        table_names = [row.table_name for row in cursor.tables(tableType='TABLE')]

        if not table_names:
            st.error("No tables found in the database.")
        else:
            with tempfile.TemporaryDirectory() as tempdir:
                zip_path = os.path.join(tempdir, "tables_csv.zip")
                with ZipFile(zip_path, 'w') as zipf:
                    for table in table_names:
                        df = pd.read_sql(f"SELECT * FROM [{table}]", conn)
                        csv_filename = f"{table}.csv"
                        csv_path = os.path.join(tempdir, csv_filename)
                        df.to_csv(csv_path, index=False)
                        zipf.write(csv_path, arcname=csv_filename)

                with open(zip_path, "rb") as f:
                    st.download_button(
                        label="Download all tables as ZIP",
                        data=f,
                        file_name="tables_csv.zip",
                        mime="application/zip"
                    )
        cursor.close()
        conn.close()
    except Exception as e:
        st.error(f"Error processing the database: {e}")
    finally:
        os.remove(access_path)
