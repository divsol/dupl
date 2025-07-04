import streamlit as st
import tempfile
import urllib
from sqlalchemy import create_engine

def connect_access_db(access_db):
    cnnstr = (
        r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'
        f'DBQ={access_db};'
    )
    cnnurl = f"access+pyodbc:///?odbc_connect={urllib.parse.quote_plus(cnnstr)}"
    return create_engine(cnnurl)

st.title("MS Access to CSV Converter")
uploaded_file = st.file_uploader("Upload your .mdb or .accdb file", type=["mdb", "accdb"])

if uploaded_file is not None:
    with tempfile.NamedTemporaryFile(delete=False, suffix='.' + uploaded_file.name.split('.')[-1]) as tmp:
        tmp.write(uploaded_file.getbuffer())
        access_path = tmp.name

    try:
        engine = connect_access_db(access_path)
        # Now you can use engine to run queries, e.g.:
        # df = pd.read_sql("SELECT * FROM [YourTable]", engine)
        st.success("Connected successfully!")
    except Exception as e:
        st.error(f"Connection error: {e}")
