import streamlit as st
import docx
import pandas as pd
import numpy as np
import requests
from io import BytesIO

st.set_page_config(page_title="MSC DSA Graduation Visualization", layout="wide")

#Loading the Word document
@st.cache_data()
def download_docx():
    url = "https://github.com/gathukalinet/msc_dsa/raw/main/SIMS%20Masters%20Graduation%20Report_Class%20of%202025_Draft%201%20(2).docx"
    response = requests.get(url)
    response.raise_for_status()  # Ensures the download succeeded
    return docx.Document(BytesIO(response.content))

doc = docx.Document(BytesIO(download_docx()))
# Getting all tables
tables = doc.tables
print(f"Found {len(tables)} tables.")

st.title("MSC DSA Graduation and Backlog Visualization")
st.write("Welcome! This application visualizes the graduation and backlog data for the MSC DSA program.")

with st.sidebar:
    st.header("Navigation")
    st.write(" Choose a table to visualize:")

    table_options = [f"Table {i+1}" for i in range(len(tables))]
    selected_table = st.selectbox("Select a table:", table_options)
    
