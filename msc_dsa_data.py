import streamlit as st
import docx
import pandas as pd
import numpy as np

st.set_page_config(page_title="MSC DSA Graduation Visualization", layout="wide")

#Loading the Word document
@st.cache_data
def load_data():
    return docx.Document("https://github.com/gathukalinet/msc_dsa/blob/main/SIMS%20Masters%20Graduation%20Report_Class%20of%202025_Draft%201%20(2).docx") 

doc = load_data()
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
    
