import streamlit as st
import docx
import pandas as pd
import numpy as np
import requests
from io import BytesIO

st.set_page_config(page_title="MSC DSA Graduation Visualization", layout="wide")

#Loading the Word document
@st.cache_resource()
def load_docx():
    url = "https://github.com/gathukalinet/msc_dsa/raw/main/SIMS%20Masters%20Graduation%20Report_Class%20of%202025_Draft%201%20(2).docx"
    response = requests.get(url)
    response.raise_for_status()
    return docx.Document(BytesIO(response.content))  # Load directly here

doc = load_docx()

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

def table_to_df(table):
    data = []
    for row in table.rows:
        data.append([cell.text.strip() for cell in row.cells])
    return pd.DataFrame(data)

# Looping through tables[1] to tables[12] and creating individual dataframes
for i in range(1, 12):
    df = table_to_df(tables[i])
    df.columns = df.iloc[0]        # Setting the first row as header
    df = df[1:]                    # Removing header row from data
    df.reset_index(drop=True, inplace=True)
    globals()[f'df{i}'] = df       # Creating. individual dataframes df1, df2, ..., df15

# Changing df9 to have correct headers


# Correcting df11 headers
df11 = df11.loc[:, ~df11.columns.duplicated()]
df11.loc[-1] = df11.columns           # Add current column names as a new row
df11.index = df11.index + 1          # Shift all existing rows down
df11 = df11.sort_index()

# Renaming all the dataframes for clarity
total_graduands = df1
total_graduands = total_graduands.iloc[:10, :10]

intake_gender_2023 = df2
intake_gender_2023 = intake_gender_2023.iloc[:4, :3]

completion_rate_2022 = df3
completion_rate_2022 = completion_rate_2022.iloc[:4,]

completion_rate_2023 = df4
completion_rate_2023 = completion_rate_2023.iloc[:4,]

overall_graduands_2025 = df5
overall_graduands_2025 = overall_graduands_2025.iloc[:4, :]

grads_2025_gender = df6
grads_2025_gender = grads_2025_gender.iloc[:4, :3]

msc_dsa_grad_list = df7
msc_dsa_grad_list = msc_dsa_grad_list.loc[:, ~msc_dsa_grad_list.columns.duplicated()]
msc_dsa_grad_list.set_index('S/No.', inplace=True)

pending_students_2023 = df8
pending_students_2023 = pending_students_2023.iloc[:4, :5]

df9.columns = df9.iloc[0, :].values
df9 = df9.iloc[1:]
msc_dsa_pending_current = df9
msc_dsa_pending_current.set_index('S/No.', inplace=True)

backlog_2024 = df10
backlog_2024.set_index('S/No.', inplace=True)

# Correcting df11 headers
df11 = df11.loc[:, ~df11.columns.duplicated()]
df11.loc[-1] = df11.columns           # Add current column names as a new row
df11.index = df11.index + 1          # Shift all existing rows down
df11 = df11.sort_index()
backlog_2023 = df11
backlog_2023.set_index('S/No.', inplace=True)

# Visualization based on selected table
if selected_table == "Table 1":
    st.subheader("Total Graduands")
    st.write(total_graduands)
# elif selected_table == "Table 2":
#     st.subheader("Intake by gender (2023)")
#     st.write(intake_gender_2023)
# elif selected_table == "Table 3":
#     st.subheader("Completion Rate (2022)")
#     st.write(completion_rate_2022)
# elif selected_table == "Table 4":
#     st.subheader("Completion Rate (2023)")
#     st.write(completion_rate_2023)
# elif selected_table == "Table 5":
#     st.subheader("Overall Graduands (2025)")
#     st.write(overall_graduands_2025)
# elif selected_table == "Table 6":
#     st.subheader("Graduands by gnder (2025)")
#     st.write(grads_2025_gender)
# elif selected_table == "Table 7":
#     st.subheader("MSC DSA Graduation List")
#     st.write(msc_dsa_grad_list)
# elif selected_table == "Table 8":
#     st.subheader("Pending Students (2023)")
#     st.write(pending_students_2023)
# elif selected_table == "Table 9":
#     st.subheader("MSC DSA Pending Current")
#     st.write(msc_dsa_pending_current)
# elif selected_table == "Table 10":
#     st.subheader("MSC DSA Backlog (2024)")
#     st.write(backlog_2024)
# elif selected_table == "Table 11":
#     st.subheader("MSC DSA Backlog (2023)")
#     st.write(backlog_2023)
# else:
#     st.subheader("No data available for this table.")
#     st.write("Please select a valid table from the sidebar.")