import streamlit as st
import docx
import pandas as pd
import numpy as np
import requests
from io import BytesIO
import matplotlib.pyplot as plt
import plotly.graph_objects as go

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


# Renaming all the dataframes for clarity
total_graduands = df1
total_graduands = total_graduands.iloc[:9, :9]
total_graduands = total_graduands.loc[:, ~total_graduands.columns.duplicated()]
total_graduands = total_graduands.replace('-', 0)
total_graduands.iloc[:, 1:] = total_graduands.iloc[:, 1:].apply(pd.to_numeric)
total_graduands = total_graduands.set_index('Class')


intake_gender_2023 = df2
intake_gender_2023 = intake_gender_2023.iloc[:4, :3]
intake_gender_2023 = intake_gender_2023.loc[:, ~intake_gender_2023.columns.duplicated()]
intake_gender_2023 = intake_gender_2023.set_index('Program')
intake_gender_2023 = intake_gender_2023.apply(pd.to_numeric, errors='coerce')

completion_rate_2022 = df3
completion_rate_2022 = completion_rate_2022.iloc[:4,]
completion_rate_2022 = completion_rate_2022.loc[:, ~completion_rate_2022.columns.duplicated()]
completion_rate_2022 = completion_rate_2022.set_index('Program')

completion_rate_2023 = df4
completion_rate_2023 = completion_rate_2023.iloc[:4,]
completion_rate_2023 = completion_rate_2023.loc[:, ~completion_rate_2023.columns.duplicated()]
completion_rate_2023 = completion_rate_2023.set_index('Program')

overall_graduands_2025 = df5
overall_graduands_2025 = overall_graduands_2025.iloc[:4, :]
overall_graduands_2025 = overall_graduands_2025.loc[:, ~overall_graduands_2025.columns.duplicated()]
overall_graduands_2025 = overall_graduands_2025.set_index('Program')

grads_2025_gender = df6
grads_2025_gender = grads_2025_gender.iloc[:4, :3]
grads_2025_gender = grads_2025_gender.loc[:, ~grads_2025_gender.columns.duplicated()]
grads_2025_gender = grads_2025_gender.set_index('Program')

msc_dsa_grad_list = df7
msc_dsa_grad_list = msc_dsa_grad_list.loc[:, ~msc_dsa_grad_list.columns.duplicated()]
msc_dsa_grad_list = msc_dsa_grad_list.loc[:, ~msc_dsa_grad_list.columns.duplicated()]
msc_dsa_grad_list = msc_dsa_grad_list.set_index('S/No.')

pending_students_2023 = df8
pending_students_2023 = pending_students_2023.iloc[:4, :5]
pending_students_2023 = pending_students_2023.loc[:, ~pending_students_2023.columns.duplicated()]
pending_students_2023 = pending_students_2023.set_index('Program')

df9.columns = df9.iloc[0, :].values
df9 = df9.iloc[1:]
msc_dsa_pending_current = df9
msc_dsa_pending_current = msc_dsa_pending_current.loc[:, ~msc_dsa_pending_current.columns.duplicated()]
msc_dsa_pending_current = msc_dsa_pending_current.set_index('S/No.')


backlog_2024 = df10
backlog_2024 = backlog_2024.loc[:, ~backlog_2024.columns.duplicated()]
backlog_2024 = backlog_2024.set_index('S/No.')


# Correcting df11 headers
df11 = df11.loc[:, ~df11.columns.duplicated()]
df11.loc[-1] = df11.columns           # Add current column names as a new row
df11.index = df11.index + 1          # Shift all existing rows down
df11 = df11.sort_index()
backlog_2023 = df11
backlog_2023 = backlog_2023.loc[:, ~backlog_2023.columns.duplicated()]



# Sidebar for table selection
st.sidebar.title("Choose Table to Visualize")
table_options = [
    "Table 1: Total Graduands",
    "Table 2: Intake by Gender (2023)",
    "Table 3: Completion Rate (2022)",
    "Table 4: Completion Rate (2023)",
    "Table 5: Overall Graduands (2025)",
    "Table 6: Graduands by Gender (2025)",
    "Table 7: MSC DSA Graduation List",
    "Table 8: Pending Students (2023)",
    "Table 9: MSC DSA Pending Current",
    "Table 10: MSC DSA Backlog (2024)",
    "Table 11: MSC DSA Backlog (2023)"]
selected_table = st.sidebar.selectbox("Choose a table to visualize:", table_options)



# Visualization based on selected table
if selected_table == "Table 1: Total Graduands":
    st.header("Total Graduands")
    st.write("This table shows the total number of graduands for each program in the MSC")
    st.write(total_graduands)
    
    # Visualization of total graduands
    st.subheader("Line Plot of Graduation Trends by MSc Program (2017–2025)")

    fig, ax = plt.subplots(figsize=(12, 6))

    fig = go.Figure()

    for col in total_graduands.columns:
        fig.add_trace(go.Scatter(
            x=total_graduands.index,
            y=total_graduands[col],
            mode='lines+markers',
            name=col
        ))

    fig.update_layout(
        xaxis_title='Year',
        yaxis_title='Number of Students',
        xaxis=dict(tickmode='linear'),
        hovermode='x unified',
        template='plotly_white',
        height=600
    )
# Display in Streamlit
    st.plotly_chart(fig, use_container_width=True)

elif selected_table == "Table 2: Intake by Gender (2023)":
    st.subheader("Intake by Gender (2023)")
    st.write(intake_gender_2023)
    fig = go.Figure()

# Add Male bars
    fig.add_trace(go.Bar(
        x=intake_gender_2023.index,
        y=intake_gender_2023['Male Students'],
        name='Male',
        marker_color='steelblue'
    ))

# Add Female bars
    fig.add_trace(go.Bar(
        x=intake_gender_2023.index,
        y=intake_gender_2023['Female Students'],
        name='Female',
        marker_color='lightcoral'
    ))

# Customize layout
    fig.update_layout(
        barmode='stack',  # Use 'stack' for stacked bars
        title='2023 Intake by Gender and Program',
        xaxis_title='Program',
        yaxis_title='Number of Students',
        template='plotly_white',
        height=500
    )

# Display in Streamlit
    st.plotly_chart(fig, use_container_width=True)

elif selected_table == "Table 3: Completion Rate (2022)":
    st.subheader("Completion Rate (2022)")
    st.write(completion_rate_2022)
elif selected_table == "Table 4: Completion Rate (2023)":
    st.subheader("Completion Rate (2023)")
    st.write(completion_rate_2023)
elif selected_table == "Table 5: Overall Graduands (2025)":
    st.subheader("Overall Graduands (2025)")
    st.write(overall_graduands_2025)
elif selected_table == "Table 6: Graduands by Gender (2025)":
    st.subheader("Graduands by Gender (2025)")
    st.write(grads_2025_gender)
elif selected_table == "Table 7: MSC DSA Graduation List":
    st.subheader("MSC DSA Graduation List")
    st.write(msc_dsa_grad_list)
elif selected_table == "Table 8: Pending Students (2023)":
    st.subheader("Pending Students (2023)")
    st.write(pending_students_2023)
elif selected_table == "Table 9: MSC DSA Pending Current":
    st.subheader("MSC DSA Pending Current")
    st.write(msc_dsa_pending_current)
elif selected_table == "Table 10: MSC DSA Backlog (2024)":
    st.subheader("MSC DSA Backlog (2024)")
    st.write(backlog_2024)
elif selected_table == "Table 11: MSC DSA Backlog (2023)":
    st.subheader("MSC DSA Backlog (2023)")
    st.write(backlog_2023)
else:
    st.subheader("No data available for this table.")
    st.write("Please select a valid table from the sidebar.")

program_options = ['MSc. Data Science and Analytics', 'MSc. Mathematical Finance & Risk Analysis', 'MSc. Statistical Science', 'MSc. Biomathematics']
selected_program = st.sidebar.selectbox("Choose a table to visualize:", program_options)
st.header(f"Overview for {selected_program}")

# 1. Total Graduands (Line chart for the selected program)
st.subheader("Graduation Trend")
years = total_graduands.columns.tolist()
values = total_graduands.loc[selected_program].tolist()

fig = go.Figure()
fig.add_trace(go.Scatter(x=years, y=values, mode='lines+markers', name=selected_program))
fig.update_layout(title="Graduation Trend (2017–2025)", xaxis_title="Year", yaxis_title="Number of Graduands")
st.plotly_chart(fig, use_container_width=True)

# 2. Gender Breakdown (if available)
if selected_program in intake_gender_2023.index:
    st.subheader("Intake Gender Breakdown (2023)")
    gender_data = intake_gender_2023.loc[selected_program]
    fig = go.Figure(data=[
        go.Bar(name='Gender', x=gender_data.index, y=gender_data.values)
    ])
    fig.update_layout(title="Intake by Gender", yaxis_title="Students")
    st.plotly_chart(fig)

# 3. Completion Rate
st.subheader("Completion Rate (2022 & 2023)")
for year, df in zip(['2022', '2023'], [completion_rate_2022, completion_rate_2023]):
    if selected_program in df['Program'].values:
        comp_row = df[df['Program'] == selected_program].iloc[0]
        st.markdown(f"**{year} Completion Rate:** {comp_row['Completion Rate']}")

# 4. Pending Students (if available)
st.subheader("Pending Students")
pending_dfs = [pending_students_2023, msc_dsa_pending_current, backlog_2024, backlog_2023]
for df in pending_dfs:
    matching_rows = df[df.iloc[:, 0].str.contains(selected_program, na=False)]
    if not matching_rows.empty:
        st.write(matching_rows)
