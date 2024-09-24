import streamlit as st
import pandas as pd
import openpyxl
import plotly.express as px
from openpyxl import Workbook
import csv

# Function to convert CSV to OpenPyXL workbook
def csv_to_workbook(csv_file):
    # Read CSV into pandas DataFrame
    df = pd.read_csv(csv_file)
    
    # Create an OpenPyXL workbook
    wb = Workbook()
    ws = wb.active
    
    # Write DataFrame to workbook
    for r_idx, row in df.iterrows():
        ws.append(list(row))

    return wb

title = "Excel Visualizer"
st.set_page_config(
    page_title=title,
    page_icon=":bar_chart:",
    layout="wide"#,
#    primaryColor = "#E694FF",
#    backgroundColor = "#00172B",
#    secondaryBackgroundColor = "#0083B8",
#    textColor = "#FFF",
#    font = "sans serif"
)

st.header(title)

data_file = st.sidebar.file_uploader(
    "Select Excel file to visualize", 
    type=["xlsx", "xlsm", "xltx", "xltm","csv"], 
#    type=["xlsx", "xlsm", "xltx", "xltm"], 
    accept_multiple_files=False, 
    label_visibility="visible"
)

#print(data_file.name)
if data_file:
#    st.write("Filename: ", data_file.name)
#    st.write("Filename: ", type(data_file))
    if data_file.name.lower().endswith(".csv"):
        st.write("Process CSV file")
#        wb = Workbook()
#        ws = wb.active
#        with open(data_file, 'r') as f:
#            for row in csv.reader(f):
#                ws.append(row)
        wb = csv_to_workbook(data_file)
    else:
        wb = openpyxl.load_workbook(data_file)
        
#exit()

#if data_file:
    st.sidebar.markdown("""---""")
#    wb = openpyxl.load_workbook(data_file)

    ## Select sheet
    sheet_selector = st.sidebar.selectbox(
        "Select sheet to visualize",
        wb.sheetnames,
        index=None,
        placeholder="Select sheet..."
    )
    if sheet_selector:
        st.sidebar.markdown("""---""")
        df = pd.read_excel(data_file,sheet_selector)
        x_axis_selector = st.sidebar.selectbox("Select time column",
            df.columns,
            index=None,
            placeholder="Select time column..."
        )
        if x_axis_selector:
            y_axis_selector = st.sidebar.multiselect(
                "Select measures to display",
                list(filter(lambda x : x != x_axis_selector, df.columns))
            )
            if y_axis_selector:
                fig = px.line(
                    df, 
                    x=x_axis_selector, 
                    y=y_axis_selector
                )
                fig.update_layout(
                    yaxis=dict(rangemode='tozero',fixedrange=True),  # Fix the Y-axis starting at 0 and prevent zooming
                    showlegend=True,
                    xaxis=dict(fixedrange=False),  # Allow zooming on X-axis
                    plot_bgcolor="rgba(0,0,0,0)",
                    xaxis_title=None,
                    yaxis_title=None,
                    legend_title=None
                )
                st.plotly_chart(fig, use_container_width=True)
