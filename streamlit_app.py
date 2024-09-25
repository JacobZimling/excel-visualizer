import streamlit as st
import pandas as pd
import openpyxl
import plotly.express as px
from openpyxl import Workbook
import csv

title = "Excel Visualizer"
csv_delimiters = (";",",")

st.set_page_config(
    page_title=title,
    page_icon=":bar_chart:",
    layout="wide"
)

## Set app title
st.header(title)

## Display file upload widget
#  Restrict to Excel and CSV file type
data_file = st.sidebar.file_uploader(
    "Select file to visualize. Excel and CSV supported.", 
    type=["xlsx", "xlsm", "xltx", "xltm","csv"], 
#    type=["xlsx", "xlsm", "xltx", "xltm"], 
    accept_multiple_files=False, 
    label_visibility="visible"
)

def csv_delimiter_selector(csv_delimiters):
    csv_delimiter = st.sidebar.selectbox("Select data delimiter",
        csv_delimiters,
        index=None,
        placeholder="Select data delimiter..."
    )
    return csv_delimiter

def time_column_selector(df):
    x_axis_selector = st.sidebar.selectbox("Select time column",
        df.columns,
        index=None,
        placeholder="Select time column..."
    )
    return x_axis_selector

def measure_column_selector(df, exclude_column=''):
    y_axis_selector = st.sidebar.multiselect(
        "Select measures to display",
        list(filter(lambda x : x != exclude_column, df.columns))
    )
    return y_axis_selector

def display_line_chart(df, time_column='', measures=''):
    fig = px.line(
        df, 
        x=time_column, 
        y=measures
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
    
    ## Display Line chart
    st.plotly_chart(fig, use_container_width=True)
    return

## File uploaded
if data_file:
    ## Set Delimiter
    csv_delimiter = csv_delimiter_selector(csv_delimiters)
    
    ## Process CSV file
    if data_file.name.lower().endswith(".csv"):
        ## Read CSV file as Pandas dataframe
        df = pd.read_csv(data_file,sep=csv_delimiter)
#        st.write(df)
        
        ## Display Column selector using dataframe columns for Time column selection
        st.sidebar.markdown("""---""")
        x_axis_selector = time_column_selector(df)

        ## Display Column selector using dataframe columns for Measure column selection. Exclude column chosen for Time
        if x_axis_selector:
            y_axis_selector = measure_column_selector(df, exclude_column=x_axis_selector)

            ## Configure Line chart using dataframe and selected columns
            ## Display Line chart
            if y_axis_selector:
                display_line_chart(df, time_column=x_axis_selector, measures=y_axis_selector)
        
    ## Process Excel file
    else:
        ## Read Excel file as openpyxl workbook object
        wb = openpyxl.load_workbook(data_file)

        ## Display dropdown to select sheet name
        st.sidebar.markdown("""---""")
        sheet_selector = st.sidebar.selectbox(
            "Select sheet to visualize",
            wb.sheetnames,
            index=None,
            placeholder="Select sheet..."
        )

        ## Read Excel sheet as Pandas dataframe
        if sheet_selector:
            st.sidebar.markdown("""---""")
            df = pd.read_excel(data_file,sheet_selector)

            ## Display Column selector using dataframe columns for Time column selection
            x_axis_selector = time_column_selector(df)

            ## Display Column selector using dataframe columns for Measure column selection. Exclude column chosen for Time
            if x_axis_selector:
                y_axis_selector = measure_column_selector(df, exclude_column=x_axis_selector)

                ## Configure Line chart using dataframe and selected columns
                ## Display Line chart
                if y_axis_selector:
                    display_line_chart(df, time_column=x_axis_selector, measures=y_axis_selector)
