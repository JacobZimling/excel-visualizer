import streamlit as st
import pandas as pd
import openpyxl
import plotly.express as pltx
from dateutil import parser
from datetime import datetime

### App configuration
# Title for the app
title = "Excel Visualizer"

# List of delimiters so select from when importing csv files
csv_delimiters = (";",",")

# Control if "wave" duration will be shown on the graph
show_duration = True
#show_duration = False

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

# Diaplay delimiter selector
def csv_delimiter_selector(csv_delimiters):
    csv_delimiter = st.sidebar.radio("Select data delimiter",
        csv_delimiters,
        index=None
    )
    return csv_delimiter

# Display time column selector
def time_column_selector(df):
    x_axis_selector = st.sidebar.selectbox("Select time column",
        df.columns,
        index=None,
        placeholder="Select time column..."
    )
    return x_axis_selector

# Display measure columns selector
def measure_column_selector(df, exclude_column=''):
    y_axis_selector = st.sidebar.multiselect(
        "Select measures to display",
        list(filter(lambda x : x != exclude_column, df.columns))
    )
    return y_axis_selector

# Display line chart
def display_line_chart(df, time_column='', measures='', show_duration=False):
    df.drop_duplicates(inplace=True)

    fig = pltx.line(
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

    # Annotate "wave" duration to the graph
    if show_duration:
        # Identify waves
        wave_start = ''
        wave_end = ''

        isInWave = False
        waves = []

        for index, row in df.iterrows():


            if not isInWave:
                if int(row[measures[0]])>0:
                    isInWave = True
                    wave_start = row[time_column]

            if isInWave:
                if int(row[measures[0]])==0:
                    isInWave = False
                    wave_end = row[time_column]
                    waves.append({'start': wave_start, 'end': wave_end})

        for wave in waves:
            # Caculate wave duration
            wave_duration = parser.parse(wave['end']) - parser.parse(wave['start'])

            # Caclulate midpoint for duration text annotation
            text_x = parser.parse(wave['start']) + wave_duration / 2

            # Annotate wave duration indicator
            fig.add_shape(
                type="rect",
                x0=wave['start'], y0=-10, x1=wave['end'], y1=-20,  # Define the coordinates of the rectangle's corners
                line=dict(
                    color="rgb(0,131,184)",
                    width=2,
                ),
                fillcolor="rgb(0,131,184)",  # Set the color to fill the rectangle
            )
            # Annotate text for wave duration indicator
            text = f"{wave_duration.total_seconds()} sec"
            fig.add_annotation(
        #        x='2024-09-20T09:18:16.678', y=-15,  # Text annotation position
                x=text_x, y=-15,  # Text annotation position
                xref="x", yref="y",  # Coordinate reference system
                text=text,  # Text content
                showarrow=False  # Hide arrow
            )

    ## Display Line chart
    st.plotly_chart(fig, use_container_width=True)
    return

## File uploaded
if data_file:
    
    ## Process CSV file
    if data_file.name.lower().endswith(".csv"):
        ## Set Delimiter
        csv_delimiter = csv_delimiter_selector(csv_delimiters)

        if csv_delimiter:
            ## Read CSV file as Pandas dataframe
            df = pd.read_csv(data_file,sep=csv_delimiter)

            ## Display Column selector using dataframe columns for Time column selection
            st.sidebar.markdown("""---""")
            x_axis_selector = time_column_selector(df)
    
            ## Display Column selector using dataframe columns for Measure column selection. Exclude column chosen for Time
            if x_axis_selector:
                y_axis_selector = measure_column_selector(df, exclude_column=x_axis_selector)
    
                ## Configure Line chart using dataframe and selected columns
                ## Display Line chart
                if y_axis_selector:
                    display_line_chart(df, time_column=x_axis_selector, measures=y_axis_selector, show_duration=show_duration)
        
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
                    display_line_chart(df, time_column=x_axis_selector, measures=y_axis_selector, show_duration=show_duration)

## Footer
footer="""<style>
a:link , a:visited{
color: blue;
background-color: transparent;
text-decoration: underline;
}

a:hover,  a:active {
color: red;
background-color: transparent;
text-decoration: underline;
}

.footer {
position: fixed;
left: 0;
bottom: 0;
width: 100%;
background-color: white;
color: black;
text-align: center;
}
.footer p {
    margin-bottom: 0px;
}

</style>
<div class="footer">
<p>
    <b>Made with</b>: Python 3.11 <a style='text-align: center;' href="https://www.python.org/" target="_blank"><img style="width: 18px; height: 18px; margin: 0em;" src="https://i.imgur.com/ml09ccU.png"></a>
    and Streamlit <a style='text-align: center;' href="https://streamlit.io/" target="_blank"><img style="width: 24px; height: 25px; margin: 0em;" src="https://docs.streamlit.io/logo.svg"></a>
    by Jacob Zimling. <b>Hosted by:</b> Streamlit Community Cloud <a style='text-align: center;' href="https://streamlit.io/cloud" target="_blank"><img style="width: 24px; height: 25px; margin: 0em;" src="https://docs.streamlit.io/logo.svg"></a>
    
</p>
</div>
"""

st.markdown(footer,unsafe_allow_html=True)
