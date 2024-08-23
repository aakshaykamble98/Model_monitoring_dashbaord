import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO
from openpyxl import load_workbook
import os
import pickle
from openpyxl.utils.dataframe import dataframe_to_rows
import base64
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# from Code.PSI import create_powerpoint_download_button_PSI
# from Code.gini import create_ppt_download_button_gini
# from Code.Calibration import create_ppt_download_button_calibration

from PSI import create_powerpoint_download_button_PSI
from gini import create_ppt_download_button_gini
from Calibration import create_ppt_download_button_calibration

# Streamlit app for the Data module
def app():
    st.markdown(
        """
        <div style='background-color: #008080; border-radius: 5px; padding: 1px;'>
            <h1 style='text-align: center; font-size: 28px; color: white;'>PL - Scorecard Model Monitoring Data</h1>
        </div>

        <div>
            <p></p>
        </div>

        """,
        unsafe_allow_html=True)
    
    # Display the DataFrame without the index column    
    st.markdown(
                """
                <div style='text-align: center;
                font-size: 20px;'>
                    <strong>SUPPORT DATA</strong>
                </div>
                """,
                unsafe_allow_html=True)
    
    # Define the custom CSS for the download button
    custom_css = """
    <style>
        .download-icon {
            position: absolute;
            top: -23px;
            right: 70px;
            font-size: 22px;
            cursor: pointer;
            color: #007BFF; /* Change this to your desired icon color */
        }
        .download-icon:active {
        transform: scale(0.85);
        }
    </style>
    """
    st.markdown(custom_css, unsafe_allow_html=True)

    
    ppt_data_gini = ppt_data_calibration = ppt_data_psi = ppt_data_overview = ppt_data_change_log = ppt_data_summary = None
    
    if 'ppt_data_overview' in st.session_state:
        ppt_data_overview = st.session_state.ppt_data_overview

    #     st.download_button(
    #         label="Download Overview PowerPoint",
    #         data=ppt_data_overview,
    #         file_name='overview_presentation.pptx',
    #         mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
    #         help="Click here to download the Overview PowerPoint presentation",
    #         key='download_button_overview'
    #     )
    # else:
    #     st.write("No data available for download. Please run the Overview module first.")
        
    if 'ppt_data_change_log' in st.session_state:
        ppt_data_change_log = st.session_state.ppt_data_change_log
        
        # Ensure ppt_data_change_log is converted to BytesIO if necessary
        if isinstance(ppt_data_change_log, bytes):
            ppt_data_change_log = BytesIO(ppt_data_change_log)

    #     st.download_button(
    #         label="Download Change Log PowerPoint",
    #         data=ppt_data_change_log,
    #         file_name='change_log_presentation.pptx',
    #         mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
    #         help="Click here to download the Change Log PowerPoint presentation",
    #         key='download_button_change_log'
    #     )
    # else:
    #     st.write("No data available for download. Please run the Change Log module first.")

    if 'ppt_data_summary' in st.session_state:
        ppt_data_summary = st.session_state.ppt_data_summary
        
        # Ensure ppt_data_change_log is converted to BytesIO if necessary
        if isinstance(ppt_data_summary, bytes):
            ppt_data_summary = BytesIO(ppt_data_summary)

    #     st.download_button(
    #         label="Download Summary PowerPoint",
    #         data=ppt_data_summary,
    #         file_name='summary_presentation.pptx',
    #         mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
    #         help="Click here to download the Summary PowerPoint presentation",
    #         key='download_button_summary'
    #     )
    # else:
    #     st.write("No data available for download. Please run the Summary module first.")
    
    # Check if data for each module is available and create download buttons
    if 'df_gini' in st.session_state and 'fig_bytes' in st.session_state and 'thresholds_gini' in st.session_state:
        df_gini = st.session_state.df_gini
        fig_bytes = st.session_state.fig_bytes
        thresholds_gini = st.session_state.thresholds_gini
        data_comment_gini = st.session_state.get("data_comment_gini", "")
        graph_comment_gini = st.session_state.get("graph_comment_gini", "")

        ppt_data_gini = create_ppt_download_button_gini(df_gini, fig_bytes, thresholds_gini, data_comment_gini, graph_comment_gini)

    #     st.download_button(
    #         label="Download Gini PowerPoint",
    #         data=ppt_data_gini,
    #         file_name='gini_presentation.pptx',
    #         mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
    #         help="Click here to download the Gini PowerPoint presentation",
    #         key='download_button_gini'
    #     )
    # else:
    #     st.write("No data available for download. Please run the Gini module first.")
        
    if 'df_calibration' in st.session_state and 'fig_bytes_calibration' in st.session_state and 'thresholds_calibration' in st.session_state:
        df_calibration = st.session_state.df_calibration
        fig_bytes_calibration = st.session_state.fig_bytes_calibration
        thresholds_calibration = st.session_state.thresholds_calibration
        data_comment_calibration = st.session_state.get("data_comment_calibration", "")
        graph_comment_calibration = st.session_state.get("graph_comment_calibration", "")

        ppt_data_calibration = create_ppt_download_button_calibration(df_calibration, fig_bytes_calibration, thresholds_calibration, data_comment_calibration, graph_comment_calibration)

    #     st.download_button(
    #         label="Download Calibration PowerPoint",
    #         data=ppt_data_calibration,
    #         file_name='calibration_presentation.pptx',
    #         mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
    #         help="Click here to download the Calibration PowerPoint presentation",
    #         key='download_button_calibration'
    #     )
    # else:
    #     st.write("No data available for download. Please run the Calibration module first.")
        
    if 'df_psi' in st.session_state and 'fig1_bytes' in st.session_state and 'fig2_bytes' in st.session_state and 'thresholds_psi' in st.session_state:
        df_psi = st.session_state.df_psi
        fig1_bytes = st.session_state.fig1_bytes
        fig2_bytes = st.session_state.fig2_bytes
        thresholds_psi = st.session_state.thresholds_psi
        data_comment_psi = st.session_state.get("data_comment_psi", "")
        graph_comment_psi = st.session_state.get("graph_comment_psi", "")

        ppt_data_psi = create_powerpoint_download_button_PSI(df_psi, fig1_bytes, fig2_bytes, thresholds_psi, data_comment_psi, graph_comment_psi)

    #     st.download_button(
    #         label="Download PSI PowerPoint",
    #         data=ppt_data_psi,
    #         file_name='psi_presentation.pptx',
    #         mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
    #         help="Click here to download the PSI PowerPoint presentation",
    #         key='download_button_psi'
    #     )
    # else:
    #     st.write("No data available for download. Please run the PSI module first.")
    
    
    # Function to convert the DataFrame to Excel without the index and header
    def convert_df_to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, header=False, sheet_name='support')
        processed_data = output.getvalue()
        return processed_data
    
    # Read the Excel file without treating the first row as headers
    # Base directory of the current script
    base_dir = os.path.dirname(__file__)
    
    # Read the Excel file
    df = pd.read_excel(os.path.join(base_dir, 'Datasets', 'support_2.xlsx', header = None))  # Ensure the file is in the 'Datasets' directory
    
    # Replace None values with empty strings for better display
    df = df.fillna("")
    
    xlsx = convert_df_to_excel(df)
    
    # Encode the Excel file to base64
    b64 = base64.b64encode(xlsx).decode()
    
    # Help text
    help_text = "Click here to download the support data into excel"
    
    # Define the HTML for the download icon
    download_html = f'''
    <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="support_2.xlsx" title="{help_text}">
        <i class="download-icon">⬇️
    </a>
    '''
    
    #Hide download button given by streamlit by default
    st.markdown(
                """
                <style>
                [data-testid="stElementToolbar"] {
                    display: none;
                }
                </style>
                """,
                unsafe_allow_html=True
            )
    
    # Display the download icon
    st.markdown(download_html, unsafe_allow_html=True)
    
    st.dataframe(df, width=1100)
    
    # Ensure that you check if the button HTML is available in session state
    if 'all_ppt_button_html' in st.session_state:
        st.markdown(st.session_state.all_ppt_button_html, unsafe_allow_html=True) 
    else:
        st.write("No data available for download. Please run the all modules first.")
    
    # Define the custom CSS for the download button
    custom_css = """
    <style>
        .excel-download-button {
            position: absolute;
            top: -60px;
            left: 85px;
            cursor: pointer
        }
    </style>
    """
    st.markdown(custom_css, unsafe_allow_html=True)
    
    # Ensure that you check if the button HTML is available in session state
    if 'excel_button_html' in st.session_state:
        st.markdown(st.session_state.excel_button_html, unsafe_allow_html=True)
    else:
        st.write("No data available for download. Please run the all Overview and Change Log module first.")
    

if __name__ == "__main__":
    app()