
import streamlit as st
import pandas as pd
import os
from io import BytesIO
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Excel Consolidation Portal",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# ============================================================
# CUSTOM UI DESIGN
# ============================================================

st.markdown("""
<style>
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1250px;
    }

    .main-header {
        font-size: 30px;
        font-weight: 700;
        color: #17365D;
        margin-bottom: 5px;
    }

    .sub-header {
        color: #64748b;
        font-size: 14px;
        margin-bottom: 25px;
    }

    .section-heading {
        font-size: 19px;
        font-weight: 650;
        color: #17365D;
        margin-top: 15px;
    }

    div[data-testid="stMetric"] {
        background-color: #f5f8fc;
        padding: 16px;
        border-radius: 10px;
        border: 1px solid #e2e8f0;
    }

    .footer {
        text-align: center;
        color: #94a3b8;
        font-size: 12px;
        padding-top: 35px;
    }

    div.stButton > button {
        width: 100%;
        border-radius: 8px;
        height: 45px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
# HEADER
# ============================================================

st.markdown(
    '<div class="main-header">Excel Consolidation Portal</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="sub-header">'
    'Upload multiple Excel workbooks and consolidate their '
    'measurement data into one structured file.'
    '</div>',
    unsafe_allow_html=True
)

st.divider()


# ============================================================
# EXTRACTION FUNCTION
# ============================================================

def process_excel_file(uploaded_file):

    file_name = uploaded_file.name

    all_data = []
    logs = []

    try:

        file_bytes = uploaded_file.getvalue()

        sheets = pd.read_excel(
            BytesIO(file_bytes),
            sheet_name=None,
            header=None
        )

        for sheet_name, df in sheets.items():

            if df.empty:
                logs.append({
                    "File": file_name,
                    "Sheet": sheet_name,
                    "Status": "Skipped",
                    "Details": "Empty worksheet"
                })
                continue

            # Minimum 25 rows and 17 columns.
            # Column index 16 corresponds to Excel column Q.

            if df.shape[0] < 25 or df.shape[1] < 17:

                logs.append({
                    "File": file_name,
                    "Sheet": sheet_name,
                    "Status": "Skipped",
                    "Details": "Insufficient rows or columns"
                })
                continue

            # ------------------------------------------------
            # SPAN IDENTIFICATION
            # ------------------------------------------------

            span_start = df.iloc[2, 16]
            span_end = df.iloc[3, 16]

            span_start = (
                "" if pd.isna(span_start)
                else str(span_start).strip()
            )

            span_end = (
                "" if pd.isna(span_end)
                else str(span_end).strip()
            )

            span = f"{span_start}-{span_end}"

            # ------------------------------------------------
            # EXTRACT DATA: ROWS 8-25, COLUMNS D-I
            # ------------------------------------------------

            extracted = df.iloc[
                7:25,
                [3, 4, 5, 6, 7, 8]
            ].copy()

            extracted.columns = [
                "From Chainage",
                "From lat/long",
                "To Chainage",
                "To lat/long",
                "Length",
                "Depth"
            ]

            # Add span information

            extracted["Span"] = span

            # Add source information

            extracted["Source File"] = file_name
            extracted["Source Sheet"] = sheet_name

            # Remove rows where all six measurement
            # fields are blank

            extracted = extracted.dropna(
                subset=[
                    "From Chainage",
                    "From lat/long",
                    "To Chainage",
                    "To lat/long",
                    "Length",
                    "Depth"
                ],
                how="all"
            )

            # Also remove rows containing only whitespace

            
            measurement_columns = [
                "From Chainage",
                "From lat/long",
                "To Chainage",
                "To lat/long",
                "Length",
                "Depth"
            ]

            def is_blank(value):
                return (
                    pd.isna(value)
                    or (isinstance(value, str) and not value.strip())
                )

            valid_rows = extracted.apply(
                lambda row: not all(
                    is_blank(value)
                    for value in row[measurement_columns]
                ),
                axis=1
            )

            extracted = extracted.loc[valid_rows].copy()


            if extracted.empty:

                logs.append({
                    "File": file_name,
                    "Sheet": sheet_name,
                    "Status": "Skipped",
                    "Details": "No measurement records found"
                })

                continue

            all_data.append(extracted)

            logs.append({
                "File": file_name,
                "Sheet": sheet_name,
                "Status": "Processed",
                "Details": f"{len(extracted)} records | Span: {span}"
            })

    except Exception as e:

        logs.append({
            "File": file_name,
            "Sheet": "-",
            "Status": "Error",
            "Details": str(e)
        })

    return all_data, logs


# ============================================================
# EXCEL OUTPUT GENERATION
# ============================================================

def create_output_excel(final_df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        final_df.to_excel(
            writer,
            index=False,
            sheet_name="Consolidated Data"
        )

        ws = writer.sheets["Consolidated Data"]

        # Header formatting

        for cell in ws[1]:

            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )

            cell.fill = PatternFill(
                start_color="17365D",
                end_color="17365D",
                fill_type="solid"
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        # Freeze header row

        ws.freeze_panes = "A2"

        # Enable filters

        ws.auto_filter.ref = ws.dimensions

        # Set column widths

        widths = {
            "A": 25,
            "B": 18,
            "C": 25,
            "D": 18,
            "E": 25,
            "F": 15,
            "G": 15,
            "H": 35,
            "I": 25
        }

        for col, width in widths.items():
            ws.column_dimensions[col].width = width

        ws.row_dimensions[1].height = 25

    output.seek(0)

    return output


# ============================================================
# FILE UPLOAD SECTION
# ============================================================

st.markdown(
    '<div class="section-heading">1. Upload Excel Files</div>',
    unsafe_allow_html=True
)

uploaded_files = st.file_uploader(
    "Select Excel workbooks to consolidate",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    help="You can select multiple Excel files at once."
)

if uploaded_files:

    st.success(
        f"{len(uploaded_files)} file(s) selected successfully."
    )

    file_details = pd.DataFrame([
        {
            "File Name": file.name,
            "Size (KB)": round(len(file.getvalue()) / 1024, 2)
        }
        for file in uploaded_files
    ])

    with st.expander("View selected files"):

        st.dataframe(
            file_details,
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# PROCESSING SECTION
# ============================================================

st.markdown(
    '<div class="section-heading">2. Consolidate Data</div>',
    unsafe_allow_html=True
)

if st.button(
    "🚀 Consolidate Files",
    type="primary",
    disabled=not uploaded_files
):

    all_data = []
    all_logs = []

    progress = st.progress(0)
    status = st.empty()

    total_files = len(uploaded_files)

    for index, uploaded_file in enumerate(uploaded_files):

        status.info(
            f"Processing {uploaded_file.name}..."
        )

        file_data, file_logs = process_excel_file(
            uploaded_file
        )

        all_data.extend(file_data)
        all_logs.extend(file_logs)

        progress.progress(
            (index + 1) / total_files
        )

    if all_data:

        final_df = pd.concat(
            all_data,
            ignore_index=True
        )

        # Final column arrangement

        final_df = final_df[
            [
                "Span",
                "From Chainage",
                "From lat/long",
                "To Chainage",
                "To lat/long",
                "Length",
                "Depth",
                "Source File",
                "Source Sheet"
            ]
        ]

        output_file = create_output_excel(final_df)

        st.session_state["final_df"] = final_df
        st.session_state["output_file"] = output_file.getvalue()
        st.session_state["logs"] = all_logs
        st.session_state["processed"] = True

    else:

        st.session_state["final_df"] = None
        st.session_state["output_file"] = None
        st.session_state["logs"] = all_logs
        st.session_state["processed"] = False

    status.empty()
    progress.empty()


# ============================================================
# RESULTS SECTION
# ============================================================

if "logs" in st.session_state:

    st.divider()

    st.markdown(
        '<div class="section-heading">3. Processing Results</div>',
        unsafe_allow_html=True
    )

    logs_df = pd.DataFrame(
        st.session_state["logs"]
    )

    processed_count = (
        len(logs_df[logs_df["Status"] == "Processed"])
        if not logs_df.empty else 0
    )

    skipped_count = (
        len(logs_df[logs_df["Status"] == "Skipped"])
        if not logs_df.empty else 0
    )

    error_count = (
        len(logs_df[logs_df["Status"] == "Error"])
        if not logs_df.empty else 0
    )

    col1, col2, col3 = st.columns(3)

    col1.metric("Processed Sheets", processed_count)
    col2.metric("Skipped Sheets", skipped_count)
    col3.metric("Errors", error_count)

    st.markdown("**Processing Log**")

    if not logs_df.empty:

        st.dataframe(
            logs_df,
            use_container_width=True,
            hide_index=True
        )

    if st.session_state["processed"]:

        final_df = st.session_state["final_df"]

        st.success("Consolidation completed successfully!")

        st.markdown("**Consolidated Data Preview**")

        st.metric(
            "Total Consolidated Records",
            len(final_df)
        )

        st.dataframe(
            final_df.head(100),
            use_container_width=True,
            hide_index=True
        )

        st.caption(
            "Showing the first 100 records. "
            "All records are included in the download."
        )

        st.markdown(
            '<div class="section-heading">4. Download Output</div>',
            unsafe_allow_html=True
        )

        st.download_button(
            label="⬇ Download Consolidated Excel",
            data=st.session_state["output_file"],
            file_name="consolidated_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    else:

        st.warning(
            "No valid measurement data was found. "
            "Please check the uploaded files and processing log."
        )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    '<div class="footer">'
    'Excel Consolidation Portal | Internal Team Utility'
    '</div>',
    unsafe_allow_html=True
)
