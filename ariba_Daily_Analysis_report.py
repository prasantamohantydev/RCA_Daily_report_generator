import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO, StringIO

st.set_page_config(page_title="Ariba Daily Report", layout="wide")
st.title("Ariba Daily Report Generator")

# ==================================
# Upload Files
# ==================================

open_file = st.file_uploader(
    "Upload Open Incident Report",
    type=["xls", "xlsx"]
)

resolved_file = st.file_uploader(
    "Upload Resolved Incident Report",
    type=["xls", "xlsx"]
)

# ==================================
# Read Report
# ==================================

def read_report_file(uploaded_file):

    uploaded_file.seek(0)

    header = uploaded_file.read(8)

    uploaded_file.seek(0)

    # Real XLS file
    if header.startswith(b'\xd0\xcf\x11\xe0'):
        return pd.read_excel(
            uploaded_file,
            engine="xlrd"
        )

    # XLSX file
    if header.startswith(b'PK'):
        return pd.read_excel(
            uploaded_file,
            engine="openpyxl",
            header=1
        )

    # HTML exported XLS
    raw = uploaded_file.read()

    try:
        html = raw.decode("utf-16")
    except:
        try:
            html = raw.decode("utf-8")
        except:
            html = raw.decode(
                "latin-1",
                errors="ignore"
            )

    tables = pd.read_html(
        StringIO(html)
    )

    if len(tables) > 1:
        return tables[1]

    return tables[0]

# ==================================
# SLA Parser
# ==================================

def sla_to_minutes(value):

    value = str(value)

    hrs = re.search(
        r'(\d+)\s*Hrs',
        value
    )

    mins = re.search(
        r'(\d+)\s*Mins',
        value
    )

    h = int(
        hrs.group(1)
    ) if hrs else 0

    m = int(
        mins.group(1)
    ) if mins else 0

    total = h * 60 + m

    if "(-" in value:
        total = -total

    return total

# ==================================
# Category Counter
# ==================================

def category_count(series, keyword):

    return (
        series.astype(str)
        .str.contains(
            keyword,
            case=False,
            na=False
        )
        .sum()
    )

# ==================================
# Main
# ==================================

if open_file and resolved_file:

    try:

        open_df = read_report_file(
            open_file
        )

        resolved_df = read_report_file(
            resolved_file
        )

        open_df.columns = (
            open_df.columns.astype(str)
            .str.strip()
        )

        resolved_df.columns = (
            resolved_df.columns.astype(str)
            .str.strip()
        )

        st.success(
            "Files loaded successfully"
        )

        # Debug
        st.write(
            "Open Columns:",
            open_df.columns.tolist()
        )

        st.write(
            "Resolved Columns:",
            resolved_df.columns.tolist()
        )

        logged_col = "Logged Time"
        priority_col = "Priority"
        category_col = "Category"
        sla_col = "Remaining SLA Time"

        open_df[logged_col] = pd.to_datetime(
            open_df[logged_col],
            errors="coerce"
        )

        resolved_df[logged_col] = pd.to_datetime(
            resolved_df[logged_col],
            errors="coerce"
        )

        report_date = max(
            open_df[logged_col].dt.date.max(),
            resolved_df[logged_col].dt.date.max()
        )

        open_today = open_df[
            open_df[logged_col].dt.date
            == report_date
        ]

        resolved_today = resolved_df[
            resolved_df[logged_col].dt.date
            == report_date
        ]

        todays_open_df = pd.concat(
            [open_today, resolved_today],
            ignore_index=True
        )

        latest_time = max(
            open_df[logged_col].max(),
            resolved_df[logged_col].max()
        )

        todays_open = len(
            todays_open_df
        )

        todays_closure = len(
            resolved_df
        )

        backlog = len(
            open_df
        )

        # SLA Meet
        resolved_sla = (
            resolved_df[sla_col]
            .apply(sla_to_minutes)
        )

        sla_meet = (
            resolved_sla > 0
        ).sum()

        # SLA Breach
        open_sla = (
            open_df[sla_col]
            .apply(sla_to_minutes)
        )

        sla_breach = (
        open_sla < 0
        ).sum()

        # Priorities
        p1_count = (
            resolved_df[priority_col]
            .astype(str)
            .str.upper()
            .str.contains("P1")
            .sum()
        )

        p2_count = (
            resolved_df[priority_col]
            .astype(str)
            .str.upper()
            .str.contains("P2")
            .sum()
        )

        p3_count = (
            resolved_df[priority_col]
            .astype(str)
            .str.upper()
            .str.contains("P3")
            .sum()
        )

        # Categories
        user_understanding = category_count(
        resolved_df[category_col],
        "User understanding"
        )

        data_fixing = category_count(
            resolved_df[category_col],
            "Data fixing"
        )

        master_data = category_count(
            resolved_df[category_col],
            "Master data"
        )

        interface_related = category_count(
            resolved_df[category_col],
            "Interface"
        )

        access = category_count(
            resolved_df[category_col],
            "Access"
        )

        application_bug = category_count(
            resolved_df[category_col],
            "Application bug"
        )

        repeated = category_count(
            resolved_df[category_col],
            "Repeated"
        )

        result = pd.DataFrame([{

            "Module": "Ariba",
            "Date": report_date.strftime("%d-%m-%Y"),
            "Time": latest_time.strftime("%H:%M:%S"),

            "Today's Open": todays_open,
            "Today's Closure": todays_closure,
            "Backlog": backlog,

            "SLA Meet Cases": int(sla_meet),
            "SLA Breach Cases": int(sla_breach),

            "P1": int(p1_count),
            "P2": int(p2_count),
            "P3": int(p3_count),

            "User Understanding": int(user_understanding),
            "Data Fixing": int(data_fixing),
            "Master Data": int(master_data),
            "Interface Related": int(interface_related),
            "Access": int(access),
            "Application Bug": int(application_bug),
            "Repeated": int(repeated)

        }])

        st.subheader(
            "Generated Report"
        )

        st.dataframe(result)

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            result.to_excel(
                writer,
                index=False
            )

        st.download_button(
            "Download Report",
            output.getvalue(),
            file_name=f"Ariba_Daily_Report_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(
            f"Error: {str(e)}"
        )

        