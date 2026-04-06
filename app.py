from __future__ import annotations

from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

from excel_to_pdf.pdf_export import (
    SheetPdfOptions,
    build_pdf_bytes_from_sheets,
    build_student_marksheet_pdf_bytes_from_row,
    build_student_marksheets_pdf_bytes_from_df,
    build_student_table_profiles_pdf_bytes,
    read_excel_sheets,
)
from excel_to_pdf.db import DbConfig, load_sqlserver_odbc_connection_string, read_students_table


st.set_page_config(page_title="Excel → PDF", layout="wide")
st.title("Excel → PDF")
st.write("Upload an Excel file, preview the data, then generate a PDF.")

with st.expander("Export from SQL database (optional)", expanded=False):
    st.write("You can also load data from a database table and export it as a PDF.")
    db_type = st.selectbox("Database type", options=["sqlite", "sqlserver"], index=0)
    table_name = st.text_input("Table name", value="students")
    export_kind = st.selectbox("Export format", options=["Marksheet (subjects + marks)", "Student details (profile table)"], index=1)
    sqlite_path = None
    sqlserver_conn_str = None
    if db_type == "sqlite":
        sqlite_path = st.text_input("SQLite .db file path", value=r"D:\path\to\school.db")
    else:
        st.caption(
            "SQL Server uses `.streamlit/secrets.toml` ([sqlserver] connection_string) "
            "or environment variable `SQLSERVER_ODBC_CONN_STR`. Connection string is not entered in the UI."
        )
        # Old UI (removed): do not put real passwords in code. Kept only as reference.
        # sqlserver_conn_str = st.text_input(
        #     "SQL Server ODBC connection string",
        #     value=(
        #         "DRIVER={ODBC Driver 18 for SQL Server};SERVER=YOUR_SERVER;DATABASE=YOUR_DB;"
        #         "UID=YOUR_USER;PWD=YOUR_PASSWORD;Encrypt=yes;TrustServerCertificate=yes;"
        #     ),
        # )

    if st.button("Generate PDF from DB"):
        try:
            if db_type == "sqlserver":
                sqlserver_conn_str = load_sqlserver_odbc_connection_string()
            df = read_students_table(
                DbConfig(db_type=db_type, sqlite_path=sqlite_path, sqlserver_conn_str=sqlserver_conn_str),
                table=table_name.strip() or "student",
            )
            if export_kind.startswith("Marksheet"):
                pdf = build_student_marksheets_pdf_bytes_from_df(df, document_title="Student Marksheet")
                out_name = "students_marksheet_from_db.pdf"
            else:
                pdf = build_student_table_profiles_pdf_bytes(df, document_title="Student Details", name_column="Name")
                out_name = "students_details_from_db.pdf"
        except Exception as e:
            st.error(f"DB export failed: {e}")
            st.stop()

        st.download_button(
            "Download DB PDF",
            data=pdf,
            file_name=out_name,
            mime="application/pdf",
        )

with st.expander("Need a sample marksheet Excel file?", expanded=True):
    st.write("Download a ready-to-upload sample with student names and subject marks.")

    sample_df = pd.DataFrame(
        [
            {
                "Student Name": "Aarav Sharma",
                "English": 86,
                "Physics": 78,
                "Chemistry": 81,
                "History": 74,
                "Maths": 92,
                "Computer Science": 95,
            },
            {
                "Student Name": "Diya Patel",
                "English": 91,
                "Physics": 88,
                "Chemistry": 90,
                "History": 84,
                "Maths": 89,
                "Computer Science": 93,
            },
            {
                "Student Name": "Mohammed Khan",
                "English": 72,
                "Physics": 69,
                "Chemistry": 75,
                "History": 80,
                "Maths": 77,
                "Computer Science": 83,
            },
        ]
    )

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        sample_df.to_excel(writer, index=False, sheet_name="Marksheet")
    st.download_button(
        "Download sample Excel (.xlsx)",
        data=out.getvalue(),
        file_name="sample_students_marksheet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

uploaded = st.file_uploader("Upload Excel", type=["xlsx", "xlsm", "xls"])

if not uploaded:
    st.info("Upload an `.xlsx` file to begin.")
    st.stop()

excel_bytes = uploaded.getvalue()

try:
    xls = pd.ExcelFile(BytesIO(excel_bytes))
except Exception as e:
    st.error(f"Could not read Excel file: {e}")
    st.stop()

sheet_names = xls.sheet_names

with st.sidebar:
    st.subheader("Options")
    selected_sheets = st.multiselect(
        "Sheets to include in PDF",
        options=sheet_names,
        default=sheet_names[:1] if sheet_names else [],
    )
    per_student_pdf = st.checkbox(
        "Download one student's marksheet (PDF)",
        value=False,
        help="Uses the first selected sheet: one row per student. Pick a student name and download only that student's PDF.",
    )
    max_rows = st.number_input("Max rows per sheet (PDF)", min_value=1, max_value=5000, value=250, step=50)
    max_cols = st.number_input("Max columns per sheet (PDF)", min_value=1, max_value=200, value=20, step=5)

if not selected_sheets:
    st.warning("Select at least one sheet in the sidebar.")
    st.stop()

tabs = st.tabs([f"Preview: {name}" for name in selected_sheets])
for tab, name in zip(tabs, selected_sheets, strict=True):
    with tab:
        try:
            df = pd.read_excel(xls, sheet_name=name)
        except Exception as e:
            st.error(f"Failed to read sheet `{name}`: {e}")
            continue
        st.dataframe(df, use_container_width=True)

st.divider()

col1, col2 = st.columns([1, 2], vertical_alignment="center")
with col1:
    pdf_title = st.text_input("PDF title", value=f"{uploaded.name} export")
with col2:
    st.caption("Tip: if your sheet is very large, increase max rows/cols carefully (PDF generation time will grow).")

if st.button("Generate PDF", type="primary"):
    try:
        sheets = read_excel_sheets(excel_bytes, sheet_names=selected_sheets)
        pdf = build_pdf_bytes_from_sheets(
            sheets,
            document_title=pdf_title.strip() or "Excel export",
            options=SheetPdfOptions(
                title=pdf_title.strip() or "Excel export",
                max_rows=int(max_rows),
                max_cols=int(max_cols),
            ),
        )
    except Exception as e:
        st.error(f"PDF generation failed: {e}")
        st.stop()

    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    safe_base = (uploaded.name.rsplit(".", 1)[0] or "excel").replace(" ", "_")
    filename = f"{safe_base}_{ts}.pdf"

    st.success("PDF created.")
    st.download_button(
        "Download PDF",
        data=pdf,
        file_name=filename,
        mime="application/pdf",
    )

    if per_student_pdf:
        try:
            sheet_name = selected_sheets[0]
            df = pd.read_excel(xls, sheet_name=sheet_name)
            name_col = None
            for c in df.columns:
                if str(c).strip().lower() in {"student name", "name", "student"}:
                    name_col = str(c)
                    break
            if not name_col:
                raise ValueError("Could not find 'Student Name' column in the selected sheet.")

            student_names = [str(v).strip() for v in df[name_col].tolist() if str(v).strip()]
            if not student_names:
                raise ValueError("No student names found in the selected sheet.")

            chosen = st.selectbox("Select student", options=student_names, index=0)
            row = df[df[name_col].astype(str).str.strip() == chosen].head(1)
            if row.empty:
                raise ValueError("Selected student row not found.")

            subjects = [str(c) for c in df.columns if str(c) != name_col]
            subjects_and_marks = [(sub, None if pd.isna(row.iloc[0][sub]) else float(row.iloc[0][sub])) for sub in subjects]

            student_pdf = build_student_marksheet_pdf_bytes_from_row(
                student_name=chosen,
                subjects_and_marks=subjects_and_marks,
                document_title=pdf_title.strip() or "Student Marksheet",
            )
        except Exception as e:
            st.error(f"Student PDF generation failed: {e}")
            st.stop()

        st.download_button(
            "Download selected student's PDF",
            data=student_pdf,
            file_name=f"{safe_base}_{ts}_{chosen.replace(' ', '_')}.pdf",
            mime="application/pdf",
        )

