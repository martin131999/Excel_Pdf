from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile
from typing import Iterable, Sequence

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


@dataclass(frozen=True)
class SheetPdfOptions:
    title: str
    max_rows: int = 250
    max_cols: int = 20


def _to_str_table(df: pd.DataFrame, *, max_rows: int, max_cols: int) -> list[list[str]]:
    view = df.copy()

    if view.shape[1] > max_cols:
        view = view.iloc[:, :max_cols].copy()
        view.columns = list(view.columns)[:-1] + [f"{view.columns[-1]} (truncated)"]

    if view.shape[0] > max_rows:
        view = view.iloc[:max_rows, :].copy()
        view.loc[view.index[-1], view.columns[0]] = f"{view.iloc[-1, 0]} (truncated)"

    view = view.fillna("")

    header = [str(c) for c in view.columns.tolist()]
    rows = [[str(v) for v in row] for row in view.to_numpy().tolist()]
    return [header, *rows]


def _looks_like_marksheet(df: pd.DataFrame) -> bool:
    if df is None or df.empty:
        return False
    cols = [str(c).strip().lower() for c in df.columns]
    has_student = any(c in {"student name", "name", "student"} for c in cols)
    return has_student and len(cols) >= 3


def _find_student_col(df: pd.DataFrame) -> str | None:
    for c in df.columns:
        key = str(c).strip().lower()
        if key in {"student name", "name", "student"}:
            return str(c)
    return None


def _subject_columns(df: pd.DataFrame, *, student_col: str) -> list[str]:
    subjects: list[str] = []
    for c in df.columns:
        if str(c) == student_col:
            continue
        subjects.append(str(c))
    return subjects


def _to_number(x: object) -> float | None:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if not s:
        return None
    try:
        return float(s)
    except Exception:
        return None


def _grade_from_percent(pct: float) -> str:
    if pct >= 90:
        return "A+"
    if pct >= 80:
        return "A"
    if pct >= 70:
        return "B"
    if pct >= 60:
        return "C"
    return "D"


def _add_marksheet_section(
    story: list[object],
    *,
    sheet_name: str,
    df: pd.DataFrame,
    styles,
    table_style: TableStyle,
) -> None:
    student_col = _find_student_col(df)
    if not student_col:
        story.append(Paragraph(f"Sheet: {sheet_name}", styles["Heading2"]))
        story.append(Spacer(1, 8))
        story.append(Paragraph("Could not find a 'Student Name' column.", styles["Normal"]))
        return

    subjects = _subject_columns(df, student_col=student_col)
    story.append(Paragraph("Student Marksheet", styles["Heading2"]))
    story.append(Paragraph(f"Sheet: {sheet_name}", styles["Normal"]))
    story.append(Spacer(1, 8))

    note = (
        "This marksheet shows subject-wise performance for each student. Marks are taken directly from the uploaded "
        "Excel file. Total and percentage are calculated across the listed subjects, and a simple grade is assigned "
        "for quick understanding. Please verify entries and keep this document for academic reference."
    )
    story.append(Paragraph(note, styles["BodyText"]))
    story.append(Spacer(1, 10))

    for idx, row in df.iterrows():
        student_name = str(row.get(student_col, "")).strip() or f"Student {idx + 1}"

        marks: list[tuple[str, float | None]] = []
        for sub in subjects:
            marks.append((sub, _to_number(row.get(sub))))

        nums = [m for _, m in marks if m is not None]
        total = float(sum(nums)) if nums else 0.0
        max_total = float(len(subjects) * 100) if subjects else 0.0
        pct = (total / max_total * 100.0) if max_total > 0 else 0.0
        grade = _grade_from_percent(pct)

        story.append(Paragraph(f"Student: {student_name}", styles["Heading3"]))

        table_data: list[list[str]] = [["Subject", "Marks (out of 100)"]]
        for sub, m in marks:
            table_data.append([sub, "" if m is None else (str(int(m)) if float(m).is_integer() else f"{m:g}")])
        table_data.append(["Total", f"{total:g} / {max_total:g}"])
        table_data.append(["Percentage", f"{pct:.2f}%"])
        table_data.append(["Grade", grade])

        t = Table(table_data, repeatRows=1, hAlign="LEFT")
        t.setStyle(table_style)
        story.append(t)
        story.append(Spacer(1, 14))

def build_student_marksheet_pdf_bytes_from_row(
    *,
    student_name: str,
    subjects_and_marks: Sequence[tuple[str, float | None]],
    document_title: str = "Student Marksheet",
    paragraph: str | None = None,
) -> bytes:
    """
    Build ONE student marksheet as ONE PDF (single page).
    """
    if paragraph is None:
        paragraph = (
            "This marksheet presents subject-wise marks and overall performance for the student. Totals and percentage are "
            "calculated from the provided subjects, and a simple grade is assigned for quick reference. Please review the "
            "uploaded Excel entries carefully; this document is generated automatically based on the data provided."
        )

    nums = [m for _, m in subjects_and_marks if m is not None]
    total = float(sum(nums)) if nums else 0.0
    max_total = float(len(subjects_and_marks) * 100) if subjects_and_marks else 0.0
    pct = (total / max_total * 100.0) if max_total > 0 else 0.0
    grade = _grade_from_percent(pct)

    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, title=document_title, author="excel-to-pdf")
    styles = getSampleStyleSheet()

    table_style = TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]
    )

    story: list[object] = []
    story.append(Paragraph(document_title, styles["Title"]))
    story.append(Spacer(1, 8))
    story.append(Paragraph(f"Student Name: {student_name}", styles["Heading2"]))
    story.append(Spacer(1, 8))
    story.append(Paragraph(paragraph, styles["BodyText"]))
    story.append(Spacer(1, 12))

    table_data: list[list[str]] = [["Subject", "Marks (out of 100)"]]
    for sub, m in subjects_and_marks:
        table_data.append([sub, "" if m is None else (str(int(m)) if float(m).is_integer() else f"{m:g}")])
    table_data.append(["Total", f"{total:g} / {max_total:g}"])
    table_data.append(["Percentage", f"{pct:.2f}%"])
    table_data.append(["Grade", grade])

    t = Table(table_data, repeatRows=1, hAlign="LEFT")
    t.setStyle(table_style)
    story.append(t)

    doc.build(story)
    return buf.getvalue()


def build_student_marksheets_pdf_bytes_from_df(
    students_df: pd.DataFrame,
    *,
    document_title: str = "Student Marksheet",
) -> bytes:
    """
    Build ONE PDF with many pages:
    - one page per student (one row per student)
    - same paragraph on every page
    """
    student_col = _find_student_col(students_df)
    if not student_col:
        raise ValueError("Data must contain a 'Student Name' (or 'Name' / 'Student') column.")

    subjects = _subject_columns(students_df, student_col=student_col)
    if not subjects:
        raise ValueError("Data must contain at least one subject column with marks.")

    paragraph = (
        "This marksheet presents subject-wise marks and overall performance for the student. Totals and percentage are "
        "calculated from the provided subjects, and a simple grade is assigned for quick reference. Please review the "
        "entries carefully; this document is generated automatically based on the data provided."
    )

    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, title=document_title, author="excel-to-pdf")
    styles = getSampleStyleSheet()

    table_style = TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]
    )

    story: list[object] = []
    for i, row in students_df.iterrows():
        if i > 0:
            story.append(PageBreak())

        student_name = str(row.get(student_col, "")).strip() or f"Student {i + 1}"

        marks: list[tuple[str, float | None]] = []
        for sub in subjects:
            marks.append((sub, _to_number(row.get(sub))))

        story.append(Paragraph(document_title, styles["Title"]))
        story.append(Spacer(1, 8))
        story.append(Paragraph(f"Student Name: {student_name}", styles["Heading2"]))
        story.append(Spacer(1, 8))
        story.append(Paragraph(paragraph, styles["BodyText"]))
        story.append(Spacer(1, 12))

        nums = [m for _, m in marks if m is not None]
        total = float(sum(nums)) if nums else 0.0
        max_total = float(len(subjects) * 100)
        pct = (total / max_total * 100.0) if max_total > 0 else 0.0
        grade = _grade_from_percent(pct)

        table_data: list[list[str]] = [["Subject", "Marks (out of 100)"]]
        for sub, m in marks:
            table_data.append([sub, "" if m is None else (str(int(m)) if float(m).is_integer() else f"{m:g}")])
        table_data.append(["Total", f"{total:g} / {max_total:g}"])
        table_data.append(["Percentage", f"{pct:.2f}%"])
        table_data.append(["Grade", grade])

        t = Table(table_data, repeatRows=1, hAlign="LEFT")
        t.setStyle(table_style)
        story.append(t)

    doc.build(story)
    return buf.getvalue()


def build_student_table_profiles_pdf_bytes(
    students_df: pd.DataFrame,
    *,
    document_title: str = "Student Details",
    name_column: str = "Name",
) -> bytes:
    """
    For DB tables like your `student` table:
    - One page per student
    - Shows fields like StudentID, Name, Phonenumber, Address, Gender, etc.
    - Same paragraph on every page
    """
    if students_df is None or students_df.empty:
        raise ValueError("No rows found in the student table.")

    paragraph = (
        "This document contains student details exported from the database. The information is generated automatically "
        "from the stored records and is intended for administrative use. Please review the data for accuracy and handle "
        "personal information responsibly according to your organization’s policy."
    )

    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, title=document_title, author="excel-to-pdf")
    styles = getSampleStyleSheet()

    table_style = TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]
    )

    cols = [str(c) for c in students_df.columns]
    story: list[object] = []
    for i, row in students_df.iterrows():
        if i > 0:
            story.append(PageBreak())

        name_val = row.get(name_column, None) if name_column in students_df.columns else None
        student_name = str(name_val).strip() if name_val is not None else ""
        if not student_name or student_name.lower() == "nan":
            student_name = f"Student {i + 1}"

        story.append(Paragraph(document_title, styles["Title"]))
        story.append(Spacer(1, 8))
        story.append(Paragraph(f"Name: {student_name}", styles["Heading2"]))
        story.append(Spacer(1, 8))
        story.append(Paragraph(paragraph, styles["BodyText"]))
        story.append(Spacer(1, 12))

        data: list[list[str]] = [["Field", "Value"]]
        for c in cols:
            v = row.get(c)
            if v is None or (isinstance(v, float) and pd.isna(v)):
                s = ""
            else:
                s = str(v)
            data.append([c, s])

        t = Table(data, repeatRows=1, hAlign="LEFT", colWidths=[140, 380])
        t.setStyle(table_style)
        story.append(t)

    doc.build(story)
    return buf.getvalue()


def build_student_marksheets_zip_bytes(
    students_df: pd.DataFrame,
    *,
    document_title: str = "Student Marksheet",
) -> bytes:
    """
    Build MANY PDFs (one per student) and return a ZIP file as bytes.
    """
    student_col = _find_student_col(students_df)
    if not student_col:
        raise ValueError("Marksheet sheet must contain a 'Student Name' (or 'Name' / 'Student') column.")
    subjects = _subject_columns(students_df, student_col=student_col)
    if not subjects:
        raise ValueError("Marksheet sheet must contain at least one subject column with marks.")

    zip_buf = BytesIO()
    with ZipFile(zip_buf, mode="w", compression=ZIP_DEFLATED) as zf:
        for idx, row in students_df.iterrows():
            student_name = str(row.get(student_col, "")).strip() or f"Student_{idx + 1}"
            safe_name = "".join(ch if ch.isalnum() or ch in (" ", "_", "-") else "_" for ch in student_name).strip()
            safe_name = safe_name.replace(" ", "_") or f"Student_{idx + 1}"

            subjects_and_marks: list[tuple[str, float | None]] = []
            for sub in subjects:
                subjects_and_marks.append((sub, _to_number(row.get(sub))))

            pdf_bytes = build_student_marksheet_pdf_bytes_from_row(
                student_name=student_name,
                subjects_and_marks=subjects_and_marks,
                document_title=document_title,
            )
            zf.writestr(f"{safe_name}.pdf", pdf_bytes)

    return zip_buf.getvalue()


def build_pdf_bytes_from_sheets(
    sheets: Sequence[tuple[str, pd.DataFrame]],
    *,
    document_title: str = "Excel export",
    options: SheetPdfOptions | None = None,
) -> bytes:
    """
    Build a single PDF containing one section per sheet.

    Notes:
    - Uses ReportLab (pure Python) for best Windows compatibility.
    - Large sheets are truncated to keep PDF generation fast and reliable.
    """
    if options is None:
        options = SheetPdfOptions(title=document_title)

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        title=document_title,
        author="excel-to-pdf",
    )
    styles = getSampleStyleSheet()

    story: list[object] = []
    story.append(Paragraph(document_title, styles["Title"]))
    story.append(Spacer(1, 12))

    table_style = TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]
    )

    for i, (sheet_name, df) in enumerate(sheets):
        if i > 0:
            story.append(PageBreak())

        if df is None or df.empty:
            story.append(Paragraph(f"Sheet: {sheet_name}", styles["Heading2"]))
            story.append(Spacer(1, 8))
            story.append(Paragraph("No rows found.", styles["Normal"]))
            continue

        if _looks_like_marksheet(df):
            _add_marksheet_section(story, sheet_name=sheet_name, df=df, styles=styles, table_style=table_style)
        else:
            story.append(Paragraph(f"Sheet: {sheet_name}", styles["Heading2"]))
            story.append(Spacer(1, 8))
            data = _to_str_table(df, max_rows=options.max_rows, max_cols=options.max_cols)
            table = Table(data, repeatRows=1)
            table.setStyle(table_style)
            story.append(table)

    doc.build(story)
    return buf.getvalue()


def read_excel_sheets(
    excel_bytes: bytes,
    *,
    sheet_names: Iterable[str] | None = None,
) -> list[tuple[str, pd.DataFrame]]:
    xls = pd.ExcelFile(BytesIO(excel_bytes))
    names = list(sheet_names) if sheet_names is not None else xls.sheet_names
    out: list[tuple[str, pd.DataFrame]] = []
    for name in names:
        df = pd.read_excel(xls, sheet_name=name)
        out.append((name, df))
    return out

