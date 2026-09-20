import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.express as px
import openpyxl
import io
import os
import re
import zipfile


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="MSE & ESE Analysis",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# ============================================================
# PASSING CRITERIA
# ============================================================

ESE_PASS = 24          # 40% of 60
MSE_PASS = 8           # 40% of 20
INTERNAL_PASS = 16     # 40% of 40


# ============================================================
# TITLE
# ============================================================

st.title("📊 MSE & ESE Analysis Dashboard")

st.caption(
    "Academia ERP – Evaluation Method Wise Report"
)


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_text(value):

    if value is None:
        return ""

    return str(value).strip()


def clean_number(value):

    if value is None:
        return None

    text = str(value).strip()

    if text == "":
        return None

    text = text.replace(",", "")

    if text.upper() in [
        "AB",
        "ABS",
        "ABSENT",
        "NA",
        "N/A",
        "-"
    ]:
        return None

    try:
        return float(text)

    except:
        return None


def natural_sort_key(value):

    value = str(value)

    match = re.search(
        r"\d+",
        value
    )

    if match:

        return (
            re.sub(
                r"\d+",
                "",
                value
            ).lower(),
            int(match.group()),
            value.lower()
        )

    return (
        value.lower(),
        0,
        value.lower()
    )


# ============================================================
# READ EXCEL FILE
# ============================================================

def read_excel_file(uploaded_file):

    file_bytes = uploaded_file.getvalue()

    # --------------------------------------------------------
    # IMPORTANT:
    # Your ERP .xls file is actually an XLSX/ZIP file.
    # Therefore use openpyxl through BytesIO.
    # --------------------------------------------------------

    if file_bytes[:2] == b"PK":

        wb = openpyxl.load_workbook(
            io.BytesIO(file_bytes),
            data_only=True
        )

        return wb

    # --------------------------------------------------------
    # Old real XLS format
    # --------------------------------------------------------

    try:

        import xlrd

        workbook = xlrd.open_workbook(
            file_contents=file_bytes
        )

        sheet = workbook.sheet_by_index(0)

        rows = []

        for r in range(
            sheet.nrows
        ):

            rows.append(
                sheet.row_values(r)
            )

        return rows

    except Exception as e:

        raise ValueError(
            f"Unable to read ERP file: {e}"
        )


# ============================================================
# PROCESS XLSX ERP
# ============================================================

def process_xlsx(
    wb,
    filename
):

    ws = wb[
        wb.sheetnames[0]
    ]

    # --------------------------------------------------------
    # Convert worksheet to dataframe
    # --------------------------------------------------------

    raw = []

    for row in ws.iter_rows(
        values_only=True
    ):

        raw.append(
            list(row)
        )

    if len(raw) < 7:

        raise ValueError(
            "ERP report does not contain enough rows."
        )

    # ========================================================
    # ACTUAL STRUCTURE OF YOUR ERP FILE
    # ========================================================
    #
    # Row 1:
    # S.No.
    # Student Code
    # Admission Code
    # Student Name
    # First Name
    # Last Name
    # Academy Location Name
    # Program Name
    # CGPA
    # Batch Name
    # Semester 2
    #
    # Row 2:
    # Course
    #
    # Row 5:
    # External Theory (60)
    # Assignment (5)
    # Attendance (5)
    # MOOC/Test (5)
    # MSE (20)
    # Presentation/Quiz/Assignment-II (5)
    #
    # Row 7 onwards:
    # Student records
    # ========================================================

    # --------------------------------------------------------
    # HEADER ROW
    # --------------------------------------------------------

    header_row = raw[0]

    # --------------------------------------------------------
    # Find required columns
    # --------------------------------------------------------

    column_map = {}

    for i, value in enumerate(
        header_row
    ):

        text = clean_text(
            value
        ).lower()

        if text:

            column_map[
                text
            ] = i

    # --------------------------------------------------------
    # Program column
    # --------------------------------------------------------

    program_col = None

    for i, value in enumerate(
        header_row
    ):

        text = clean_text(
            value
        ).lower()

        if text in [
            "program name",
            "programme name"
        ]:

            program_col = i

            break

    if program_col is None:

        raise ValueError(
            "Program Name column not found."
        )

    # --------------------------------------------------------
    # Student columns
    # --------------------------------------------------------

    student_code_col = column_map.get(
        "student code"
    )

    student_name_col = column_map.get(
        "student name"
    )

    semester_col = column_map.get(
        "semester 2"
    )

    # --------------------------------------------------------
    # Course
    # --------------------------------------------------------

    course_text = ""

    # Course is in Row 2
    for value in raw[1]:

        text = clean_text(
            value
        )

        if (
            "-"
            in text
            and
            any(
                ch.isdigit()
                for ch in text
            )
        ):

            course_text = text

            break

    course_code = ""

    course_name = ""

    if course_text:

        match = re.match(
            r"^\s*([A-Za-z0-9]+)\s*[-–]\s*(.+)$",
            course_text
        )

        if match:

            course_code = (
                match.group(1)
                .strip()
            )

            course_name = (
                match.group(2)
                .strip()
            )

            # Remove report grade
            course_name = re.sub(
                r"\s*\([A-Za-z0-9+ ]+\)\s*$",
                "",
                course_name
            ).strip()

    # --------------------------------------------------------
    # Find marks columns
    # --------------------------------------------------------

    component_row = raw[4]

    ese_col = None
    assignment_col = None
    attendance_col = None
    mooc_col = None
    mse_col = None
    presentation_col = None

    for i, value in enumerate(
        component_row
    ):

        text = clean_text(
            value
        ).lower()

        if not text:
            continue

        if "external theory" in text:

            ese_col = i

        elif (
            text.startswith("assignment")
            and
            "presentation" not in text
        ):

            assignment_col = i

        elif "attendance" in text:

            attendance_col = i

        elif (
            "mooc" in text
            or
            "test" in text
        ):

            mooc_col = i

        elif text.startswith("mse"):

            mse_col = i

        elif (
            "presentation" in text
            or
            "quiz" in text
        ):

            presentation_col = i

    if ese_col is None:

        raise ValueError(
            "External Theory (60) column not found."
        )

    if mse_col is None:

        raise ValueError(
            "MSE (20) column not found."
        )

    # --------------------------------------------------------
    # Student data
    # --------------------------------------------------------

    records = []

    # Data starts at row 7
    for row in raw[6:]:

        if len(row) <= max(
            student_name_col,
            ese_col,
            mse_col
        ):

            continue

        student_name = clean_text(
            row[
                student_name_col
            ]
        )

        if not student_name:

            continue

        # ----------------------------------------------------
        # Skip totals / non-student rows
        # ----------------------------------------------------

        if student_name.lower() in [
            "total",
            "grand total",
            "student name"
        ]:

            continue

        # ----------------------------------------------------
        # Student Code
        # ----------------------------------------------------

        student_code = ""

        if student_code_col is not None:

            student_code = clean_text(
                row[
                    student_code_col
                ]
            )

        # ----------------------------------------------------
        # Program Name
        #
        # VERY IMPORTANT:
        # Take actual value from student row,
        # NOT the next header cell.
        # ----------------------------------------------------

        program_name = clean_text(
            row[
                program_col
            ]
        )

        # ----------------------------------------------------
        # Semester
        # ----------------------------------------------------

        semester = ""

        if semester_col is not None:

            semester = clean_text(
                row[
                    semester_col
                ]
            )

        # ----------------------------------------------------
        # ESE
        # ----------------------------------------------------

        ese_marks = clean_number(
            row[
                ese_col
            ]
        )

        # ----------------------------------------------------
        # MSE
        # ----------------------------------------------------

        mse_marks = clean_number(
            row[
                mse_col
            ]
        )

        # ----------------------------------------------------
        # Internal components
        # ----------------------------------------------------

        assignment = 0
        attendance = 0
        mooc = 0
        presentation = 0
        mse_internal = 0

        if assignment_col is not None:

            value = clean_number(
                row[
                    assignment_col
                ]
            )

            if value is not None:

                assignment = value

        if attendance_col is not None:

            value = clean_number(
                row[
                    attendance_col
                ]
            )

            if value is not None:

                attendance = value

        if mooc_col is not None:

            value = clean_number(
                row[
                    mooc_col
                ]
            )

            if value is not None:

                mooc = value

        if mse_marks is not None:

            mse_internal = mse_marks

        if presentation_col is not None:

            value = clean_number(
                row[
                    presentation_col
                ]
            )

            if value is not None:

                presentation = value

        # ----------------------------------------------------
        # INTERNAL = 40
        # ----------------------------------------------------

        internal_marks = (
            assignment
            +
            attendance
            +
            mooc
            +
            mse_internal
            +
            presentation
        )

        # ----------------------------------------------------
        # RESULTS
        #
        # Blank ESE = Fail/Absent
        # Blank MSE = Fail
        # ----------------------------------------------------

        ese_result = (
            "Pass"
            if ese_marks is not None
            and ese_marks >= ESE_PASS
            else "Fail"
        )

        mse_result = (
            "Pass"
            if mse_marks is not None
            and mse_marks >= MSE_PASS
            else "Fail"
        )

        internal_result = (
            "Pass"
            if internal_marks >= INTERNAL_PASS
            else "Fail"
        )

        records.append({

            "Program Name":
                program_name,

            "Semester":
                semester,

            "Course Code":
                course_code,

            "Course Name":
                course_name,

            "Student Code":
                student_code,

            "Student Name":
                student_name,

            "ESE Marks":
                ese_marks,

            "ESE Result":
                ese_result,

            "MSE Marks":
                mse_marks,

            "MSE Result":
                mse_result,

            "Internal Marks":
                internal_marks,

            "Internal Result":
                internal_result,

            "Source File":
                filename
        })

    if not records:

        raise ValueError(
            "No student records found."
        )

    return pd.DataFrame(
        records
    )


# ============================================================
# PROCESS OLD XLS
# ============================================================

def process_old_xls(
    rows,
    filename
):

    df = pd.DataFrame(
        rows
    )

    if df.empty:

        raise ValueError(
            "Empty XLS file."
        )

    # Same structure as ERP XLSX
    raw = df.values.tolist()

    header = raw[0]

    program_col = None
    student_code_col = None
    student_name_col = None
    semester_col = None

    for i, value in enumerate(
        header
    ):

        text = clean_text(
            value
        ).lower()

        if text in [
            "program name",
            "programme name"
        ]:

            program_col = i

        elif text == "student code":

            student_code_col = i

        elif text == "student name":

            student_name_col = i

        elif text.startswith(
            "semester"
        ):

            semester_col = i

    if program_col is None:

        raise ValueError(
            "Program Name column not found."
        )

    # Course
    course_text = ""

    for value in raw[1]:

        text = clean_text(
            value
        )

        if (
            "-"
            in text
            and
            any(
                x.isdigit()
                for x in text
            )
        ):

            course_text = text

            break

    course_code = ""
    course_name = ""

    match = re.match(
        r"^\s*([A-Za-z0-9]+)\s*[-–]\s*(.+)$",
        course_text
    )

    if match:

        course_code = match.group(1).strip()

        course_name = match.group(2).strip()

        course_name = re.sub(
            r"\s*\([A-Za-z0-9+ ]+\)\s*$",
            "",
            course_name
        )

    # Component row
    component_row = raw[4]

    ese_col = None
    assignment_col = None
    attendance_col = None
    mooc_col = None
    mse_col = None
    presentation_col = None

    for i, value in enumerate(
        component_row
    ):

        text = clean_text(
            value
        ).lower()

        if "external theory" in text:

            ese_col = i

        elif (
            text.startswith("assignment")
            and
            "presentation" not in text
        ):

            assignment_col = i

        elif "attendance" in text:

            attendance_col = i

        elif (
            "mooc" in text
            or
            "test" in text
        ):

            mooc_col = i

        elif text.startswith("mse"):

            mse_col = i

        elif (
            "presentation" in text
            or
            "quiz" in text
        ):

            presentation_col = i

    records = []

    for row in raw[6:]:

        if len(row) <= student_name_col:

            continue

        student_name = clean_text(
            row[
                student_name_col
            ]
        )

        if not student_name:

            continue

        program_name = clean_text(
            row[
                program_col
            ]
        )

        ese = clean_number(
            row[
                ese_col
            ]
        )

        mse = clean_number(
            row[
                mse_col
            ]
        )

        def get_mark(col):

            if col is None:
                return 0

            value = clean_number(
                row[col]
            )

            return (
                value
                if value is not None
                else 0
            )

        internal = (
            get_mark(assignment_col)
            +
            get_mark(attendance_col)
            +
            get_mark(mooc_col)
            +
            get_mark(mse_col)
            +
            get_mark(presentation_col)
        )

        records.append({

            "Program Name":
                program_name,

            "Semester":
                clean_text(
                    row[semester_col]
                )
                if semester_col is not None
                else "",

            "Course Code":
                course_code,

            "Course Name":
                course_name,

            "Student Code":
                clean_text(
                    row[
                        student_code_col
                    ]
                )
                if student_code_col is not None
                else "",

            "Student Name":
                student_name,

            "ESE Marks":
                ese,

            "ESE Result":
                "Pass"
                if ese is not None
                and ese >= ESE_PASS
                else "Fail",

            "MSE Marks":
                mse,

            "MSE Result":
                "Pass"
                if mse is not None
                and mse >= MSE_PASS
                else "Fail",

            "Internal Marks":
                internal,

            "Internal Result":
                "Pass"
                if internal >= INTERNAL_PASS
                else "Fail",

            "Source File":
                filename
        })

    if not records:

        raise ValueError(
            "No students found."
        )

    return pd.DataFrame(
        records
    )


# ============================================================
# MAIN FILE PROCESSOR
# ============================================================

def process_file(
    uploaded_file
):

    file_bytes = uploaded_file.getvalue()

    # --------------------------------------------------------
    # Your ERP .xls starts with PK = XLSX ZIP format
    # --------------------------------------------------------

    if file_bytes[:2] == b"PK":

        wb = openpyxl.load_workbook(
            io.BytesIO(file_bytes),
            data_only=True
        )

        return process_xlsx(
            wb,
            uploaded_file.name
        )

    # --------------------------------------------------------
    # True XLS
    # --------------------------------------------------------

    try:

        import xlrd

        book = xlrd.open_workbook(
            file_contents=file_bytes
        )

        sheet = book.sheet_by_index(0)

        rows = []

        for r in range(
            sheet.nrows
        ):

            rows.append(
                sheet.row_values(r)
            )

        return process_old_xls(
            rows,
            uploaded_file.name
        )

    except Exception as e:

        raise ValueError(
            f"Unable to process file: {e}"
        )


# ============================================================
# UPLOAD
# ============================================================

uploaded_files = st.file_uploader(
    "📂 Upload Academia ERP Evaluation Method Wise Report",
    type=[
        "xls",
        "xlsx"
    ],
    accept_multiple_files=True
)


# ============================================================
# INITIAL SCREEN
# ============================================================

if not uploaded_files:

    st.info(
        "Please upload the Academia ERP Evaluation Method Wise Report."
    )

    st.markdown(
        """
### Passing Criteria

| Component | Maximum | Passing |
|---|---:|---:|
| ESE / External Theory | 60 | **24** |
| MSE | 20 | **8** |
| Internal | 40 | **16** |
"""
    )

    st.stop()


# ============================================================
# PROCESS FILES
# ============================================================

all_data = []

errors = []

for uploaded_file in uploaded_files:

    try:

        result = process_file(
            uploaded_file
        )

        if (
            result is not None
            and
            not result.empty
        ):

            all_data.append(
                result
            )

    except Exception as e:

        errors.append(
            f"{uploaded_file.name}: {e}"
        )


# ============================================================
# ERROR DISPLAY
# ============================================================

if errors:

    with st.expander(
        "⚠️ Processing Errors",
        expanded=True
    ):

        for error in errors:

            st.error(
                error
            )


# ============================================================
# NO DATA
# ============================================================

if not all_data:

    st.error(
        "❌ No valid data could be extracted."
    )

    st.stop()


# ============================================================
# COMBINE
# ============================================================

data = pd.concat(
    all_data,
    ignore_index=True
)


# ============================================================
# REMOVE DUPLICATES
# ============================================================

data = data.drop_duplicates(
    subset=[
        "Program Name",
        "Semester",
        "Course Code",
        "Course Name",
        "Student Code",
        "Student Name"
    ]
)


# ============================================================
# FILTER
# ============================================================

show_filters = st.toggle(
    "🔎 Show Filters",
    value=False
)


selected_programs = sorted(
    data[
        "Program Name"
    ]
    .dropna()
    .unique()
    .tolist(),
    key=natural_sort_key
)

selected_semesters = sorted(
    data[
        "Semester"
    ]
    .dropna()
    .unique()
    .tolist(),
    key=natural_sort_key
)

selected_courses = sorted(
    data[
        "Course Name"
    ]
    .dropna()
    .unique()
    .tolist(),
    key=natural_sort_key
)


if show_filters:

    st.sidebar.header(
        "🔎 Filters"
    )

    selected_programs = st.sidebar.multiselect(
        "Program Name",
        selected_programs,
        default=selected_programs
    )

    selected_semesters = st.sidebar.multiselect(
        "Semester",
        selected_semesters,
        default=selected_semesters
    )

    selected_courses = st.sidebar.multiselect(
        "Course Name",
        selected_courses,
        default=selected_courses
    )


# ============================================================
# FILTER DATA
# ============================================================

filtered_data = data[
    data[
        "Program Name"
    ].isin(
        selected_programs
    )
    &
    data[
        "Semester"
    ].isin(
        selected_semesters
    )
    &
    data[
        "Course Name"
    ].isin(
        selected_courses
    )
].copy()


# ============================================================
# SUMMARY
# ============================================================

summary = (
    filtered_data
    .groupby(
        [
            "Program Name",
            "Course Code",
            "Course Name"
        ],
        dropna=False
    )
    .agg(

        ESE_Passed=(
            "ESE Result",
            lambda x:
                int(
                    (x == "Pass").sum()
                )
        ),

        ESE_Failed=(
            "ESE Result",
            lambda x:
                int(
                    (x == "Fail").sum()
                )
        ),

        MSE_Passed=(
            "MSE Result",
            lambda x:
                int(
                    (x == "Pass").sum()
                )
        ),

        MSE_Failed=(
            "MSE Result",
            lambda x:
                int(
                    (x == "Fail").sum()
                )
        ),

        Internal_Passed=(
            "Internal Result",
            lambda x:
                int(
                    (x == "Pass").sum()
                )
        ),

        Internal_Failed=(
            "Internal Result",
            lambda x:
                int(
                    (x == "Fail").sum()
                )
        ),

        Total_Students=(
            "Student Name",
            "count"
        )
    )
    .reset_index()
)


# ============================================================
# S.NO
# ============================================================

summary.insert(
    0,
    "S.NO",
    range(
        1,
        len(summary) + 1
    )
)


# ============================================================
# RENAME COLUMNS
# ============================================================

summary = summary.rename(
    columns={

        "ESE_Passed":
            "No of Students Passed ESE",

        "ESE_Failed":
            "No of Students Failed ESE",

        "MSE_Passed":
            "No of Students Passed MSE",

        "MSE_Failed":
            "No of Students Failed MSE",

        "Internal_Passed":
            "No of Students Passed Internal",

        "Internal_Failed":
            "No of Students Failed Internal",

        "Total_Students":
            "Total Students"
    }
)


# ============================================================
# EXACT OUTPUT COLUMN ORDER
# ============================================================

summary = summary[
    [
        "S.NO",

        "Program Name",

        "Course Code",

        "Course Name",

        "No of Students Passed ESE",

        "No of Students Failed ESE",

        "No of Students Passed MSE",

        "No of Students Failed MSE",

        "No of Students Passed Internal",

        "No of Students Failed Internal",

        "Total Students"
    ]
]


# ============================================================
# KPI
# ============================================================

st.subheader(
    "📌 Examination Overview"
)

students = filtered_data[
    "Student Code"
].replace(
    "",
    pd.NA
).dropna().nunique()

if students == 0:

    students = filtered_data[
        "Student Name"
    ].nunique()


ese_passed = int(
    (
        filtered_data[
            "ESE Result"
        ] == "Pass"
    ).sum()
)

ese_failed = int(
    (
        filtered_data[
            "ESE Result"
        ] == "Fail"
    ).sum()
)

mse_passed = int(
    (
        filtered_data[
            "MSE Result"
        ] == "Pass"
    ).sum()
)

mse_failed = int(
    (
        filtered_data[
            "MSE Result"
        ] == "Fail"
    ).sum()
)

internal_passed = int(
    (
        filtered_data[
            "Internal Result"
        ] == "Pass"
    ).sum()
)

internal_failed = int(
    (
        filtered_data[
            "Internal Result"
        ] == "Fail"
    ).sum()
)


c1, c2, c3, c4, c5 = st.columns(5)

c1.metric(
    "Students",
    f"{students:,}"
)

c2.metric(
    "ESE Passed",
    f"{ese_passed:,}"
)

c3.metric(
    "ESE Failed",
    f"{ese_failed:,}"
)

c4.metric(
    "MSE Failed",
    f"{mse_failed:,}"
)

c5.metric(
    "Internal Failed",
    f"{internal_failed:,}"
)


# ============================================================
# MAIN TABLE
# ============================================================

st.subheader(
    "📋 Course-wise MSE & ESE Analysis"
)

# ============================================================
# MAIN TABLE - RESPONSIVE FULL-WIDTH DISPLAY
# ============================================================

st.subheader(
    "📋 Course-wise MSE & ESE Analysis"
)

# ------------------------------------------------------------
# Keep the original full column names for CSV/calculation.
# Use shorter display headings so the complete table fits
# on the screen without horizontal dragging.
# ------------------------------------------------------------

display_summary = summary.copy()

def format_count_percentage(count, total):
    if pd.isna(count):
        return "-"
    count = int(count)
    total = int(total)

    if total == 0:
        return f"{count} (0.00%)"

    percentage = (count / total) * 100

    return f"{count} ({percentage:.2f}%)"


display_summary["No of Students Passed ESE"] = display_summary.apply(
    lambda row: format_count_percentage(
        row["No of Students Passed ESE"],
        row["Total Students"]
    ),
    axis=1
)

display_summary["No of Students Failed ESE"] = display_summary.apply(
    lambda row: format_count_percentage(
        row["No of Students Failed ESE"],
        row["Total Students"]
    ),
    axis=1
)

display_summary["No of Students Passed MSE"] = display_summary.apply(
    lambda row: format_count_percentage(
        row["No of Students Passed MSE"],
        row["Total Students"]
    ),
    axis=1
)

display_summary["No of Students Failed MSE"] = display_summary.apply(
    lambda row: format_count_percentage(
        row["No of Students Failed MSE"],
        row["Total Students"]
    ),
    axis=1
)

display_summary["No of Students Passed Internal"] = display_summary.apply(
    lambda row: format_count_percentage(
        row["No of Students Passed Internal"],
        row["Total Students"]
    ),
    axis=1
)

display_summary["No of Students Failed Internal"] = display_summary.apply(
    lambda row: format_count_percentage(
        row["No of Students Failed Internal"],
        row["Total Students"]
    ),
    axis=1
)

display_summary["Total Students"] = display_summary[
    "Total Students"
].apply(lambda x: f"{int(x):,}")


# ------------------------------------------------------------
# Short display headings
# ------------------------------------------------------------

display_summary = display_summary.rename(
    columns={
        "No of Students Passed ESE": "ESE Passed",
        "No of Students Failed ESE": "ESE Failed",
        "No of Students Passed MSE": "MSE Passed",
        "No of Students Failed MSE": "MSE Failed",
        "No of Students Passed Internal": "Internal Passed",
        "No of Students Failed Internal": "Internal Failed",
        "Total Students": "Total Students"
    }
)


# ------------------------------------------------------------
# HTML
# ------------------------------------------------------------

html_table = display_summary.to_html(
    index=False,
    escape=True,
    classes="marks-summary-table",
    border=0
)


table_html = f"""
<!DOCTYPE html>

<html>

<head>

<meta charset="UTF-8">

<style>

* {{
    box-sizing: border-box;
}}

html,
body {{
    margin: 0;
    padding: 0;
    width: 100%;
    background: white;
    font-family: Arial, sans-serif;
}}

#fullscreen-area {{
    width: 100%;
    background: white;
}}

.toolbar {{
    display: flex;
    justify-content: flex-end;
    margin-bottom: 6px;
}}

.fullscreen-btn {{
    border: 1px solid #d1d5db;
    background: white;
    border-radius: 6px;
    padding: 5px 10px;
    font-size: 12px;
    cursor: pointer;
    color: #374151;
}}

.fullscreen-btn:hover {{
    background: #f3f4f6;
}}


/* ---------------------------------------------------------
   TABLE
   --------------------------------------------------------- */

.table-container {{
    width: 100%;
    overflow: hidden;
    border: 1px solid #e5e7eb;
    border-radius: 8px;
}}

.marks-summary-table {{
    width: 100%;
    table-layout: fixed;
    border-collapse: collapse;
    background: white;
    font-size: 12px;
}}

.marks-summary-table th {{
    background: #f5f7fa;
    color: #5f6b7a;
    font-weight: 600;
    text-align: center !important;
    vertical-align: middle;
    padding: 7px 5px;
    border: 1px solid #e5e7eb;
    white-space: normal;
    line-height: 1.2;
}}

.marks-summary-table td {{
    padding: 8px 5px;
    border: 1px solid #e5e7eb;
    vertical-align: middle;
    white-space: normal;
    line-height: 1.25;
    overflow-wrap: anywhere;
}}


/* ---------------------------------------------------------
   COLUMN WIDTHS
   Total = 100%
   --------------------------------------------------------- */

.marks-summary-table th:nth-child(1),
.marks-summary-table td:nth-child(1) {{
    width: 4%;
    text-align: center !important;
}}

.marks-summary-table th:nth-child(2),
.marks-summary-table td:nth-child(2) {{
    width: 15%;
    text-align: left !important;
}}

.marks-summary-table th:nth-child(3),
.marks-summary-table td:nth-child(3) {{
    width: 7%;
    text-align: center !important;
}}

.marks-summary-table th:nth-child(4),
.marks-summary-table td:nth-child(4) {{
    width: 13%;
    text-align: left !important;
}}

.marks-summary-table th:nth-child(5),
.marks-summary-table td:nth-child(5),
.marks-summary-table th:nth-child(6),
.marks-summary-table td:nth-child(6),
.marks-summary-table th:nth-child(7),
.marks-summary-table td:nth-child(7),
.marks-summary-table th:nth-child(8),
.marks-summary-table td:nth-child(8),
.marks-summary-table th:nth-child(9),
.marks-summary-table td:nth-child(9),
.marks-summary-table th:nth-child(10),
.marks-summary-table td:nth-child(10) {{
    width: 9%;
    text-align: center !important;
}}

.marks-summary-table th:nth-child(11),
.marks-summary-table td:nth-child(11) {{
    width: 6%;
    text-align: center !important;
}}


/* ---------------------------------------------------------
   FULL SCREEN
   --------------------------------------------------------- */

#fullscreen-area:fullscreen {{
    width: 100vw;
    height: 100vh;
    padding: 20px;
    background: white;
    overflow: auto;
}}

#fullscreen-area:fullscreen .table-container {{
    overflow: visible;
}}

#fullscreen-area:fullscreen .marks-summary-table {{
    font-size: 14px;
}}

#fullscreen-area:fullscreen .marks-summary-table th {{
    padding: 10px 7px;
}}

#fullscreen-area:fullscreen .marks-summary-table td {{
    padding: 10px 7px;
}}

</style>

</head>

<body>

<div id="fullscreen-area">

<div class="toolbar">

<button
    class="fullscreen-btn"
    onclick="toggleFullscreen()"
    id="fullscreenButton"
>
    ⛶ Full Screen
</button>

</div>

<div class="table-container">

{html_table}

</div>

</div>


<script>

function toggleFullscreen() {{

    const element =
        document.getElementById("fullscreen-area");

    if (!document.fullscreenElement) {{

        if (element.requestFullscreen) {{
            element.requestFullscreen();
        }}

    }} else {{

        if (document.exitFullscreen) {{
            document.exitFullscreen();
        }}

    }}

}}


document.addEventListener(
    "fullscreenchange",
    function() {{

        const button =
            document.getElementById("fullscreenButton");

        if (document.fullscreenElement) {{

            button.innerHTML =
                "⛶ Exit Full Screen";

        }} else {{

            button.innerHTML =
                "⛶ Full Screen";

        }}

    }}
);

</script>

</body>

</html>
"""


components.html(
    table_html,
    height=155,
    scrolling=False
)

# ============================================================
# DOWNLOAD
# ============================================================

csv_data = summary.to_csv(
    index=False
).encode(
    "utf-8-sig"
)

st.download_button(
    "⬇️ Download Analysis CSV",
    data=csv_data,
    file_name="MSE_ESE_Analysis.csv",
    mime="text/csv"
)


# ============================================================
# CHART
# ============================================================

st.subheader(
    "📊 Failed Students by Course"
)

chart_data = summary[
    [
        "Course Name",
        "No of Students Failed ESE",
        "No of Students Failed MSE",
        "No of Students Failed Internal"
    ]
].copy()

chart_data = chart_data.rename(
    columns={

        "No of Students Failed ESE":
            "ESE",

        "No of Students Failed MSE":
            "MSE",

        "No of Students Failed Internal":
            "Internal"
    }
)

chart_long = chart_data.melt(
    id_vars=[
        "Course Name"
    ],
    value_vars=[
        "ESE",
        "MSE",
        "Internal"
    ],
    var_name="Component",
    value_name="Failed Students"
)

fig = px.bar(
    chart_long,
    x="Course Name",
    y="Failed Students",
    color="Component",
    barmode="group",
    text="Failed Students"
)

fig.update_traces(
    textposition="outside"
)

fig.update_layout(
    xaxis_title="Course Name",
    yaxis_title="Number of Failed Students",
    xaxis_tickangle=-45
)

st.plotly_chart(
    fig,
    use_container_width=True
)


# ============================================================
# FAILED STUDENTS
# ============================================================

st.subheader(
    "🚨 Failed Student Details"
)

failed_students = filtered_data[
    (
        filtered_data[
            "ESE Result"
        ] == "Fail"
    )
    |
    (
        filtered_data[
            "MSE Result"
        ] == "Fail"
    )
    |
    (
        filtered_data[
            "Internal Result"
        ] == "Fail"
    )
].copy()


failed_students = failed_students[
    [
        "Program Name",
        "Semester",
        "Course Code",
        "Course Name",
        "Student Code",
        "Student Name",
        "ESE Marks",
        "ESE Result",
        "MSE Marks",
        "MSE Result",
        "Internal Marks",
        "Internal Result"
    ]
]


st.dataframe(
    failed_students,
    use_container_width=True,
    hide_index=True,
    height=500
)


# ============================================================
# DOWNLOAD FAILED STUDENTS
# ============================================================

failed_csv = failed_students.to_csv(
    index=False
).encode(
    "utf-8-sig"
)

st.download_button(
    "⬇️ Download Failed Student Details",
    data=failed_csv,
    file_name="Failed_Student_Details.csv",
    mime="text/csv"
)