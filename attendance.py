import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.express as px
import xml.etree.ElementTree as ET
import re
import html


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Student Attendance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# ============================================================
# TITLE
# ============================================================

st.title("📊 Student Attendance Dashboard")

st.caption(
    "Academia ERP Attendance Analysis"
)


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def local_name(tag):
    """
    Get XML tag name without namespace.
    """
    return tag.split("}")[-1].lower()


def make_unique_columns(columns):
    """
    Make blank and duplicate column names unique.
    """

    result = []
    counts = {}

    for col in columns:

        col = str(col).strip()

        if col == "":
            col = "Unnamed"

        if col not in counts:

            counts[col] = 0
            result.append(col)

        else:

            counts[col] += 1

            result.append(
                f"{col}_{counts[col]}"
            )

    return result


def normalize_text(value):
    """
    Normalize text for column matching.
    """

    value = str(value).strip().lower()

    value = value.replace(
        "_",
        " "
    )

    value = value.replace(
        "-",
        " "
    )

    value = value.replace(
        "/",
        " "
    )

    value = re.sub(
        r"\s+",
        " ",
        value
    )

    return value.strip()


def get_cell_text(cell):
    """
    Extract all text from an XML cell,
    including nested XML elements.
    """

    return "".join(
        cell.itertext()
    ).strip()


# ============================================================
# NATURAL SORT
# ============================================================

def natural_sort_value(value):

    value = str(value).strip()

    # --------------------------------------------------------
    # Semester sorting
    # --------------------------------------------------------

    semester_match = re.search(
        r"(?:semester|sem)\s*(\d+)",
        value,
        flags=re.IGNORECASE
    )

    if semester_match:

        semester_number = int(
            semester_match.group(1)
        )

        text_part = re.sub(
            r"(?:semester|sem)\s*\d+",
            "",
            value,
            flags=re.IGNORECASE
        ).strip().lower()

        return (
            f"{text_part}|"
            f"{semester_number:010d}|"
            f"{value.lower()}"
        )

    # --------------------------------------------------------
    # General numeric sorting
    # --------------------------------------------------------

    number_match = re.search(
        r"\d+",
        value
    )

    if number_match:

        number = int(
            number_match.group()
        )

        text_part = re.sub(
            r"\d+",
            "",
            value
        ).strip().lower()

        return (
            f"{text_part}|"
            f"{number:010d}|"
            f"{value.lower()}"
        )

    return (
        f"{value.lower()}|"
        f"0000000000|"
        f"{value.lower()}"
    )


# ============================================================
# EXTRACT SCHOOL NAME
# ============================================================

def extract_school_name(root):

    # --------------------------------------------------------
    # Search individual XML rows
    # --------------------------------------------------------

    for row in root.iter():

        if local_name(
            row.tag
        ) != "row":

            continue

        cells = []

        for cell in row:

            if local_name(
                cell.tag
            ) != "cell":

                continue

            cells.append(
                get_cell_text(cell)
            )

        # ----------------------------------------------------
        # Search for School Name
        # ----------------------------------------------------

        for i, text in enumerate(cells):

            text_lower = text.lower()

            if "school name" in text_lower:

                # Example:
                # School Name : GNA Business School

                if ":" in text:

                    value = text.split(
                        ":",
                        1
                    )[1].strip()

                    if value:

                        return value

                # Example:
                # School Name
                # GNA Business School

                for next_value in cells[
                    i + 1:
                ]:

                    next_value = str(
                        next_value
                    ).strip()

                    if next_value:

                        return next_value

    # --------------------------------------------------------
    # Fallback: complete XML text
    # --------------------------------------------------------

    full_text = " ".join(
        text.strip()
        for text in root.itertext()
        if text and text.strip()
    )

    patterns = [

        r"School\s*Name\s*:\s*(.+?)(?=\s+(?:Program|Programme|Period|Semester|Report)\b|$)",

        r"School\s*Name\s+(.+?)(?=\s+(?:Program|Programme|Period|Semester|Report)\b|$)"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            full_text,
            flags=re.IGNORECASE
        )

        if match:

            value = match.group(
                1
            ).strip()

            if value:

                return value

    return "Not Available"


# ============================================================
# READ ACADEMIA ERP XML / XLS
# ============================================================

def read_academia_xml(file_bytes):

    try:

        root = ET.fromstring(
            file_bytes
        )

    except Exception as e:

        raise ValueError(
            f"Unable to read Academia XML file: {e}"
        )

    # ========================================================
    # SCHOOL NAME
    # ========================================================

    school_name = extract_school_name(
        root
    )

    # ========================================================
    # FIND WORKSHEETS
    # ========================================================

    worksheets = []

    for element in root.iter():

        if local_name(
            element.tag
        ) == "worksheet":

            worksheets.append(
                element
            )

    if not worksheets:

        raise ValueError(
            "No worksheet found in Academia ERP file."
        )

    # ========================================================
    # FIND TABLES
    # ========================================================

    candidate_tables = []

    for worksheet in worksheets:

        for element in worksheet:

            if local_name(
                element.tag
            ) != "table":

                continue

            rows = []

            # ------------------------------------------------
            # Read XML rows
            # ------------------------------------------------

            for row in element:

                if local_name(
                    row.tag
                ) != "row":

                    continue

                row_values = []

                current_index = 1

                # --------------------------------------------
                # Read cells
                # --------------------------------------------

                for cell in row:

                    if local_name(
                        cell.tag
                    ) != "cell":

                        continue

                    # ----------------------------------------
                    # SpreadsheetML ss:Index
                    # ----------------------------------------

                    cell_index = None

                    for attr_name, attr_value in cell.attrib.items():

                        if (
                            attr_name
                            .split("}")[-1]
                            .lower()
                            == "index"
                        ):

                            try:

                                cell_index = int(
                                    attr_value
                                )

                            except:

                                cell_index = None

                    # ----------------------------------------
                    # Fill skipped cells
                    # ----------------------------------------

                    if cell_index is not None:

                        while (
                            current_index
                            <
                            cell_index
                        ):

                            row_values.append(
                                ""
                            )

                            current_index += 1

                    # ----------------------------------------
                    # Cell value
                    # ----------------------------------------

                    row_values.append(
                        get_cell_text(cell)
                    )

                    current_index += 1

                if row_values:

                    rows.append(
                        row_values
                    )

            if rows:

                candidate_tables.append(
                    rows
                )

    if not candidate_tables:

        raise ValueError(
            "No table found in Academia ERP XML file."
        )

    # ========================================================
    # FIND ACTUAL STUDENT ATTENDANCE HEADER
    # ========================================================

    selected_rows = None
    selected_header_index = None

    best_score = -1

    for rows in candidate_tables:

        for i, row in enumerate(rows):

            row_text = " ".join(
                normalize_text(x)
                for x in row
            )

            score = 0

            # ------------------------------------------------
            # Student
            # ------------------------------------------------

            if (
                "student" in row_text
                or
                "student name" in row_text
            ):

                score += 3

            # ------------------------------------------------
            # Program
            # ------------------------------------------------

            if (
                "program" in row_text
                or
                "programme" in row_text
            ):

                score += 4

            # ------------------------------------------------
            # Period / Semester
            # ------------------------------------------------

            if (
                "period" in row_text
                or
                "semester" in row_text
                or
                re.search(
                    r"\bsem\b",
                    row_text
                )
            ):

                score += 4

            # ------------------------------------------------
            # Attendance
            # ------------------------------------------------

            if (
                "attendance" in row_text
                or
                "percentage" in row_text
                or
                "present" in row_text
            ):

                score += 4

            # ------------------------------------------------
            # Serial
            # ------------------------------------------------

            if (
                "serial" in row_text
                or
                "s.no" in row_text
                or
                "sno" in row_text
            ):

                score += 2

            # ------------------------------------------------
            # Check if this is really the header
            # ------------------------------------------------

            has_student = (
                "student" in row_text
                or
                "student name" in row_text
            )

            has_program = (
                "program" in row_text
                or
                "programme" in row_text
            )

            has_attendance = (
                "attendance" in row_text
                or
                "percentage" in row_text
                or
                "present" in row_text
            )

            if (
                has_student
                and
                has_program
            ) or (
                has_program
                and
                has_attendance
            ):

                if score > best_score:

                    best_score = score

                    selected_rows = rows

                    selected_header_index = i

    # ========================================================
    # FALLBACK HEADER SEARCH
    # ========================================================

    if selected_rows is None:

        for rows in candidate_tables:

            for i, row in enumerate(rows):

                normalized = [
                    normalize_text(x)
                    for x in row
                ]

                has_student = any(
                    "student" in x
                    for x in normalized
                )

                has_program = any(
                    (
                        "program" in x
                        or
                        "programme" in x
                    )
                    for x in normalized
                )

                has_period = any(
                    (
                        "period" in x
                        or
                        "semester" in x
                        or
                        x == "sem"
                    )
                    for x in normalized
                )

                if (
                    has_student
                    and
                    has_program
                    and
                    has_period
                ):

                    selected_rows = rows

                    selected_header_index = i

                    break

            if selected_rows is not None:

                break

    # ========================================================
    # VALIDATE
    # ========================================================

    if selected_rows is None:

        raise ValueError(
            "Could not identify the student attendance table."
        )

    # ========================================================
    # HEADER
    # ========================================================

    headers = selected_rows[
        selected_header_index
    ]

    headers = [
        str(x).strip()
        for x in headers
    ]

    # Remove trailing blanks
    while (
        headers
        and
        headers[-1] == ""
    ):

        headers.pop()

    headers = make_unique_columns(
        headers
    )

    # ========================================================
    # DATA ROWS
    # ========================================================

    data_rows = selected_rows[
        selected_header_index + 1:
    ]

    cleaned_rows = []

    for row in data_rows:

        # Ignore blank rows
        if not any(
            str(x).strip()
            for x in row
        ):

            continue

        # Normalize length
        if len(row) < len(headers):

            row = row + [
                ""
                for _ in range(
                    len(headers) - len(row)
                )
            ]

        elif len(row) > len(headers):

            row = row[
                :len(headers)
            ]

        # Ignore report/footer rows
        row_text = " ".join(
            str(x).strip().lower()
            for x in row
        )

        if (
            "report name:" in row_text
            or
            "generated on:" in row_text
            or
            "printed on:" in row_text
        ):

            continue

        cleaned_rows.append(
            row
        )

    # ========================================================
    # DATAFRAME
    # ========================================================

    df = pd.DataFrame(
        cleaned_rows,
        columns=headers
    )

    # Remove empty rows
    df = df.dropna(
        how="all"
    )

    df = df[
        df.astype(str)
        .apply(
            lambda row: any(
                str(x).strip() != ""
                for x in row
            ),
            axis=1
        )
    ].copy()

    # Clean column names
    df.columns = make_unique_columns(
        df.columns
    )

    # Add School Name
    df["School Name"] = (
        school_name
    )

    return df


# ============================================================
# READ NORMAL EXCEL
# ============================================================

def read_normal_excel(
    uploaded_file
):

    filename = (
        uploaded_file.name
        .lower()
    )

    if filename.endswith(
        ".xlsx"
    ):

        df = pd.read_excel(
            uploaded_file,
            engine="openpyxl"
        )

    elif filename.endswith(
        ".xls"
    ):

        df = pd.read_excel(
            uploaded_file,
            engine="xlrd"
        )

    else:

        raise ValueError(
            "Unsupported Excel format."
        )

    df.columns = make_unique_columns(
        df.columns
    )

    return df


# ============================================================
# READ ATTENDANCE FILE
# ============================================================

def read_attendance_file(
    uploaded_file
):

    file_bytes = (
        uploaded_file.getvalue()
    )

    first_bytes = (
        file_bytes[:1000]
        .lower()
    )

    # Academia XML Spreadsheet
    if (
        b"<?xml" in first_bytes
        or
        b"<workbook" in first_bytes
    ):

        return read_academia_xml(
            file_bytes
        )

    # Normal Excel
    return read_normal_excel(
        uploaded_file
    )


# ============================================================
# FIND COLUMN
# ============================================================

def find_column(
    df,
    possible_names
):

    normalized_columns = {
        col: normalize_text(col)
        for col in df.columns
    }

    normalized_names = [
        normalize_text(name)
        for name in possible_names
    ]

    # ========================================================
    # EXACT MATCH
    # ========================================================

    for col, normalized_col in normalized_columns.items():

        if (
            normalized_col
            in
            normalized_names
        ):

            return col

    # ========================================================
    # PARTIAL MATCH
    # ========================================================

    for name in normalized_names:

        for col, normalized_col in normalized_columns.items():

            if (
                name in normalized_col
                or
                normalized_col in name
            ):

                return col

    return None


# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_files = st.file_uploader(
    "📂 Upload Academia ERP Attendance Summary files",
    type=[
        "xls",
        "xlsx"
    ],
    accept_multiple_files=True
)


# ============================================================
# PROCESS FILES
# ============================================================

if uploaded_files:

    all_data = []

    errors = []

    # ========================================================
    # PROCESS EACH FILE
    # ========================================================

    for uploaded_file in uploaded_files:

        try:

            df = read_attendance_file(
                uploaded_file
            )

            if (
                df is None
                or
                df.empty
            ):

                errors.append(
                    f"{uploaded_file.name}: No data found."
                )

                continue

            # Add source file
            df["Source File"] = (
                uploaded_file.name
            )

            all_data.append(
                df
            )

        except Exception as e:

            errors.append(
                f"{uploaded_file.name}: {str(e)}"
            )

    # ========================================================
    # ERRORS
    # ========================================================

    if errors:

        with st.expander(
            "⚠️ Files with errors"
        ):

            for error in errors:

                st.error(
                    error
                )

    # ========================================================
    # STOP IF NO DATA
    # ========================================================

    if not all_data:

        st.error(
            "No valid attendance files could be processed."
        )

        st.stop()

    # ========================================================
    # COMBINE FILES
    # ========================================================

    data = pd.concat(
        all_data,
        ignore_index=True
    )

    # ========================================================
    # FIND COLUMNS
    # ========================================================

    student_col = find_column(
        data,
        [
            "Student Name",
            "Student",
            "Student Name / ID",
            "Student ID / Name",
            "Name"
        ]
    )

    program_col = find_column(
        data,
        [
            "Program Name",
            "Program",
            "Programme Name",
            "Programme",
            "Program Code/Name",
            "Program Code / Name",
            "Programme Code/Name",
            "Programme Code / Name"
        ]
    )

    period_col = find_column(
        data,
        [
            "Period",
            "Period/Semester",
            "Period / Semester",
            "Semester",
            "Semester Name",
            "Academic Semester",
            "Sem"
        ]
    )

    attendance_col = find_column(
        data,
        [
            "Present Percentage",
            "Present %",
            "Present Percentage (%)",
            "Present %age",
            "Attendance Percentage",
            "Attendance %",
            "Attendance",
            "Attendance %age",
            "Percentage"
        ]
    )

    # ========================================================
    # VALIDATION
    # ========================================================

    missing = []

    if student_col is None:
        missing.append(
            "Student Name"
        )

    if program_col is None:
        missing.append(
            "Program Name"
        )

    if period_col is None:
        missing.append(
            "Period / Semester"
        )

    if attendance_col is None:
        missing.append(
            "Attendance %"
        )

    if missing:

        st.error(
            "The following required columns could not be identified:"
        )

        for item in missing:

            st.write(
                f"- **{item}**"
            )

        st.write(
            "### Available columns:"
        )

        for col in data.columns:

            st.write(
                f"- `{col}`"
            )

        st.stop()

    # ========================================================
    # STANDARDIZE
    # ========================================================

    data["Student Name"] = (
        data[student_col]
        .astype(str)
        .str.strip()
    )

    data["Program Name"] = (
        data[program_col]
        .astype(str)
        .str.strip()
    )

    data["Period"] = (
        data[period_col]
        .astype(str)
        .str.strip()
    )

    data["School Name"] = (
        data["School Name"]
        .astype(str)
        .str.strip()
    )

    # ========================================================
    # CLEAN ATTENDANCE
    # ========================================================

    data["Attendance"] = (
        data[attendance_col]
        .astype(str)
        .str.replace(
            "%",
            "",
            regex=False
        )
        .str.replace(
            ",",
            "",
            regex=False
        )
        .str.strip()
    )

    data["Attendance"] = pd.to_numeric(
        data["Attendance"],
        errors="coerce"
    )

    # Remove invalid
    data = data[
        data["Attendance"].notna()
    ].copy()

    data = data[
        (data["Attendance"] >= 0)
        &
        (data["Attendance"] <= 100)
    ].copy()

    # ========================================================
    # ATTENDANCE SLABS
    # ========================================================

    labels = [
        "Below 20%",
        "20-40%",
        "40-60%",
        "60-75%",
        "75% and Above"
    ]

    bins = [
        float("-inf"),
        20,
        40,
        60,
        75,
        float("inf")
    ]

    data["Attendance Slab"] = pd.cut(
        data["Attendance"],
        bins=bins,
        labels=labels,
        right=False
    )

    # ========================================================
    # FILTER TOGGLE
    # ========================================================

    show_filters = st.toggle(
        "🔎 Show Filters",
        value=False
    )

    # ========================================================
    # DEFAULT FILTER VALUES
    # ========================================================

    selected_schools = sorted(
        data["School Name"]
        .dropna()
        .unique(),
        key=natural_sort_value
    )

    selected_programs = sorted(
        data["Program Name"]
        .dropna()
        .unique(),
        key=natural_sort_value
    )

    selected_periods = sorted(
        data["Period"]
        .dropna()
        .unique(),
        key=natural_sort_value
    )

    # ========================================================
    # SIDEBAR FILTERS
    # ========================================================

    if show_filters:

        st.sidebar.header(
            "🔎 Filters"
        )

        # ----------------------------------------------------
        # School
        # ----------------------------------------------------

        schools = sorted(
            data["School Name"]
            .dropna()
            .unique(),
            key=natural_sort_value
        )

        selected_schools = st.sidebar.multiselect(
            "School",
            schools,
            default=schools
        )

        # ----------------------------------------------------
        # Program
        # ----------------------------------------------------

        available_programs = data[
            data["School Name"].isin(
                selected_schools
            )
        ]["Program Name"].dropna().unique()

        programs = sorted(
            available_programs,
            key=natural_sort_value
        )

        selected_programs = st.sidebar.multiselect(
            "Program",
            programs,
            default=programs
        )

        # ----------------------------------------------------
        # Period
        # ----------------------------------------------------

        available_periods = data[
            data["School Name"].isin(
                selected_schools
            )
            &
            data["Program Name"].isin(
                selected_programs
            )
        ]["Period"].dropna().unique()

        periods = sorted(
            available_periods,
            key=natural_sort_value
        )

        selected_periods = st.sidebar.multiselect(
            "Period / Semester",
            periods,
            default=periods
        )

    # ========================================================
    # FILTER DATA
    # ========================================================

    filtered_data = data[
        data["School Name"].isin(
            selected_schools
        )
        &
        data["Program Name"].isin(
            selected_programs
        )
        &
        data["Period"].isin(
            selected_periods
        )
    ].copy()

    # ========================================================
    # KPI
    # ========================================================

    st.subheader(
        "📌 Attendance Overview"
    )

    total_students = len(
        filtered_data
    )

    average_attendance = (
        filtered_data["Attendance"].mean()
        if total_students > 0
        else 0
    )

    below_75 = len(
        filtered_data[
            filtered_data["Attendance"] < 75
        ]
    )

    above_75 = len(
        filtered_data[
            filtered_data["Attendance"] >= 75
        ]
    )

    k1, k2, k3, k4 = st.columns(4)

    with k1:

        st.metric(
            "Total Students",
            f"{total_students:,}"
        )

    with k2:

        st.metric(
            "Average Attendance",
            f"{average_attendance:.2f}%"
        )

    with k3:

        st.metric(
            "Students Below 75%",
            f"{below_75:,}"
        )

    with k4:

        st.metric(
            "Students 75% and Above",
            f"{above_75:,}"
        )

    # ========================================================
    # CREATE SUMMARY
    # ========================================================

    summary = (
        filtered_data
        .groupby(
            [
                "School Name",
                "Program Name",
                "Period"
            ],
            dropna=False
        )
        .size()
        .reset_index(
            name="Total Students"
        )
    )

    # ========================================================
    # SLAB COUNTS
    # ========================================================

    slab_counts = (
        filtered_data
        .groupby(
            [
                "School Name",
                "Program Name",
                "Period",
                "Attendance Slab"
            ],
            observed=False
        )
        .size()
        .unstack(
            fill_value=0
        )
        .reset_index()
    )

    # Ensure all slabs
    for slab in labels:

        if slab not in slab_counts.columns:

            slab_counts[slab] = 0

    # ========================================================
    # MERGE
    # ========================================================

    summary_pivot = summary.merge(
        slab_counts,
        on=[
            "School Name",
            "Program Name",
            "Period"
        ],
        how="left"
    )

    # ========================================================
    # NUMERIC
    # ========================================================

    for col in labels:

        summary_pivot[col] = pd.to_numeric(
            summary_pivot[col],
            errors="coerce"
        ).fillna(0)

    summary_pivot[
        "Total Students"
    ] = pd.to_numeric(
        summary_pivot[
            "Total Students"
        ],
        errors="coerce"
    ).fillna(0)

    # ========================================================
    # SORT
    # ========================================================

    summary_pivot["_School_Sort"] = (
        summary_pivot[
            "School Name"
        ]
        .astype(str)
        .apply(
            natural_sort_value
        )
    )

    summary_pivot["_Program_Sort"] = (
        summary_pivot[
            "Program Name"
        ]
        .astype(str)
        .apply(
            natural_sort_value
        )
    )

    summary_pivot["_Period_Sort"] = (
        summary_pivot[
            "Period"
        ]
        .astype(str)
        .apply(
            natural_sort_value
        )
    )

    summary_pivot = (
        summary_pivot
        .sort_values(
            by=[
                "_School_Sort",
                "_Program_Sort",
                "_Period_Sort"
            ],
            kind="stable"
        )
        .drop(
            columns=[
                "_School_Sort",
                "_Program_Sort",
                "_Period_Sort"
            ]
        )
        .reset_index(
            drop=True
        )
    )

    # ========================================================
    # SCHOOL ORDER
    # ========================================================

    school_order = sorted(
        summary_pivot[
            "School Name"
        ].dropna().unique(),
        key=natural_sort_value
    )

    # ========================================================
    # BUILD FINAL SUMMARY
    # ========================================================

    final_rows = []

    for school in school_order:

        school_data = summary_pivot[
            summary_pivot[
                "School Name"
            ] == school
        ].copy()

        # ----------------------------------------------------
        # Sort program
        # ----------------------------------------------------

        school_data["_Program_Sort"] = (
            school_data[
                "Program Name"
            ]
            .astype(str)
            .apply(
                natural_sort_value
            )
        )

        # ----------------------------------------------------
        # Sort semester
        # ----------------------------------------------------

        school_data["_Period_Sort"] = (
            school_data[
                "Period"
            ]
            .astype(str)
            .apply(
                natural_sort_value
            )
        )

        school_data = (
            school_data
            .sort_values(
                by=[
                    "_Program_Sort",
                    "_Period_Sort"
                ],
                kind="stable"
            )
            .drop(
                columns=[
                    "_Program_Sort",
                    "_Period_Sort"
                ]
            )
        )

        # ----------------------------------------------------
        # NORMAL ROWS
        # ----------------------------------------------------

        for _, row in school_data.iterrows():

            row_dict = row.to_dict()

            # Percentage is only required
            # for School TOTAL / Grand Total
            for slab in labels:

                row_dict[
                    f"{slab} %"
                ] = None

            final_rows.append(
                row_dict
            )

        # ----------------------------------------------------
        # SCHOOL TOTAL
        # ----------------------------------------------------

        school_total_students = int(
            school_data[
                "Total Students"
            ].sum()
        )

        school_total = {

            "School Name":
                f"{school} TOTAL",

            "Program Name":
                "",

            "Period":
                "",

            "Total Students":
                school_total_students
        }

        # ----------------------------------------------------
        # Calculate totals
        # ----------------------------------------------------

        for slab in labels:

            count = int(
                school_data[
                    slab
                ].sum()
            )

            school_total[
                slab
            ] = count

            if school_total_students > 0:

                percentage = (
                    count
                    /
                    school_total_students
                    *
                    100
                )

            else:

                percentage = 0

            school_total[
                f"{slab} %"
            ] = percentage

        final_rows.append(
            school_total
        )

    # ========================================================
    # GRAND TOTAL
    # ========================================================

    grand_total_students = int(
        summary_pivot[
            "Total Students"
        ].sum()
    )

    grand_total = {

        "School Name":
            "GRAND TOTAL",

        "Program Name":
            "",

        "Period":
            "",

        "Total Students":
            grand_total_students
    }

    for slab in labels:

        count = int(
            summary_pivot[
                slab
            ].sum()
        )

        grand_total[
            slab
        ] = count

        if grand_total_students > 0:

            percentage = (
                count
                /
                grand_total_students
                *
                100
            )

        else:

            percentage = 0

        grand_total[
            f"{slab} %"
        ] = percentage

    final_rows.append(
        grand_total
    )

    # ========================================================
    # FINAL DATAFRAME
    # ========================================================

    final_summary = pd.DataFrame(
        final_rows
    )

    # ========================================================
    # SUMMARY TITLE
    # ========================================================

    st.subheader(
        "📋 School / Program / Semester Attendance Summary"
    )

    st.caption(
        "School TOTAL and GRAND TOTAL display "
        "attendance count followed by the school-level percentage in brackets."
    )


    # ========================================================
    # CUSTOM HTML TABLE
    # ========================================================

    table_html = """

    <style>

    /* -----------------------------------------------------
       MAIN WRAPPER
       ----------------------------------------------------- */

    .attendance-wrapper {

        width: 100%;

        overflow-x: auto;

        overflow-y: auto;

        border: 1px solid #d9d9d9;

        border-radius: 8px;

        background: white;

        max-height: 650px;

    }


    /* -----------------------------------------------------
       TABLE
       ----------------------------------------------------- */

    .attendance-table {

        border-collapse: collapse;

        width: 100%;

        min-width: 1250px;

        font-family:
            Arial,
            Helvetica,
            sans-serif;

        font-size: 14px;

    }


    /* -----------------------------------------------------
       HEADER
       ----------------------------------------------------- */

    .attendance-table th {

        background-color: #f5f5f5;

        color: #333333;

        font-weight: 600;

        text-align: center;

        padding: 12px 10px;

        border-bottom: 1px solid #d9d9d9;

        border-right: 1px solid #e5e5e5;

        white-space: nowrap;

        position: sticky;

        top: 0;

        z-index: 5;

    }


    .attendance-table th:nth-child(1) {

        text-align: left;

        min-width: 220px;

    }


    .attendance-table th:nth-child(2) {

        text-align: left;

        min-width: 420px;

    }


    .attendance-table th:nth-child(3) {

        text-align: left;

        min-width: 120px;

    }


    /* -----------------------------------------------------
       CELLS
       ----------------------------------------------------- */

    .attendance-table td {

        padding: 10px;

        border-bottom: 1px solid #e5e5e5;

        border-right: 1px solid #eeeeee;

        vertical-align: middle;

        white-space: nowrap;

    }


    .attendance-table td.school {

        text-align: left;

    }


    .attendance-table td.program {

        text-align: left;

    }


    .attendance-table td.period {

        text-align: left;

    }


    .attendance-table td.number {

        text-align: center;

    }


    /* -----------------------------------------------------
       TOTAL ROW
       ----------------------------------------------------- */

    .attendance-table tr.total-row {

        font-weight: bold;

        background-color: #f0f0f0;

    }


    /* -----------------------------------------------------
       GRAND TOTAL
       ----------------------------------------------------- */

    .attendance-table tr.grand-total-row {

        font-weight: bold;

        background-color: #e6e6e6;

        border-top: 2px solid #555555;

    }


    /* -----------------------------------------------------
       TOTAL VALUES
       ----------------------------------------------------- */

    .total-value {

        font-weight: bold;

        font-size: 14px;

    }


    /* -----------------------------------------------------
       FULLSCREEN BUTTON
       ----------------------------------------------------- */

    .fullscreen-button {

        margin-bottom: 10px;

        padding: 7px 15px;

        border: 1px solid #cccccc;

        border-radius: 6px;

        background-color: white;

        cursor: pointer;

        font-size: 14px;

    }


    .fullscreen-button:hover {

        background-color: #f2f2f2;

    }


    </style>


    <button
        class="fullscreen-button"
        onclick="openAttendanceFullScreen()">
        ⛶ Full Screen
    </button>


    <div
        class="attendance-wrapper"
        id="attendance-container">

    <table class="attendance-table">

    <thead>

    <tr>

        <th>School Name</th>

        <th>Program Name</th>

        <th>Period</th>

        <th>Below 20%</th>

        <th>20-40%</th>

        <th>40-60%</th>

        <th>60-75%</th>

        <th>75% and Above</th>

        <th>Total Students</th>

    </tr>

    </thead>

    <tbody>

    """

    # ========================================================
    # ADD TABLE ROWS
    # ========================================================

    def is_total_row(
        school_name
    ):

        school_name = str(
            school_name
        ).strip()

        return (
            school_name.endswith(
                " TOTAL"
            )
            or
            school_name == "GRAND TOTAL"
        )


    for _, row in final_summary.iterrows():

        school_name = str(
            row["School Name"]
        ).strip()

        program_name = str(
            row["Program Name"]
        ).strip()

        period = str(
            row["Period"]
        ).strip()

        total_row = (
            school_name.endswith(
                " TOTAL"
            )
        )

        grand_total_row = (
            school_name
            ==
            "GRAND TOTAL"
        )

        # ----------------------------------------------------
        # Row class
        # ----------------------------------------------------

        if grand_total_row:

            table_html += """
            <tr class="grand-total-row">
            """

        elif total_row:

            table_html += """
            <tr class="total-row">
            """

        else:

            table_html += """
            <tr>
            """

        # ----------------------------------------------------
        # Escape text safely
        # ----------------------------------------------------

        school_display = html.escape(
            school_name
        )

        program_display = html.escape(
            program_name
        )

        period_display = html.escape(
            period
        )

        # ----------------------------------------------------
        # School
        # ----------------------------------------------------

        table_html += f"""

        <td class="school">
            {school_display}
        </td>

        """

        # ----------------------------------------------------
        # Program
        # ----------------------------------------------------

        table_html += f"""

        <td class="program">
            {program_display}
        </td>

        """

        # ----------------------------------------------------
        # Period
        # ----------------------------------------------------

        table_html += f"""

        <td class="period">
            {period_display}
        </td>

        """

        # ----------------------------------------------------
        # Slabs
        # ----------------------------------------------------

        for slab in labels:

            count = int(
                pd.to_numeric(
                    row[slab],
                    errors="coerce"
                )
                if pd.notna(
                    row[slab]
                )
                else 0
            )

            if (
                total_row
                or
                grand_total_row
            ):

                percentage = float(
                    pd.to_numeric(
                        row[
                            f"{slab} %"
                        ],
                        errors="coerce"
                    )
                    if pd.notna(
                        row[
                            f"{slab} %"
                        ]
                    )
                    else 0
                )

                # --------------------------------------------
                # COUNT (PERCENTAGE)
                # --------------------------------------------

                table_html += f"""

                <td class="number">

                    <span class="total-value">

                        {count} ({percentage:.2f}%)

                    </span>

                </td>

                """

            else:

                table_html += f"""

                <td class="number">

                    {count}

                </td>

                """

        # ----------------------------------------------------
        # Total Students
        # ----------------------------------------------------

        total_students_value = int(
            pd.to_numeric(
                row[
                    "Total Students"
                ],
                errors="coerce"
            )
            if pd.notna(
                row[
                    "Total Students"
                ]
            )
            else 0
        )

        if (
            total_row
            or
            grand_total_row
        ):

            table_html += f"""

            <td class="number">

                <span class="total-value">

                    {total_students_value}

                </span>

            </td>

            """

        else:

            table_html += f"""

            <td class="number">

                {total_students_value}

            </td>

            """

        table_html += """
        </tr>
        """

    # ========================================================
    # CLOSE TABLE
    # ========================================================

    table_html += """

    </tbody>

    </table>

    </div>


    <script>

    function openAttendanceFullScreen() {

        const container =
            document.getElementById(
                "attendance-container"
            );

        if (
            container.requestFullscreen
        ) {

            container.requestFullscreen();

        }

        else if (
            container.webkitRequestFullscreen
        ) {

            container.webkitRequestFullscreen();

        }

        else if (
            container.msRequestFullscreen
        ) {

            container.msRequestFullscreen();

        }

    }

    </script>

    """


    # ========================================================
    # DISPLAY TABLE
    # ========================================================

    components.html(
        table_html,
        height=700,
        scrolling=True
    )


    # ========================================================
    # DOWNLOAD SUMMARY CSV
    # ========================================================

    csv_summary = final_summary.copy()

    csv_columns = [

        "School Name",

        "Program Name",

        "Period",

        "Below 20%",

        "Below 20% %",

        "20-40%",

        "20-40% %",

        "40-60%",

        "40-60% %",

        "60-75%",

        "60-75% %",

        "75% and Above",

        "75% and Above %",

        "Total Students"
    ]

    csv_columns = [
        col
        for col in csv_columns
        if col in csv_summary.columns
    ]

    csv_summary = csv_summary[
        csv_columns
    ]

    summary_csv = (
        csv_summary
        .to_csv(
            index=False
        )
        .encode(
            "utf-8-sig"
        )
    )

    st.download_button(
        label="⬇️ Download Attendance Summary CSV",
        data=summary_csv,
        file_name="attendance_summary.csv",
        mime="text/csv"
    )


    # ========================================================
    # OVERALL ATTENDANCE DISTRIBUTION
    # ========================================================

    st.subheader(
        "📊 Overall Attendance Slab Distribution"
    )

    overall_counts = (
        filtered_data[
            "Attendance Slab"
        ]
        .value_counts(
            sort=False
        )
        .reindex(
            labels,
            fill_value=0
        )
        .reset_index()
    )

    overall_counts.columns = [
        "Attendance Slab",
        "Students"
    ]

    fig_overall = px.bar(
        overall_counts,
        x="Attendance Slab",
        y="Students",
        text="Students",
        title="Students by Attendance Slab"
    )

    fig_overall.update_traces(
        textposition="outside"
    )

    fig_overall.update_layout(
        xaxis_title="Attendance Slab",
        yaxis_title="Number of Students"
    )

    st.plotly_chart(
        fig_overall,
        use_container_width=True
    )


    # ========================================================
    # SCHOOL-WISE AVERAGE
    # ========================================================

    st.subheader(
        "🏫 School-wise Average Attendance"
    )

    school_avg = (
        filtered_data
        .groupby(
            "School Name",
            as_index=False
        )["Attendance"]
        .mean()
    )

    school_avg = school_avg.sort_values(
        "Attendance",
        ascending=False
    )

    fig_school = px.bar(
        school_avg,
        x="School Name",
        y="Attendance",
        text=school_avg[
            "Attendance"
        ].round(2),
        title="Average Attendance by School"
    )

    fig_school.update_traces(
        textposition="outside"
    )

    fig_school.update_layout(
        xaxis_title="School",
        yaxis_title="Average Attendance (%)"
    )

    st.plotly_chart(
        fig_school,
        use_container_width=True
    )


    # ========================================================
    # PROGRAM-WISE AVERAGE
    # ========================================================

    st.subheader(
        "🎓 Program-wise Average Attendance"
    )

    program_avg = (
        filtered_data
        .groupby(
            "Program Name",
            as_index=False
        )["Attendance"]
        .mean()
    )

    program_avg = program_avg.sort_values(
        "Attendance",
        ascending=False
    )

    fig_program = px.bar(
        program_avg,
        x="Program Name",
        y="Attendance",
        text=program_avg[
            "Attendance"
        ].round(2),
        title="Average Attendance by Program"
    )

    fig_program.update_traces(
        textposition="outside"
    )

    fig_program.update_layout(
        xaxis_title="Program",
        yaxis_title="Average Attendance (%)"
    )

    st.plotly_chart(
        fig_program,
        use_container_width=True
    )


    # ========================================================
    # STUDENT-WISE DETAILS
    # ========================================================

    st.subheader(
        "👨‍🎓 Student-wise Attendance Details"
    )

    student_display = filtered_data[
        [
            "School Name",
            "Program Name",
            "Period",
            "Student Name",
            "Attendance",
            "Attendance Slab"
        ]
    ].copy()

    # --------------------------------------------------------
    # Sort student data
    # --------------------------------------------------------

    student_display = student_display.sort_values(
        by=[
            "School Name",
            "Program Name",
            "Period",
            "Student Name"
        ],
        key=lambda col: col.map(
            natural_sort_value
        )
    )

    st.dataframe(
        student_display,
        use_container_width=True,
        hide_index=True
    )


    # ========================================================
    # DOWNLOAD STUDENT DATA
    # ========================================================

    student_csv = filtered_data[
        [
            "School Name",
            "Program Name",
            "Period",
            "Student Name",
            "Attendance",
            "Attendance Slab",
            "Source File"
        ]
    ].to_csv(
        index=False
    ).encode(
        "utf-8-sig"
    )

    st.download_button(
        label="⬇️ Download Student-wise Attendance CSV",
        data=student_csv,
        file_name="student_attendance_details.csv",
        mime="text/csv"
    )


    # ========================================================
    # RAW PROCESSED DATA
    # ========================================================

    with st.expander(
        "🔍 View Processed Data"
    ):

        st.dataframe(
            filtered_data,
            use_container_width=True,
            hide_index=True
        )

        st.write(
            f"Total records: **{len(filtered_data):,}**"
        )


# ============================================================
# NO FILE UPLOADED
# ============================================================

else:

    st.info(
        "👆 Please upload one or more Academia ERP "
        "Student Attendance Summary files."
    )

    st.markdown(
        """
        ### Attendance Slabs

        | Attendance | Category |
        |---|---|
        | Below 20% | Critical |
        | 20% – 40% | Very Low |
        | 40% – 60% | Low |
        | 60% – 75% | Moderate |
        | 75% and Above | Satisfactory |

        **Multiple Academia ERP files can be uploaded together.**
        """
    )