import pandas as pd
import pdfplumber
import os
import re


# =====================================================
# NORMALIZE GENERAL TEXT
# =====================================================

def normalize_text(value):

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    value = str(value)

    # Non-breaking spaces
    value = value.replace("\xa0", " ")

    # Line breaks / tabs
    value = value.replace("\n", " ")
    value = value.replace("\r", " ")
    value = value.replace("\t", " ")

    # Normalize multiple spaces
    value = re.sub(r"\s+", " ", value)

    return value.strip().upper()


# =====================================================
# NORMALIZE HAWB / AWB
# =====================================================

def normalize_hawb(value):

    # Important:
    # Only accept scalar values.
    # This prevents the previous error where
    # pd.isna() received a Series/DataFrame.

    if value is None:
        return ""

    if isinstance(value, (pd.Series, pd.DataFrame)):
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        return ""

    return (
        str(value)
        .strip()
        .upper()
        .replace(" ", "")
        .replace("-", "")
    )


# =====================================================
# NORMALIZE HEADER
# =====================================================

def normalize_header(value):

    value = normalize_text(value)

    # Remove characters that should not matter
    value = (
        value
        .replace("º", "")
        .replace("°", "")
        .replace(".", "")
    )

    return value


# =====================================================
# IDENTIFY AWB / HAWB HEADER
# =====================================================

def is_awb_header(value):

    value = normalize_header(value)

    if not value:
        return False

    compact = value.replace(" ", "")

    # Exact common formats
    valid_headers = {

        "AWB",
        "HAWB",

        "AWBBL",
        "AWBBLNO",
        "AWBBLN",

        "HAWBBL",
        "HAWBBLNO",
        "HAWBBLN",

    }

    if compact in valid_headers:
        return True

    # Flexible detection
    if compact.startswith("AWB"):
        return True

    if compact.startswith("HAWB"):
        return True

    return False


# =====================================================
# IDENTIFY CW HEADER
# =====================================================

def is_cw_header(value):

    value = normalize_header(value)

    if not value:
        return False

    compact = value.replace(" ", "")

    return compact == "CW"


# =====================================================
# FIND EXCEL HEADER ROW
# =====================================================

def find_excel_header(excel_file):

    df = pd.read_excel(
        excel_file,
        header=None,
        nrows=50
    )

    for i in range(len(df)):

        row = df.iloc[i].tolist()

        awb_found = False
        cw_found = False

        for value in row:

            if is_awb_header(value):
                awb_found = True

            if is_cw_header(value):
                cw_found = True

        if awb_found and cw_found:

            print(
                f"Excel header detected on row {i + 1}"
            )

            return i

    # =================================================
    # DEBUG
    # =================================================

    debug_rows = []

    for i in range(len(df)):

        row = df.iloc[i].tolist()

        normalized_row = [
            normalize_header(x)
            for x in row
        ]

        debug_rows.append(
            f"ROW {i + 1}: {normalized_row}"
        )

    raise Exception(
        "Unable to locate Excel header row.\n\n"
        "The program searched the first 50 rows "
        "for an AWB/HAWB/AWB-BL column and a CW column.\n\n"
        + "\n".join(debug_rows)
    )


# =====================================================
# FIND IDENTIFIER COLUMN
# =====================================================

def find_awb_column(columns):

    # First pass: exact / normalized matches

    for column in columns:

        if is_awb_header(column):

            return column

    # Second pass: flexible matching

    for column in columns:

        normalized = normalize_header(column)
        compact = normalized.replace(" ", "")

        if compact.startswith("AWB"):
            return column

        if compact.startswith("HAWB"):
            return column

    return None


# =====================================================
# FIND CW COLUMN
# =====================================================

def find_cw_column(columns):

    for column in columns:

        if is_cw_header(column):

            return column

    return None


# =====================================================
# EXTRACT HAWB / AWB FROM PDF FILE NAME
# =====================================================

def extract_hawb_from_filename(filename):

    # Remove [123], [1], etc.
    filename = re.sub(
        r"\[\d+\]",
        "",
        filename
    )

    filename_upper = filename.upper()

    # -------------------------------------------------
    # Pattern 1
    # Alphanumeric HAWB
    #
    # Example:
    # PTY0045653
    # I879513
    # J158916
    # -------------------------------------------------

    match = re.search(
        r"([A-Z]{1,5}\d{5,})",
        filename_upper
    )

    if match:

        return normalize_hawb(
            match.group(1)
        )

    # -------------------------------------------------
    # Pattern 2
    # Numeric AWB
    #
    # Example:
    # 992-10764250
    # 202-31652291
    # -------------------------------------------------

    match = re.search(
        r"(\d{3}-?\d{5,})",
        filename_upper
    )

    if match:

        return normalize_hawb(
            match.group(1)
        )

    return None


# =====================================================
# EXTRACT CW FROM PDF
# =====================================================

def extract_pdf_cw(pdf_path):

    pdf_text = ""

    try:

        with pdfplumber.open(pdf_path) as pdf:

            for page in pdf.pages:

                text = page.extract_text()

                if text:

                    pdf_text += text + "\n"

    except Exception as e:

        print(
            f"Unable to read PDF: {pdf_path}"
        )

        print(e)

        return None

    lines = pdf_text.split("\n")

    # =================================================
    # ORIGINAL PROVIDER
    # =================================================

    for line in lines:

        match = re.search(
            r"\d+\s+\d+(?:\.\d+)?K\s+[A-Z]\s+"
            r"(\d+(?:\.\d+)?)\s+"
            r"\d+(?:\.\d+)?",
            line
        )

        if match:

            try:

                return float(
                    match.group(1)
                )

            except (ValueError, TypeError):

                pass

    # =================================================
    # SVC PROVIDER
    # =================================================

    for line in lines:

        upper_line = line.upper()

        if (
            "CHARGEABLE WEIGHT" in upper_line
            or "CHARGEABLE" in upper_line
            or "CWT" in upper_line
        ):

            numbers = re.findall(
                r"\d+(?:\.\d+)?",
                line
            )

            if numbers:

                try:

                    return float(
                        numbers[-1]
                    )

                except (ValueError, TypeError):

                    pass

    return None


# =====================================================
# VALIDATE AWB
# =====================================================

def validate_awb(
    input_folder,
    output_folder
):

    # =================================================
    # FIND EXCEL
    # =================================================

    excel_file = None

    for file in os.listdir(input_folder):

        if file.lower().endswith(".xlsx"):

            excel_file = os.path.join(
                input_folder,
                file
            )

            break

    if not excel_file:

        raise Exception(
            "No Excel file found."
        )

    print(
        f"Excel found: {excel_file}"
    )

    # =================================================
    # FIND HEADER
    # =================================================

    header_row = find_excel_header(
        excel_file
    )

    print(
        f"Header detected on Excel row: "
        f"{header_row + 1}"
    )

    # =================================================
    # READ EXCEL
    # =================================================

    df_excel = pd.read_excel(
        excel_file,
        header=header_row
    )

    # =================================================
    # NORMALIZE COLUMN NAMES
    # =================================================

    original_columns = list(
        df_excel.columns
    )

    print(
        "Excel columns detected:"
    )

    print(
        original_columns
    )

    # =================================================
    # FIND AWB / HAWB COLUMN
    # =================================================

    awb_column = find_awb_column(
        df_excel.columns
    )

    if awb_column is None:

        raise Exception(
            "AWB/HAWB column could not be identified."
        )

    print(
        f"AWB/HAWB column detected: "
        f"{awb_column}"
    )

    # =================================================
    # FIND CW COLUMN
    # =================================================

    cw_column = find_cw_column(
        df_excel.columns
    )

    if cw_column is None:

        raise Exception(
            "CW column could not be identified."
        )

    print(
        f"CW column detected: "
        f"{cw_column}"
    )

    # =================================================
    # CREATE INTERNAL COLUMNS
    # =================================================

    # IMPORTANT:
    # Use the selected column directly.
    # Do not assume it is column A, D, E, etc.

    df_excel["HAWB_INTERNAL"] = (
        df_excel[awb_column]
        .apply(normalize_hawb)
    )

    df_excel["CW_INTERNAL"] = pd.to_numeric(
        df_excel[cw_column],
        errors="coerce"
    )

    # =================================================
    # CREATE PDF INDEX
    # =================================================

    pdf_index = {}

    for file in os.listdir(input_folder):

        if file.lower().endswith(".pdf"):

            hawb = extract_hawb_from_filename(
                file
            )

            if hawb:

                pdf_index[hawb] = file

    print(
        f"PDFs found and indexed: "
        f"{len(pdf_index)}"
    )

    # =================================================
    # VALIDATION
    # =================================================

    results = []

    for _, row in df_excel.iterrows():

        excel_hawb = row["HAWB_INTERNAL"]

        excel_cw = row["CW_INTERNAL"]

        # -------------------------------------------------
        # EMPTY AWB
        # -------------------------------------------------

        if not excel_hawb:

            results.append({

                "HAWB": "",
                "Excel CW": excel_cw,
                "PDF CW": "",
                "Difference": "",
                "Result": "AWB/HAWB EMPTY",
                "PDF File": ""

            })

            continue

        # -------------------------------------------------
        # PDF NOT FOUND
        # -------------------------------------------------

        if excel_hawb not in pdf_index:

            results.append({

                "HAWB": excel_hawb,
                "Excel CW": excel_cw,
                "PDF CW": "",
                "Difference": "",
                "Result": "PDF NOT FOUND",
                "PDF File": ""

            })

            continue

        # -------------------------------------------------
        # PDF FOUND
        # -------------------------------------------------

        pdf_file = pdf_index[
            excel_hawb
        ]

        pdf_path = os.path.join(
            input_folder,
            pdf_file
        )

        pdf_cw = extract_pdf_cw(
            pdf_path
        )

        # -------------------------------------------------
        # CW NOT FOUND
        # -------------------------------------------------

        if pdf_cw is None:

            results.append({

                "HAWB": excel_hawb,
                "Excel CW": excel_cw,
                "PDF CW": "",
                "Difference": "",
                "Result": "CW NOT FOUND IN PDF",
                "PDF File": pdf_file

            })

            continue

        # -------------------------------------------------
        # COMPARE CW
        # -------------------------------------------------

        if pd.isna(excel_cw):

            results.append({

                "HAWB": excel_hawb,
                "Excel CW": "",
                "PDF CW": pdf_cw,
                "Difference": "",
                "Result": "CW EMPTY IN EXCEL",
                "PDF File": pdf_file

            })

            continue

        difference = round(
            abs(
                float(excel_cw) - float(pdf_cw)
            ),
            2
        )

        # -------------------------------------------------
        # RESULT
        # -------------------------------------------------

        if difference <= 0.01:

            result = "PASS"

        else:

            result = "FAIL"

        results.append({

            "HAWB": excel_hawb,
            "Excel CW": excel_cw,
            "PDF CW": pdf_cw,
            "Difference": difference,
            "Result": result,
            "PDF File": pdf_file

        })

    # =================================================
    # EXTRA PDFs
    # =================================================

    excel_hawb_set = set(
        df_excel["HAWB_INTERNAL"]
    )

    # Remove blanks
    excel_hawb_set.discard("")

    for pdf_hawb, pdf_file in pdf_index.items():

        if pdf_hawb not in excel_hawb_set:

            results.append({

                "HAWB": pdf_hawb,
                "Excel CW": "",
                "PDF CW": "",
                "Difference": "",
                "Result": "HAWB NOT FOUND IN EXCEL",
                "PDF File": pdf_file

            })

    # =================================================
    # SAVE RESULT
    # =================================================

    os.makedirs(
        output_folder,
        exist_ok=True
    )

    output_file = os.path.join(
        output_folder,
        "AWB_Validation_Result.xlsx"
    )

    df_results = pd.DataFrame(
        results
    )

    df_results.to_excel(
        output_file,
        index=False
    )

    print(
        "Validation completed:"
    )

    print(
        output_file
    )

    return output_file