import pandas as pd
import pdfplumber
import os
import re


# ============================================================
# NORMALIZE IDENTIFIER
# ============================================================

def normalize_hawb(value):

    if value is None:
        return ""

    # Avoid pd.isna() problems with lists/Series/DataFrames
    if isinstance(value, (list, tuple, pd.Series, pd.DataFrame)):
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    value = str(value)

    # Normalize common Excel / PDF characters
    value = (
        value
        .replace("\u00a0", " ")
        .replace("\u200b", "")
        .replace("\ufeff", "")
        .strip()
        .upper()
    )

    # Remove spaces and separators
    value = re.sub(r"[\s\-_/]+", "", value)

    return value


# ============================================================
# NORMALIZE HEADER
# ============================================================

def normalize_header(value):

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    value = str(value)

    value = (
        value
        .replace("\u00a0", " ")
        .replace("\n", " ")
        .replace("\r", " ")
        .strip()
        .upper()
    )

    # Normalize multiple spaces
    value = re.sub(r"\s+", " ", value)

    return value


# ============================================================
# IDENTIFY AWB / HAWB COLUMN
# ============================================================

def is_awb_header(value):

    header = normalize_header(value)

    if not header:
        return False

    # Remove spaces for easier comparison
    compact = header.replace(" ", "")

    possible_headers = {

        "AWB",
        "HAWB",
        "AWB/BL",
        "AWB/BLNO",
        "AWB/BLNº",
        "AWB/BLNO.",
        "AWB/BLN",
        "AWB/HBL",
        "AWB/HBLNO",
        "HAWBNO",
        "HAWBNº",
        "AWBHBLNO",

    }

    if compact in possible_headers:
        return True

    # Flexible checks
    if compact.startswith("AWB") and len(compact) <= 15:
        return True

    if compact.startswith("HAWB") and len(compact) <= 15:
        return True

    return False


# ============================================================
# IDENTIFY CW COLUMN
# ============================================================

def is_cw_header(value):

    header = normalize_header(value)

    if not header:
        return False

    compact = header.replace(" ", "")

    return compact in {
        "CW",
        "CWT",
        "CHARGEABLEWEIGHT",
        "CHARGEABLEWT",
        "CHARGEABLEW"
    }


# ============================================================
# FIND EXCEL HEADER ROW
# ============================================================

def find_excel_header(excel_file):

    # Read a reasonable number of rows.
    # We do NOT depend on row 2 specifically.
    df = pd.read_excel(
        excel_file,
        header=None,
        nrows=50
    )

    for row_number in range(len(df)):

        row_values = df.iloc[row_number].tolist()

        awb_found = False
        cw_found = False

        for value in row_values:

            if is_awb_header(value):
                awb_found = True

            if is_cw_header(value):
                cw_found = True

        # Main provider formats:
        # AWB + CW
        # HAWB + CW
        # AWB/BL Nº + CW
        if awb_found and cw_found:
            return row_number

    # --------------------------------------------------------
    # SECONDARY DETECTION
    # --------------------------------------------------------
    #
    # Some Excel files may contain strange formatting,
    # hidden columns, filters, merged cells, etc.
    #
    # If we find AWB but not CW in the same row, we inspect
    # whether another column contains a likely CW header.
    #

    for row_number in range(len(df)):

        row_values = df.iloc[row_number].tolist()

        awb_found = any(
            is_awb_header(value)
            for value in row_values
        )

        if not awb_found:
            continue

        # Search for CW-like headers
        for value in row_values:

            header = normalize_header(value)

            if (
                header == "CW"
                or "CHARGEABLE WEIGHT" in header
                or "CHARGEABLEWEIGHT" in header
            ):

                return row_number

    raise Exception(
        "Unable to locate Excel header row. "
        "Expected an AWB/HAWB/AWB-BL header and a CW header."
    )


# ============================================================
# FIND ACTUAL COLUMNS
# ============================================================

def find_column(columns, detector):

    for column in columns:

        if detector(column):
            return column

    return None


# ============================================================
# EXTRACT IDENTIFIER FROM PDF FILENAME
# ============================================================

def extract_hawb_from_filename(filename):

    if not filename:
        return None

    name = os.path.basename(filename)

    # Remove extension
    name = re.sub(
        r"\.pdf$",
        "",
        name,
        flags=re.IGNORECASE
    )

    upper_name = name.upper()

    # Normalize special spaces
    upper_name = (
        upper_name
        .replace("\u00a0", " ")
        .replace("\u200b", "")
        .replace("\ufeff", "")
    )

    # ========================================================
    # PATTERN 1
    # Explicit HAWB No / HAWB Nº
    #
    # Example:
    # Copy 5 - (Extra Copy) - HAWB No_ J560885
    # ========================================================

    match = re.search(
        r"HAWB\s*(?:NO|Nº|N°|NUMBER)?\s*[_:\-]?\s*"
        r"([A-Z0-9][A-Z0-9\-_]{4,})",
        upper_name
    )

    if match:

        identifier = normalize_hawb(
            match.group(1)
        )

        if identifier:
            return identifier

    # ========================================================
    # PATTERN 2
    # AWB No / AWB Nº
    # ========================================================

    match = re.search(
        r"AWB\s*(?:NO|Nº|N°|NUMBER)?\s*[_:\-]?\s*"
        r"([A-Z0-9][A-Z0-9\-_]{4,})",
        upper_name
    )

    if match:

        identifier = normalize_hawb(
            match.group(1)
        )

        if identifier:
            return identifier

    # ========================================================
    # PATTERN 3
    # 3 digits + hyphen + 8 digits
    #
    # Example:
    # 810-42903125
    # 202-31647291
    # 406-06772010
    # ========================================================

    match = re.search(
        r"\b\d{3}-?\d{8}\b",
        upper_name
    )

    if match:

        return normalize_hawb(
            match.group(0)
        )

    # ========================================================
    # PATTERN 4
    # Letter prefix + digits
    #
    # Examples:
    # J560885
    # J603482
    # PTY0045653
    #
    # Minimum 5 characters.
    # ========================================================

    matches = re.findall(
        r"\b[A-Z]{1,6}\d{5,12}\b",
        upper_name
    )

    if matches:

        # Prefer the last match because filenames often
        # contain descriptive text before the actual HAWB.
        return normalize_hawb(
            matches[-1]
        )

    # ========================================================
    # PATTERN 5
    # Generic numeric AWB
    # ========================================================

    matches = re.findall(
        r"\b\d{8,12}\b",
        upper_name
    )

    if matches:

        return normalize_hawb(
            matches[-1]
        )

    return None


# ============================================================
# EXTRACT CW FROM PDF
# ============================================================

def extract_pdf_cw(pdf_path):

    pdf_text = ""

    try:

        with pdfplumber.open(pdf_path) as pdf:

            for page in pdf.pages:

                text = page.extract_text()

                if text:

                    pdf_text += text + "\n"

    except Exception:

        return None

    if not pdf_text:
        return None

    lines = pdf_text.split("\n")

    # ========================================================
    # ORIGINAL PROVIDER
    #
    # Example pattern:
    #
    # 123 456.78 K 535.5 123.45
    #
    # ========================================================

    for line in lines:

        match = re.search(
            r"\d+\s+\d+(?:\.\d+)?K\s+"
            r"[A-Z]\s+"
            r"(\d+(?:\.\d+)?)\s+"
            r"\d+(?:\.\d+)?",
            line,
            flags=re.IGNORECASE
        )

        if match:

            try:

                return float(
                    match.group(1)
                )

            except Exception:
                pass

    # ========================================================
    # SEARCH FOR CHARGEABLE WEIGHT
    # ========================================================

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

                except Exception:
                    pass

    return None


# ============================================================
# READ EXCEL
# ============================================================

def read_excel_file(excel_file):

    header_row = find_excel_header(
        excel_file
    )

    df = pd.read_excel(
        excel_file,
        header=header_row
    )

    return df, header_row


# ============================================================
# VALIDATE EXCEL STRUCTURE
# ============================================================

def prepare_excel(df):

    # --------------------------------------------------------
    # Find AWB / HAWB column
    # --------------------------------------------------------

    awb_column = find_column(
        df.columns,
        is_awb_header
    )

    if awb_column is None:

        raise Exception(
            "AWB/HAWB column was not found in the Excel file."
        )

    # --------------------------------------------------------
    # Find CW column
    # --------------------------------------------------------

    cw_column = find_column(
        df.columns,
        is_cw_header
    )

    if cw_column is None:

        raise Exception(
            "CW column was not found in the Excel file."
        )

    # --------------------------------------------------------
    # Rename internally
    # --------------------------------------------------------

    df = df.rename(
        columns={
            awb_column: "HAWB",
            cw_column: "CW"
        }
    )

    return df


# ============================================================
# CREATE PDF INDEX
# ============================================================

def create_pdf_index(input_folder):

    pdf_index = {}

    for file in os.listdir(input_folder):

        if not file.lower().endswith(".pdf"):
            continue

        hawb = extract_hawb_from_filename(
            file
        )

        if hawb:

            pdf_index[hawb] = file

    return pdf_index


# ============================================================
# VALIDATE AWB
# ============================================================

def validate_awb(
    input_folder,
    output_folder
):

    # ========================================================
    # FIND EXCEL
    # ========================================================

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

    # ========================================================
    # READ EXCEL
    # ========================================================

    df_excel, header_row = read_excel_file(
        excel_file
    )

    print(
        f"Header detected on row: {header_row}"
    )

    df_excel = prepare_excel(
        df_excel
    )

    # ========================================================
    # NORMALIZE EXCEL DATA
    # ========================================================

    df_excel["HAWB"] = (
        df_excel["HAWB"]
        .apply(normalize_hawb)
    )

    df_excel["CW"] = pd.to_numeric(
        df_excel["CW"],
        errors="coerce"
    )

    # Remove completely empty HAWB rows

    df_excel = df_excel[
        df_excel["HAWB"] != ""
    ].copy()

    # ========================================================
    # CREATE PDF INDEX
    # ========================================================

    pdf_index = create_pdf_index(
        input_folder
    )

    print(
        f"PDFs indexed: {len(pdf_index)}"
    )

    # ========================================================
    # VALIDATION
    # ========================================================

    results = []

    for _, row in df_excel.iterrows():

        excel_hawb = row["HAWB"]
        excel_cw = row["CW"]

        # ----------------------------------------------------
        # HAWB empty
        # ----------------------------------------------------

        if not excel_hawb:

            continue

        # ----------------------------------------------------
        # PDF NOT FOUND
        # ----------------------------------------------------

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

        # ----------------------------------------------------
        # PDF FOUND
        # ----------------------------------------------------

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

        # ----------------------------------------------------
        # CW NOT FOUND
        # ----------------------------------------------------

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

        # ----------------------------------------------------
        # COMPARE CW
        # ----------------------------------------------------

        if pd.isna(excel_cw):

            results.append({

                "HAWB": excel_hawb,
                "Excel CW": "",
                "PDF CW": pdf_cw,
                "Difference": "",
                "Result": "CW NOT FOUND IN EXCEL",
                "PDF File": pdf_file

            })

            continue

        difference = round(
            abs(
                float(excel_cw) - float(pdf_cw)
            ),
            2
        )

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

    # ========================================================
    # FIND EXTRA PDFs
    # ========================================================

    excel_hawb_set = set(
        df_excel["HAWB"]
    )

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

    # ========================================================
    # SAVE RESULT
    # ========================================================

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
        "Validation completed:",
        output_file
    )

    return output_file