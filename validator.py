import pandas as pd
import pdfplumber
import os
import re


# ============================================================
# NORMALIZE TEXT
# ============================================================

def normalize_text(value):

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    text = str(value)

    # Replace non-breaking spaces and similar characters
    text = (
        text
        .replace("\xa0", " ")
        .replace("\u200b", "")
        .replace("\u200c", "")
        .replace("\u200d", "")
        .replace("\ufeff", "")
    )

    # Normalize spaces
    text = re.sub(r"\s+", " ", text)

    return text.strip()


# ============================================================
# NORMALIZE COLUMN NAME
# ============================================================

def normalize_column_name(value):

    text = normalize_text(value).upper()

    # Remove accents / special variants
    replacements = {
        "Á": "A",
        "É": "E",
        "Í": "I",
        "Ó": "O",
        "Ú": "U",
        "Ñ": "N",
        "º": "",
        "°": "",
        "ª": "",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    # Remove spaces around slash
    text = re.sub(r"\s*/\s*", "/", text)

    # Normalize punctuation
    text = text.replace("_", " ")
    text = re.sub(r"\s+", " ", text)

    return text.strip()


# ============================================================
# NORMALIZE HAWB / AWB
# ============================================================

def normalize_hawb(value):

    if value is None:
        return ""

    # IMPORTANT:
    # Some Excel cells can contain lists/arrays/objects.
    # pd.isna() on those can return an array and cause:
    #
    # ValueError: The truth value of an array is ambiguous
    #
    # Therefore only use pd.isna() for scalar values.

    if isinstance(value, (list, tuple, set, dict)):
        return ""

    try:
        missing = pd.isna(value)

        if isinstance(missing, bool) and missing:
            return ""

    except (TypeError, ValueError):
        pass

    text = str(value)

    text = (
        text
        .strip()
        .upper()
        .replace("\xa0", "")
        .replace(" ", "")
        .replace("-", "")
        .replace("_", "")
    )

    return text


# ============================================================
# FIND EXCEL HEADER ROW
# ============================================================

def find_excel_header(excel_file):

    df = pd.read_excel(
        excel_file,
        header=None,
        nrows=20
    )

    # Possible AWB headers
    awb_headers = {
        "AWB",
        "HAWB",
        "AWB/BL",
        "AWB/ BL",
        "AWB/BL NO",
        "AWB/ BL NO",
        "AWB/HBL",
        "AWB/ HBL",
        "AWB/HBL NO",
        "AWB/ HBL NO",
        "AWB/BL NO.",
        "AWB/ BL NO.",
        "AWB/HBL NO.",
        "AWB/ HBL NO.",
        "AWB NO",
        "HAWB NO",
    }

    # Normalize the possible headers
    normalized_awb_headers = {
        normalize_column_name(x)
        for x in awb_headers
    }

    for row_index in range(len(df)):

        row_values = []

        for value in df.iloc[row_index].tolist():

            normalized = normalize_column_name(value)

            if normalized:
                row_values.append(normalized)

        # Need at least one AWB-like column
        awb_found = any(
            value in normalized_awb_headers
            for value in row_values
        )

        # CW can sometimes contain spaces or different capitalization
        cw_found = any(
            value == "CW"
            or value.startswith("CW ")
            for value in row_values
        )

        if awb_found and cw_found:
            return row_index

    raise Exception(
        "Unable to locate Excel header row.\n\n"
        "The file must contain an AWB/HAWB column "
        "and a CW column within the first 20 rows."
    )


# ============================================================
# FIND AWB COLUMN
# ============================================================

def find_awb_column(columns):

    candidates = []

    for column in columns:

        normalized = normalize_column_name(column)

        # Exact matches
        if normalized in {
            "AWB",
            "HAWB",
            "AWB/BL",
            "AWB/ BL",
            "AWB/BL NO",
            "AWB/ BL NO",
            "AWB/HBL",
            "AWB/ HBL",
            "AWB/HBL NO",
            "AWB/ HBL NO",
            "AWB NO",
            "HAWB NO",
        }:
            return column

        # More flexible matching
        compact = normalized.replace(" ", "")

        if compact in {
            "AWB",
            "HAWB",
            "AWB/BL",
            "AWB/BLNO",
            "AWB/HBL",
            "AWB/HBLNO",
            "AWBNO",
            "HAWBNO",
        }:
            candidates.append(column)

    if candidates:
        return candidates[0]

    return None


# ============================================================
# FIND CW COLUMN
# ============================================================

def find_cw_column(columns):

    candidates = []

    for column in columns:

        normalized = normalize_column_name(column)

        if normalized == "CW":
            return column

        compact = normalized.replace(" ", "")

        if compact == "CW":
            candidates.append(column)

    if candidates:
        return candidates[0]

    return None


# ============================================================
# EXTRACT HAWB FROM PDF FILE NAME
# ============================================================

def extract_hawb_from_filename(filename):

    filename = normalize_text(filename).upper()

    # Remove [123], [1], etc.
    filename = re.sub(
        r"\[\d+\]",
        "",
        filename
    )

    # Common formats:
    #
    # PTY0045653
    # I879513
    # J158916
    # 99210748500
    # 992-10748500
    # 406-06772010
    # 729-94678345

    patterns = [

        # Alphanumeric HAWB
        r"\b([A-Z]{1,5}\d{5,})\b",

        # Numeric AWB with hyphen
        r"\b(\d{3}-\d{5,})\b",

        # Numeric AWB without hyphen
        r"\b(\d{8,})\b",

    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            filename
        )

        if match:

            return normalize_hawb(
                match.group(1)
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

    lines = pdf_text.split("\n")

    # ========================================================
    # ORIGINAL PROVIDER
    # ========================================================

    for line in lines:

        match = re.search(
            r"\d+\s+\d+(?:\.\d+)?K\s+[A-Z]\s+(\d+(?:\.\d+)?)\s+\d+(?:\.\d+)?",
            line
        )

        if match:

            try:

                return float(
                    match.group(1)
                )

            except (ValueError, TypeError):

                pass

    # ========================================================
    # GENERIC / SVC PROVIDER
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

                except (ValueError, TypeError):

                    pass

    return None


# ============================================================
# READ EXCEL
# ============================================================

def read_excel_file(excel_file):

    header_row = find_excel_header(
        excel_file
    )

    df_excel = pd.read_excel(
        excel_file,
        header=header_row
    )

    # Clean column names
    cleaned_columns = []

    for column in df_excel.columns:

        cleaned_columns.append(
            normalize_text(column)
        )

    df_excel.columns = cleaned_columns

    return df_excel, header_row


# ============================================================
# VALIDATE AWB
# ============================================================

def validate_awb(
    input_folder,
    output_folder
):

    print("")
    print("======================================")
    print("STARTING AWB VALIDATION")
    print("======================================")

    # ========================================================
    # FIND EXCEL
    # ========================================================

    excel_file = None

    for file in os.listdir(input_folder):

        if file.lower().endswith(".xlsx"):

            # Do not use a previous generated result
            if file.lower() == "awb_validation_result.xlsx":
                continue

            excel_file = os.path.join(
                input_folder,
                file
            )

            break

    if not excel_file:

        raise Exception(
            "No Excel file found in the uploaded files."
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
        f"Header detected on Excel row: {header_row + 1}"
    )

    # ========================================================
    # SHOW COLUMNS
    # ========================================================

    print("")
    print("COLUMNS READ BY PANDAS:")
    print("--------------------------------------")

    for index, column in enumerate(
        df_excel.columns,
        start=1
    ):

        print(
            f"{index}: {repr(column)}"
        )

    print("--------------------------------------")

    # ========================================================
    # FIND AWB COLUMN
    # ========================================================

    awb_column = find_awb_column(
        df_excel.columns
    )

    if awb_column is None:

        columns_debug = "\n".join(
            [
                f"{i + 1}: {repr(column)}"
                for i, column in enumerate(
                    df_excel.columns
                )
            ]
        )

        raise Exception(
            "AWB/HAWB column could not be identified.\n\n"
            "Columns detected by pandas:\n\n"
            + columns_debug
        )

    print(
        f"AWB column detected: {repr(awb_column)}"
    )

    # ========================================================
    # FIND CW COLUMN
    # ========================================================

    cw_column = find_cw_column(
        df_excel.columns
    )

    if cw_column is None:

        columns_debug = "\n".join(
            [
                f"{i + 1}: {repr(column)}"
                for i, column in enumerate(
                    df_excel.columns
                )
            ]
        )

        raise Exception(
            "CW column could not be identified.\n\n"
            "Columns detected by pandas:\n\n"
            + columns_debug
        )

    print(
        f"CW column detected: {repr(cw_column)}"
    )

    # ========================================================
    # STANDARDIZE INTERNAL COLUMN NAMES
    # ========================================================

    df_excel = df_excel.rename(
        columns={
            awb_column: "HAWB",
            cw_column: "CW"
        }
    )

    # ========================================================
    # NORMALIZE HAWB
    # ========================================================

    df_excel["HAWB"] = df_excel[
        "HAWB"
    ].apply(
        normalize_hawb
    )

    # ========================================================
    # NORMALIZE CW
    # ========================================================

    df_excel["CW"] = pd.to_numeric(
        df_excel["CW"],
        errors="coerce"
    )

    # ========================================================
    # REMOVE EMPTY HAWB ROWS
    # ========================================================

    df_excel = df_excel[
        df_excel["HAWB"] != ""
    ].copy()

    print("")
    print(
        f"Excel shipment rows: {len(df_excel)}"
    )

    # ========================================================
    # CREATE PDF INDEX
    # ========================================================

    pdf_index = {}

    pdf_count = 0

    for file in os.listdir(input_folder):

        if file.lower().endswith(".pdf"):

            pdf_count += 1

            hawb = extract_hawb_from_filename(
                file
            )

            if hawb:

                pdf_index[hawb] = file

    print(
        f"PDF files uploaded: {pdf_count}"
    )

    print(
        f"PDFs with identifiable AWB/HAWB: {len(pdf_index)}"
    )

    # ========================================================
    # VALIDATION
    # ========================================================

    results = []

    for _, row in df_excel.iterrows():

        excel_hawb = normalize_hawb(
            row["HAWB"]
        )

        excel_cw = row["CW"]

        # ====================================================
        # PDF NOT FOUND
        # ====================================================

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

        # ====================================================
        # PDF FOUND
        # ====================================================

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

        # ====================================================
        # CW NOT FOUND
        # ====================================================

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

        # ====================================================
        # EXCEL CW EMPTY
        # ====================================================

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

        # ====================================================
        # CALCULATE DIFFERENCE
        # ====================================================

        difference = round(
            abs(
                float(excel_cw) - float(pdf_cw)
            ),
            2
        )

        # ====================================================
        # RESULT
        # ========================================================

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
    # EXTRA PDFS
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

    # ========================================================
    # SUMMARY
    # ========================================================

    print("")
    print("======================================")
    print("VALIDATION COMPLETED")
    print("======================================")

    print(
        f"Total results: {len(df_results)}"
    )

    if not df_results.empty:

        print(
            f"PASS: {(df_results['Result'] == 'PASS').sum()}"
        )

        print(
            f"FAIL: {(df_results['Result'] == 'FAIL').sum()}"
        )

        print(
            f"PDF NOT FOUND: "
            f"{(df_results['Result'] == 'PDF NOT FOUND').sum()}"
        )

        print(
            f"CW NOT FOUND IN PDF: "
            f"{(df_results['Result'] == 'CW NOT FOUND IN PDF').sum()}"
        )

    print(
        f"Output: {output_file}"
    )

    print("======================================")

    return output_file