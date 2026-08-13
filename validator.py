import pandas as pd
import pdfplumber
import os
import re


# =====================================================
# TEXT NORMALIZATION
# =====================================================

def normalize_text(value):

    if pd.isna(value):
        return ""

    text = str(value)

    # Replace non-breaking spaces
    text = text.replace("\xa0", " ")

    # Normalize multiple spaces
    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip().upper()


# =====================================================
# NORMALIZE HAWB
# =====================================================

def normalize_hawb(value):

    if pd.isna(value):
        return ""

    return (
        str(value)
        .strip()
        .upper()
        .replace(" ", "")
        .replace("-", "")
    )


# =====================================================
# NORMALIZE COLUMN NAME
# =====================================================

def normalize_column_name(value):

    return normalize_text(value)


# =====================================================
# IDENTIFY HAWB / AWB COLUMN
# =====================================================

def is_hawb_column(value):

    text = normalize_column_name(value)

    if not text:
        return False

    # Remove spaces and punctuation
    simplified = re.sub(
        r"[^A-Z0-9]",
        "",
        text
    )

    possible_names = {

        "AWB",
        "AWBNO",
        "AWBNUMBER",

        "HAWB",
        "HAWBNO",
        "HAWBNUMBER",

        "AWBBL",
        "AWBBLNO",
        "AWBBLNUMBER",

    }

    return simplified in possible_names


# =====================================================
# IDENTIFY CW COLUMN
# =====================================================

def is_cw_column(value):

    text = normalize_column_name(value)

    if not text:
        return False

    simplified = re.sub(
        r"[^A-Z0-9]",
        "",
        text
    )

    possible_names = {

        "CW",
        "CWT",

        "CHARGEABLEWEIGHT",
        "CHARGEABLEWT",
        "CHARGEABLEWEIGHTKG",

    }

    return simplified in possible_names


# =====================================================
# FIND EXCEL HEADER ROW
# =====================================================

def find_excel_header(excel_file):

    # Read first 30 rows without assuming a header
    df = pd.read_excel(
        excel_file,
        header=None,
        nrows=30
    )

    for i in range(len(df)):

        row = df.iloc[i].tolist()

        hawb_found = False
        cw_found = False

        for value in row:

            if is_hawb_column(value):
                hawb_found = True

            if is_cw_column(value):
                cw_found = True

        # Header must contain BOTH:
        # HAWB/AWB/AWB-BL
        # AND
        # CW
        if hawb_found and cw_found:

            print(
                f"Excel header detected on row {i + 1}"
            )

            return i

    # =================================================
    # DIAGNOSTIC INFORMATION
    # =================================================

    print(
        "Unable to locate Excel header row."
    )

    print(
        "Rows detected in Excel:"
    )

    for i in range(min(len(df), 30)):

        row_values = [
            normalize_column_name(x)
            for x in df.iloc[i].tolist()
        ]

        print(
            f"Row {i + 1}: {row_values}"
        )

    raise Exception(
        "Unable to locate Excel header row. "
        "The file must contain an AWB/HAWB "
        "identification column and a CW column."
    )


# =====================================================
# NORMALIZE EXCEL COLUMNS
# =====================================================

def normalize_excel_columns(df):

    normalized_columns = []

    for column in df.columns:

        if is_hawb_column(column):

            normalized_columns.append(
                "HAWB"
            )

        elif is_cw_column(column):

            normalized_columns.append(
                "CW"
            )

        else:

            normalized_columns.append(
                normalize_column_name(column)
            )

    df.columns = normalized_columns

    return df


# =====================================================
# EXTRACT HAWB FROM PDF FILE NAME
# =====================================================

def extract_hawb_from_filename(filename):

    # Remove extension
    filename_without_extension = os.path.splitext(
        filename
    )[0]

    filename_upper = filename_without_extension.upper()

    # -------------------------------------------------
    # PRIORITY 1
    # Look specifically after HAWB No / HAWB Number
    # -------------------------------------------------

    match = re.search(
        r"HAWB\s*(?:NO|NUMBER)?\s*[_:\-]?\s*([A-Z0-9]{5,})",
        filename_upper
    )

    if match:

        return normalize_hawb(
            match.group(1)
        )

    # -------------------------------------------------
    # PRIORITY 2
    # Existing general pattern
    # -------------------------------------------------

    match = re.search(
        r"([A-Z]{1,5}\d{5,}|[0-9]{3}-?[0-9]{5,})",
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
            f"Error reading PDF {pdf_path}: {e}"
        )

        return None

    lines = pdf_text.split("\n")

    # =================================================
    # ORIGINAL PROVIDER
    # =================================================

    for line in lines:

        match = re.search(
            r'\d+\s+\d+(?:\.\d+)?K\s+[A-Z]\s+(\d+(?:\.\d+)?)\s+\d+(?:\.\d+)?',
            line
        )

        if match:

            try:

                return float(
                    match.group(1)
                )

            except Exception:

                pass

    # =================================================
    # SVC / OTHER PROVIDER
    # =================================================

    for line in lines:

        upper_line = line.upper()

        if (
            "CHARGEABLE WEIGHT" in upper_line
            or "CHARGEABLE" in upper_line
            or "CWT" in upper_line
        ):

            numbers = re.findall(
                r'\d+(?:\.\d+)?',
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


# =====================================================
# VALIDATE AWB
# =====================================================

def validate_awb(
    input_folder,
    output_folder
):

    # =================================================
    # FIND EXCEL FILE
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
        f"Header detected on row: {header_row + 1}"
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

    df_excel = normalize_excel_columns(
        df_excel
    )

    print(
        "Detected Excel columns:"
    )

    print(
        list(df_excel.columns)
    )

    # =================================================
    # VALIDATE HAWB COLUMN
    # =================================================

    if "HAWB" not in df_excel.columns:

        raise Exception(
            "AWB / HAWB / AWB-BL column not found."
        )

    # =================================================
    # VALIDATE CW COLUMN
    # =================================================

    if "CW" not in df_excel.columns:

        raise Exception(
            "CW / Chargeable Weight column not found."
        )

    # =================================================
    # NORMALIZE HAWB VALUES
    # =================================================

    df_excel["HAWB"] = df_excel["HAWB"].apply(
        normalize_hawb
    )

    # =================================================
    # NORMALIZE CW VALUES
    # =================================================

    df_excel["CW"] = pd.to_numeric(
        df_excel["CW"],
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
                    f"PDF indexed: {hawb} -> {file}"
                )

    print(
        f"PDFs found: {len(pdf_index)}"
    )

    # =================================================
    # VALIDATE EXCEL ROWS
    # =================================================

    results = []

    for _, row in df_excel.iterrows():

        excel_hawb = row["HAWB"]
        excel_cw = row["CW"]

        # Skip empty rows
        if not excel_hawb:

            continue

        # =================================================
        # PDF NOT FOUND
        # =================================================

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

        # =================================================
        # GET PDF
        # =================================================

        pdf_file = pdf_index[
            excel_hawb
        ]

        pdf_path = os.path.join(
            input_folder,
            pdf_file
        )

        # =================================================
        # EXTRACT PDF CW
        # =================================================

        pdf_cw = extract_pdf_cw(
            pdf_path
        )

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

        # =================================================
        # CALCULATE DIFFERENCE
        # =================================================

        try:

            difference = round(
                abs(
                    float(excel_cw) - pdf_cw
                ),
                2
            )

        except Exception:

            difference = ""

        # =================================================
        # PASS / FAIL
        # =================================================

        if (
            difference != ""
            and difference <= 0.01
        ):

            result = "PASS"

        else:

            result = "FAIL"

        # =================================================
        # SAVE RESULT
        # =================================================

        results.append({

            "HAWB": excel_hawb,
            "Excel CW": excel_cw,
            "PDF CW": pdf_cw,
            "Difference": difference,
            "Result": result,
            "PDF File": pdf_file

        })

    # =================================================
    # EXTRA PDFS
    # =================================================

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
        "Validation completed:",
        output_file
    )

    return output_file