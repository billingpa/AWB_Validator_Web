import pandas as pd
import pdfplumber
import os
import re


# ============================================================
# BASIC TEXT CLEANING
# ============================================================

def clean_text(value):

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    text = str(value)

    text = (
        text
        .replace("\xa0", " ")
        .replace("\u200b", "")
        .replace("\u200c", "")
        .replace("\u200d", "")
        .replace("\ufeff", "")
        .replace("\n", " ")
        .replace("\r", " ")
        .replace("\t", " ")
    )

    text = re.sub(r"\s+", " ", text)

    return text.strip()


# ============================================================
# NORMALIZE COLUMN NAME
# ============================================================

def normalize_column(value):

    text = clean_text(value).upper()

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

    # Remove all spaces
    text = re.sub(r"\s+", "", text)

    return text


# ============================================================
# NORMALIZE HAWB / AWB
# ============================================================

def normalize_hawb(value):

    if value is None:
        return ""

    if isinstance(value, (list, tuple, set, dict)):
        return ""

    try:

        missing = pd.isna(value)

        if isinstance(missing, bool) and missing:
            return ""

    except Exception:
        pass

    text = str(value)

    return (
        text
        .strip()
        .upper()
        .replace("\xa0", "")
        .replace(" ", "")
        .replace("-", "")
        .replace("_", "")
    )


# ============================================================
# IDENTIFY AWB COLUMN
# ============================================================

def is_awb_column(column):

    value = normalize_column(column)

    possible = {
        "AWB",
        "HAWB",
        "AWB/BL",
        "AWB/BLNO",
        "AWB/HBL",
        "AWB/HBLNO",
        "AWBNO",
        "HAWBNO",
        "AWBHBLNO",
    }

    if value in possible:
        return True

    # Flexible cases
    if value.startswith("AWB"):

        if (
            "BL" in value
            or "HBL" in value
            or "NO" in value
        ):
            return True

    if value.startswith("HAWB"):
        return True

    return False


# ============================================================
# IDENTIFY CW COLUMN
# ============================================================

def is_cw_column(column):

    value = normalize_column(column)

    return value == "CW"


# ============================================================
# EXTRACT HAWB FROM PDF FILE NAME
# ============================================================

def extract_hawb_from_filename(filename):

    filename = clean_text(filename).upper()

    # Remove things like [1], [2], etc.
    filename = re.sub(
        r"\[\d+\]",
        "",
        filename
    )

    patterns = [

        # PTY0045653
        # I879513
        # J158916
        r"\b([A-Z]{1,5}\d{5,})\b",

        # 992-10748500
        # 406-06772010
        r"\b(\d{3}-\d{5,})\b",

        # 99210748500
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

    # --------------------------------------------------------
    # ORIGINAL PROVIDER
    # --------------------------------------------------------

    for line in lines:

        match = re.search(
            r"\d+\s+\d+(?:\.\d+)?K\s+[A-Z]\s+"
            r"(\d+(?:\.\d+)?)\s+\d+(?:\.\d+)?",
            line
        )

        if match:

            try:

                return float(
                    match.group(1)
                )

            except Exception:
                pass

    # --------------------------------------------------------
    # OTHER PROVIDERS
    # --------------------------------------------------------

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
# FIND EXCEL FILE
# ============================================================

def find_excel_file(input_folder):

    for file in os.listdir(input_folder):

        if not file.lower().endswith(".xlsx"):
            continue

        if file.lower() == "awb_validation_result.xlsx":
            continue

        return os.path.join(
            input_folder,
            file
        )

    return None


# ============================================================
# READ EXCEL
#
# IMPORTANT:
# The test file has the header on Excel row 2.
#
# pandas uses zero-based indexing:
# Excel row 2 = header=1
# ============================================================

def read_excel_file(excel_file):

    HEADER_ROW = 1

    df = pd.read_excel(
        excel_file,
        header=HEADER_ROW
    )

    # Clean all column names
    df.columns = [
        clean_text(column)
        for column in df.columns
    ]

    return df


# ============================================================
# FIND COLUMN
# ============================================================

def find_awb_column(columns):

    for column in columns:

        if is_awb_column(column):

            return column

    return None


def find_cw_column(columns):

    for column in columns:

        if is_cw_column(column):

            return column

    return None


# ============================================================
# VALIDATE
# ============================================================

def validate_awb(
    input_folder,
    output_folder
):

    print("")
    print("======================================")
    print("STARTING AWB VALIDATOR")
    print("======================================")

    # ========================================================
    # FIND EXCEL
    # ========================================================

    excel_file = find_excel_file(
        input_folder
    )

    if excel_file is None:

        raise Exception(
            "No Excel file found."
        )

    print(
        "Excel file:",
        excel_file
    )

    # ========================================================
    # READ EXCEL
    # ========================================================

    df = read_excel_file(
        excel_file
    )

    print("")
    print("COLUMNS DETECTED:")
    print("--------------------------------------")

    for number, column in enumerate(
        df.columns,
        start=1
    ):

        print(
            f"{number}: {repr(column)}"
        )

    print("--------------------------------------")

    # ========================================================
    # FIND AWB
    # ========================================================

    awb_column = find_awb_column(
        df.columns
    )

    if awb_column is None:

        detected_columns = "\n".join(
            [
                f"{i}: {repr(column)}"
                for i, column in enumerate(
                    df.columns,
                    start=1
                )
            ]
        )

        raise Exception(
            "AWB column could not be identified.\n\n"
            "Columns read from Excel:\n\n"
            + detected_columns
        )

    print(
        "AWB column:",
        repr(awb_column)
    )

    # ========================================================
    # FIND CW
    # ========================================================

    cw_column = find_cw_column(
        df.columns
    )

    if cw_column is None:

        detected_columns = "\n".join(
            [
                f"{i}: {repr(column)}"
                for i, column in enumerate(
                    df.columns,
                    start=1
                )
            ]
        )

        raise Exception(
            "CW column could not be identified.\n\n"
            "Columns read from Excel:\n\n"
            + detected_columns
        )

    print(
        "CW column:",
        repr(cw_column)
    )

    # ========================================================
    # STANDARDIZE DATA
    # ========================================================

    df["VALIDATOR_HAWB"] = df[
        awb_column
    ].apply(
        normalize_hawb
    )

    df["VALIDATOR_CW"] = pd.to_numeric(
        df[cw_column],
        errors="coerce"
    )

    # Remove empty AWB rows
    df = df[
        df["VALIDATOR_HAWB"] != ""
    ].copy()

    print("")
    print(
        "Excel records:",
        len(df)
    )

    # ========================================================
    # CREATE PDF INDEX
    # ========================================================

    pdf_index = {}

    pdf_count = 0

    for file in os.listdir(input_folder):

        if not file.lower().endswith(".pdf"):
            continue

        pdf_count += 1

        hawb = extract_hawb_from_filename(
            file
        )

        if hawb:

            pdf_index[hawb] = file

    print(
        "PDF files:",
        pdf_count
    )

    print(
        "PDFs identified:",
        len(pdf_index)
    )

    # ========================================================
    # VALIDATE EXCEL RECORDS
    # ========================================================

    results = []

    for _, row in df.iterrows():

        excel_hawb = row[
            "VALIDATOR_HAWB"
        ]

        excel_cw = row[
            "VALIDATOR_CW"
        ]

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
        # PDF CW NOT FOUND
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
        # EXCEL CW NOT FOUND
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

        # ----------------------------------------------------
        # DIFFERENCE
        # ----------------------------------------------------

        difference = round(
            abs(
                float(excel_cw)
                -
                float(pdf_cw)
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
    # PDFS NOT IN EXCEL
    # ========================================================

    excel_hawb_set = set(
        df["VALIDATOR_HAWB"]
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
    # CREATE OUTPUT
    # ========================================================

    os.makedirs(
        output_folder,
        exist_ok=True
    )

    output_file = os.path.join(
        output_folder,
        "AWB_Validation_Result.xlsx"
    )

    result_df = pd.DataFrame(
        results
    )

    result_df.to_excel(
        output_file,
        index=False
    )

    # ========================================================
    # FINISHED
    # ========================================================

    print("")
    print("======================================")
    print("VALIDATION COMPLETED")
    print("======================================")

    print(
        "Output:",
        output_file
    )

    print(
        "Total results:",
        len(result_df)
    )

    if not result_df.empty:

        print(
            "PASS:",
            (
                result_df["Result"] == "PASS"
            ).sum()
        )

        print(
            "FAIL:",
            (
                result_df["Result"] == "FAIL"
            ).sum()
        )

        print(
            "PDF NOT FOUND:",
            (
                result_df["Result"] == "PDF NOT FOUND"
            ).sum()
        )

    print("======================================")

    return output_file