import pandas as pd
import pdfplumber
import os
import re


# ============================================================
# TEXT NORMALIZATION
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
# NORMALIZE HEADER
# ============================================================

def normalize_header(value):

    text = clean_text(value).upper()

    # Remove accents / special characters
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

    # Remove spaces
    text = re.sub(r"\s+", "", text)

    # Remove punctuation except slash
    text = text.replace("_", "")
    text = text.replace("-", "")

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
        result = pd.isna(value)

        if isinstance(result, bool) and result:
            return ""

    except Exception:
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
# IDENTIFY AWB HEADER
# ============================================================

def is_awb_header(value):

    h = normalize_header(value)

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

    if h in possible:
        return True

    # Flexible checks
    if h.startswith("AWB") and (
        "NO" in h
        or "BL" in h
        or "HBL" in h
    ):
        return True

    if h.startswith("HAWB"):
        return True

    return False


# ============================================================
# IDENTIFY CW HEADER
# ============================================================

def is_cw_header(value):

    h = normalize_header(value)

    return h == "CW"


# ============================================================
# FIND EXCEL HEADER
# ============================================================

def find_excel_header(excel_file):

    # Read WITHOUT assuming any header
    raw = pd.read_excel(
        excel_file,
        header=None,
        nrows=30
    )

    diagnostics = []

    for row_number in range(len(raw)):

        values = raw.iloc[row_number].tolist()

        cleaned = [
            clean_text(x)
            for x in values
        ]

        normalized = [
            normalize_header(x)
            for x in cleaned
        ]

        awb_positions = [
            i + 1
            for i, value in enumerate(cleaned)
            if is_awb_header(value)
        ]

        cw_positions = [
            i + 1
            for i, value in enumerate(cleaned)
            if is_cw_header(value)
        ]

        # Store diagnostic information
        diagnostics.append(
            f"Excel row {row_number + 1}: "
            f"AWB={awb_positions}, "
            f"CW={cw_positions}"
        )

        # The ideal header contains both
        if awb_positions and cw_positions:

            print(
                f"HEADER FOUND: Excel row {row_number + 1}"
            )

            return row_number

    # ========================================================
    # HEADER NOT FOUND
    # ========================================================

    diagnostic_text = "\n".join(
        diagnostics
    )

    raise Exception(
        "Unable to locate Excel header row.\n\n"
        "The application searched the first 30 rows.\n\n"
        "Detection results:\n"
        + diagnostic_text
    )


# ============================================================
# FIND AWB COLUMN
# ============================================================

def find_awb_column(columns):

    for column in columns:

        if is_awb_header(column):
            return column

    return None


# ============================================================
# FIND CW COLUMN
# ============================================================

def find_cw_column(columns):

    for column in columns:

        if is_cw_header(column):
            return column

    return None


# ============================================================
# EXTRACT HAWB FROM PDF FILE NAME
# ============================================================

def extract_hawb_from_filename(filename):

    filename = clean_text(filename).upper()

    filename = re.sub(
        r"\[\d+\]",
        "",
        filename
    )

    patterns = [

        # Example:
        # PTY0045653
        # I879513
        # J158916
        r"\b([A-Z]{1,5}\d{5,})\b",

        # Example:
        # 992-10748500
        # 406-06772010
        r"\b(\d{3}-\d{5,})\b",

        # Example:
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

    text = ""

    try:

        with pdfplumber.open(pdf_path) as pdf:

            for page in pdf.pages:

                page_text = page.extract_text()

                if page_text:
                    text += page_text + "\n"

    except Exception:

        return None

    lines = text.splitlines()

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
                return float(match.group(1))
            except Exception:
                pass

    # --------------------------------------------------------
    # OTHER PROVIDERS
    # --------------------------------------------------------

    for line in lines:

        upper = line.upper()

        if (
            "CHARGEABLE WEIGHT" in upper
            or "CHARGEABLE" in upper
            or "CWT" in upper
        ):

            numbers = re.findall(
                r"\d+(?:\.\d+)?",
                line
            )

            if numbers:

                try:
                    return float(numbers[-1])
                except Exception:
                    pass

    return None


# ============================================================
# VALIDATE AWB
# ============================================================

def validate_awb(
    input_folder,
    output_folder
):

    print("================================")
    print("AWB VALIDATOR START")
    print("================================")

    # ========================================================
    # FIND EXCEL
    # ========================================================

    excel_file = None

    for file in os.listdir(input_folder):

        if not file.lower().endswith(".xlsx"):
            continue

        if file.lower() == "awb_validation_result.xlsx":
            continue

        excel_file = os.path.join(
            input_folder,
            file
        )

        break

    if excel_file is None:

        raise Exception(
            "No Excel file found."
        )

    print(
        "Excel:",
        excel_file
    )

    # ========================================================
    # FIND HEADER
    # ========================================================

    header_row = find_excel_header(
        excel_file
    )

    print(
        "Header row:",
        header_row + 1
    )

    # ========================================================
    # READ EXCEL
    # ========================================================

    df = pd.read_excel(
        excel_file,
        header=header_row
    )

    # Clean column names
    df.columns = [
        clean_text(column)
        for column in df.columns
    ]

    print("")
    print("COLUMNS:")
    print("--------------------------------")

    for i, column in enumerate(
        df.columns,
        start=1
    ):

        print(
            i,
            repr(column)
        )

    print("--------------------------------")

    # ========================================================
    # FIND AWB
    # ========================================================

    awb_column = find_awb_column(
        df.columns
    )

    if awb_column is None:

        raise Exception(
            "AWB column was not found after reading "
            "the detected header row.\n\n"
            "Columns detected:\n\n"
            + "\n".join(
                [
                    f"{i}: {repr(x)}"
                    for i, x in enumerate(
                        df.columns,
                        start=1
                    )
                ]
            )
        )

    # ========================================================
    # FIND CW
    # ========================================================

    cw_column = find_cw_column(
        df.columns
    )

    if cw_column is None:

        raise Exception(
            "CW column was not found after reading "
            "the detected header row.\n\n"
            "Columns detected:\n\n"
            + "\n".join(
                [
                    f"{i}: {repr(x)}"
                    for i, x in enumerate(
                        df.columns,
                        start=1
                    )
                ]
            )
        )

    print(
        "AWB column:",
        repr(awb_column)
    )

    print(
        "CW column:",
        repr(cw_column)
    )

    # ========================================================
    # STANDARDIZE
    # ========================================================

    df["HAWB"] = df[
        awb_column
    ].apply(
        normalize_hawb
    )

    df["CW"] = pd.to_numeric(
        df[cw_column],
        errors="coerce"
    )

    # Remove empty rows
    df = df[
        df["HAWB"] != ""
    ].copy()

    # ========================================================
    # PDF INDEX
    # ========================================================

    pdf_index = {}

    for file in os.listdir(input_folder):

        if not file.lower().endswith(".pdf"):
            continue

        hawb = extract_hawb_from_filename(
            file
        )

        if hawb:

            pdf_index[hawb] = file

    print(
        "PDFs:",
        len(pdf_index)
    )

    # ========================================================
    # VALIDATE
    # ========================================================

    results = []

    for _, row in df.iterrows():

        hawb = normalize_hawb(
            row["HAWB"]
        )

        excel_cw = row["CW"]

        # PDF not found
        if hawb not in pdf_index:

            results.append({

                "HAWB": hawb,
                "Excel CW": excel_cw,
                "PDF CW": "",
                "Difference": "",
                "Result": "PDF NOT FOUND",
                "PDF File": ""

            })

            continue

        pdf_file = pdf_index[
            hawb
        ]

        pdf_path = os.path.join(
            input_folder,
            pdf_file
        )

        pdf_cw = extract_pdf_cw(
            pdf_path
        )

        # CW not found
        if pdf_cw is None:

            results.append({

                "HAWB": hawb,
                "Excel CW": excel_cw,
                "PDF CW": "",
                "Difference": "",
                "Result": "CW NOT FOUND IN PDF",
                "PDF File": pdf_file

            })

            continue

        # Excel CW missing
        if pd.isna(excel_cw):

            results.append({

                "HAWB": hawb,
                "Excel CW": "",
                "PDF CW": pdf_cw,
                "Difference": "",
                "Result": "CW NOT FOUND IN EXCEL",
                "PDF File": pdf_file

            })

            continue

        # Difference
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

            "HAWB": hawb,
            "Excel CW": excel_cw,
            "PDF CW": pdf_cw,
            "Difference": difference,
            "Result": result,
            "PDF File": pdf_file

        })

    # ========================================================
    # PDFS NOT IN EXCEL
    # ========================================================

    excel_hawbs = set(
        df["HAWB"]
    )

    for pdf_hawb, pdf_file in pdf_index.items():

        if pdf_hawb not in excel_hawbs:

            results.append({

                "HAWB": pdf_hawb,
                "Excel CW": "",
                "PDF CW": "",
                "Difference": "",
                "Result": "HAWB NOT FOUND IN EXCEL",
                "PDF File": pdf_file

            })

    # ========================================================
    # SAVE
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

    print("")
    print("================================")
    print("VALIDATION FINISHED")
    print("Output:")
    print(output_file)
    print("================================")

    return output_file