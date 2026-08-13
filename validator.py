import pandas as pd
import pdfplumber
import os
import re


# =====================================================
# NORMALIZE TEXT
# =====================================================

def normalize_text(value):

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        return ""

    value = str(value)

    value = value.replace("\xa0", " ")
    value = value.replace("\n", " ")
    value = value.replace("\r", " ")
    value = value.replace("\t", " ")

    value = re.sub(r"\s+", " ", value)

    return value.strip().upper()


# =====================================================
# NORMALIZE HAWB / AWB
# =====================================================

def normalize_hawb(value):

    if value is None:
        return ""

    # Important:
    # Prevent pandas Series/DataFrame from
    # reaching this function.

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
# FIND EXCEL HEADER
# =====================================================

def find_excel_header(excel_file):

    print("========================================")
    print("READING EXCEL FOR DEBUG")
    print("========================================")

    # Read first 10 rows without headers
    debug_df = pd.read_excel(
        excel_file,
        header=None,
        nrows=10
    )

    print("")
    print("Excel shape for first 10 rows:")
    print(debug_df.shape)

    print("")

    for i in range(len(debug_df)):

        print(
            f"========== EXCEL ROW {i + 1} =========="
        )

        for j, value in enumerate(
            debug_df.iloc[i].tolist()
        ):

            print(
                f"Column {j + 1}: "
                f"value={repr(value)} | "
                f"type={type(value).__name__}"
            )

    print("")
    print("========================================")
    print("TEST MODE")
    print("Assuming Excel header is ROW 2")
    print("========================================")

    # Excel row 2 = pandas header=1
    return 1


# =====================================================
# FIND AWB COLUMN
# =====================================================

def find_awb_column(columns):

    print("")
    print("SEARCHING FOR AWB COLUMN...")
    print("")

    for column in columns:

        normalized = normalize_text(column)

        compact = (
            normalized
            .replace(" ", "")
            .replace("/", "")
            .replace("-", "")
            .replace("_", "")
            .replace("º", "")
            .replace("°", "")
        )

        print(
            f"Checking column: {repr(column)} "
            f"-> normalized: {repr(compact)}"
        )

        # Exact common cases
        if compact in [
            "AWB",
            "HAWB",
            "AWBBL",
            "AWBBLNO",
            "AWBBLN",
            "AWBHBL",
            "AWBHBLNO",
            "AWBHBLN",
            "HAWBBL",
            "HAWBBLNO",
            "HAWBBLN",
        ]:

            print(
                f"AWB COLUMN FOUND: {column}"
            )

            return column

        # Flexible cases
        if compact.startswith("AWB"):

            print(
                f"AWB COLUMN FOUND: {column}"
            )

            return column

        if compact.startswith("HAWB"):

            print(
                f"HAWB COLUMN FOUND: {column}"
            )

            return column

    return None


# =====================================================
# FIND CW COLUMN
# =====================================================

def find_cw_column(columns):

    print("")
    print("SEARCHING FOR CW COLUMN...")
    print("")

    for column in columns:

        normalized = normalize_text(column)

        compact = (
            normalized
            .replace(" ", "")
            .replace("/", "")
            .replace("-", "")
            .replace("_", "")
        )

        print(
            f"Checking column: {repr(column)} "
            f"-> normalized: {repr(compact)}"
        )

        if compact == "CW":

            print(
                f"CW COLUMN FOUND: {column}"
            )

            return column

        if compact == "CHARGEABLEWEIGHT":

            print(
                f"CW COLUMN FOUND: {column}"
            )

            return column

    return None


# =====================================================
# EXTRACT HAWB FROM PDF FILE NAME
# =====================================================

def extract_hawb_from_filename(filename):

    # Remove things like [1], [2], etc.
    filename = re.sub(
        r"\[\d+\]",
        "",
        filename
    )

    filename_upper = filename.upper()

    # -------------------------------------------------
    # Alphanumeric HAWB
    #
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
    # Numeric AWB
    #
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
            f"ERROR READING PDF: {pdf_path}"
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

    print("")
    print("========================================")
    print("STARTING AWB VALIDATION")
    print("========================================")

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

    print("")
    print(
        f"Excel found: {excel_file}"
    )

    # =================================================
    # FIND HEADER
    # =================================================

    header_row = find_excel_header(
        excel_file
    )

    print("")
    print(
        f"Using header row index: "
        f"{header_row}"
    )

    print(
        "This corresponds to Excel row: "
        f"{header_row + 1}"
    )

    # =================================================
    # READ EXCEL
    # =================================================

    print("")
    print("READING EXCEL WITH HEADER...")
    print("")

    df_excel = pd.read_excel(
        excel_file,
        header=header_row
    )

    # =================================================
    # SHOW COLUMNS
    # =================================================

    print("")
    print("========================================")
    print("COLUMNS READ BY PANDAS")
    print("========================================")

    for i, column in enumerate(
        df_excel.columns
    ):

        print(
            f"{i + 1}. "
            f"{repr(column)} "
            f"(type={type(column).__name__})"
        )

    print("========================================")

    # =================================================
    # FIND AWB COLUMN
    # =================================================

    awb_column = find_awb_column(
        df_excel.columns
    )

    if awb_column is None:

        raise Exception(
            "AWB/HAWB column could not be identified.\n\n"
            "Check the Streamlit logs under "
            "COLUMNS READ BY PANDAS."
        )

    # =================================================
    # FIND CW COLUMN
    # =================================================

    cw_column = find_cw_column(
        df_excel.columns
    )

    if cw_column is None:

        raise Exception(
            "CW column could not be identified.\n\n"
            "Check the Streamlit logs under "
            "COLUMNS READ BY PANDAS."
        )

    print("")
    print(
        f"FINAL AWB COLUMN: {repr(awb_column)}"
    )

    print(
        f"FINAL CW COLUMN: {repr(cw_column)}"
    )

    # =================================================
    # CREATE INTERNAL COLUMNS
    # =================================================

    df_excel["HAWB_INTERNAL"] = (
        df_excel[awb_column]
        .apply(normalize_hawb)
    )

    df_excel["CW_INTERNAL"] = pd.to_numeric(
        df_excel[cw_column],
        errors="coerce"
    )

    # =================================================
    # PDF INDEX
    # =================================================

    pdf_index = {}

    for file in os.listdir(input_folder):

        if file.lower().endswith(".pdf"):

            hawb = extract_hawb_from_filename(
                file
            )

            if hawb:

                pdf_index[hawb] = file

    print("")
    print(
        f"PDFs found and indexed: "
        f"{len(pdf_index)}"
    )

    # =================================================
    # VALIDATE
    # =================================================

    results = []

    for _, row in df_excel.iterrows():

        excel_hawb = row[
            "HAWB_INTERNAL"
        ]

        excel_cw = row[
            "CW_INTERNAL"
        ]

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
        # EXCEL CW EMPTY
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

        # -------------------------------------------------
        # COMPARE
        # -------------------------------------------------

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

    # =================================================
    # EXTRA PDFs
    # =================================================

    excel_hawb_set = set(
        df_excel["HAWB_INTERNAL"]
    )

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

    print("")
    print("========================================")
    print("VALIDATION COMPLETED")
    print("========================================")

    print(
        f"Output: {output_file}"
    )

    return output_file