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

    text = text.replace("\xa0", " ")

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
# SIMPLIFY COLUMN NAME
# =====================================================

def simplify_column_name(value):

    text = normalize_text(value)

    return re.sub(
        r"[^A-Z0-9]",
        "",
        text
    )


# =====================================================
# IDENTIFY HAWB / AWB COLUMN
# =====================================================

def get_hawb_column_type(value):

    simplified = simplify_column_name(
        value
    )

    if not simplified:
        return None

    # Highest priority
    if simplified in {
        "HAWB",
        "HAWBNO",
        "HAWBNUMBER"
    }:
        return "HAWB"

    # Second priority
    if simplified in {
        "AWB",
        "AWBNO",
        "AWBNUMBER"
    }:
        return "AWB"

    # Third priority
    if simplified in {
        "AWBBL",
        "AWBBLN",
        "AWBBLNO",
        "AWBBLNUMBER"
    }:
        return "AWB_BL"

    return None


# =====================================================
# IDENTIFY CW COLUMN
# =====================================================

def is_cw_column(value):

    simplified = simplify_column_name(
        value
    )

    return simplified in {

        "CW",
        "CWT",
        "CHARGEABLEWEIGHT",
        "CHARGEABLEWT",
        "CHARGEABLEWEIGHTKG"

    }


# =====================================================
# FIND EXCEL HEADER ROW
# =====================================================

def find_excel_header(excel_file):

    df = pd.read_excel(
        excel_file,
        header=None,
        nrows=30
    )

    for i in range(len(df)):

        row = df.iloc[i].tolist()

        identifier_found = False
        cw_found = False

        for value in row:

            if get_hawb_column_type(value):
                identifier_found = True

            if is_cw_column(value):
                cw_found = True

        if identifier_found and cw_found:

            print(
                f"Excel header detected on row {i + 1}"
            )

            return i

    # Diagnostic output

    print(
        "Unable to locate Excel header row."
    )

    for i in range(
        min(len(df), 30)
    ):

        row_values = [
            normalize_text(x)
            for x in df.iloc[i].tolist()
        ]

        print(
            f"Row {i + 1}: {row_values}"
        )

    raise Exception(
        "Unable to locate Excel header row. "
        "The file must contain an AWB/HAWB/AWB-BL "
        "identification column and a CW column."
    )


# =====================================================
# FIND IDENTIFICATION COLUMN
# =====================================================

def find_identification_column(df):

    candidates = []

    for index, column in enumerate(
        df.columns
    ):

        column_type = get_hawb_column_type(
            column
        )

        if column_type:

            candidates.append(
                (
                    index,
                    column,
                    column_type
                )
            )

    if not candidates:

        raise Exception(
            "AWB / HAWB / AWB-BL column not found."
        )

    # =================================================
    # PRIORITY
    # =================================================
    #
    # HAWB
    # AWB
    # AWB/BL
    #

    priority = {
        "HAWB": 1,
        "AWB": 2,
        "AWB_BL": 3
    }

    candidates.sort(
        key=lambda x: priority[x[2]]
    )

    selected = candidates[0]

    print(
        f"Identification column selected: "
        f"{selected[1]} "
        f"({selected[2]})"
    )

    return selected[1]


# =====================================================
# FIND CW COLUMN
# =====================================================

def find_cw_column(df):

    candidates = []

    for column in df.columns:

        if is_cw_column(column):

            candidates.append(
                column
            )

    if not candidates:

        raise Exception(
            "CW / Chargeable Weight column not found."
        )

    selected = candidates[0]

    print(
        f"CW column selected: {selected}"
    )

    return selected


# =====================================================
# EXTRACT HAWB FROM PDF FILE NAME
# =====================================================

def extract_hawb_from_filename(filename):

    filename_without_extension = (
        os.path.splitext(filename)[0]
    )

    filename_upper = (
        filename_without_extension.upper()
    )

    # =================================================
    # HAWB No_ PTY0045653
    # HAWB NO PTY0045653
    # HAWB NUMBER PTY0045653
    # =================================================

    match = re.search(
        r"HAWB\s*(?:NO|NUMBER)?\s*[_:\-]?\s*([A-Z0-9]{5,})",
        filename_upper
    )

    if match:

        return normalize_hawb(
            match.group(1)
        )

    # =================================================
    # GENERAL PATTERN
    # =================================================

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

        with pdfplumber.open(
            pdf_path
        ) as pdf:

            for page in pdf.pages:

                text = page.extract_text()

                if text:

                    pdf_text += (
                        text + "\n"
                    )

    except Exception as e:

        print(
            f"Error reading PDF {pdf_path}: {e}"
        )

        return None

    lines = pdf_text.split(
        "\n"
    )

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
    # OTHER PROVIDERS
    # =================================================

    for line in lines:

        upper_line = line.upper()

        if (
            "CHARGEABLE WEIGHT"
            in upper_line
            or
            "CHARGEABLE"
            in upper_line
            or
            "CWT"
            in upper_line
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
    # FIND EXCEL
    # =================================================

    excel_file = None

    for file in os.listdir(
        input_folder
    ):

        if file.lower().endswith(
            ".xlsx"
        ):

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

    # =================================================
    # READ EXCEL
    # =================================================

    df_excel = pd.read_excel(
        excel_file,
        header=header_row
    )

    print(
        "Original Excel columns:"
    )

    print(
        list(df_excel.columns)
    )

    # =================================================
    # FIND IDENTIFICATION COLUMN
    # =================================================

    hawb_column = find_identification_column(
        df_excel
    )

    # =================================================
    # FIND CW COLUMN
    # =================================================

    cw_column = find_cw_column(
        df_excel
    )

    # =================================================
    # CREATE INTERNAL COLUMNS
    # =================================================

    df_excel["HAWB_INTERNAL"] = (
        df_excel[hawb_column]
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

    for file in os.listdir(
        input_folder
    ):

        if file.lower().endswith(
            ".pdf"
        ):

            hawb = (
                extract_hawb_from_filename(
                    file
                )
            )

            if hawb:

                pdf_index[hawb] = file

                print(
                    f"PDF indexed: "
                    f"{hawb} -> {file}"
                )

    print(
        f"PDFs indexed: "
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
                "Result":
                    "CW NOT FOUND IN PDF",
                "PDF File": pdf_file

            })

            continue

        # =================================================
        # CALCULATE DIFFERENCE
        # =================================================

        try:

            difference = round(
                abs(
                    float(excel_cw)
                    - float(pdf_cw)
                ),
                2
            )

        except Exception:

            difference = ""

        # =================================================
        # RESULT
        # =================================================

        if (
            difference != ""
            and difference <= 0.01
        ):

            result = "PASS"

        else:

            result = "FAIL"

        # =================================================
        # ADD RESULT
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
    # EXTRA PDFs
    # =================================================

    excel_hawb_set = set(
        df_excel[
            "HAWB_INTERNAL"
        ]
    )

    for (
        pdf_hawb,
        pdf_file
    ) in pdf_index.items():

        if pdf_hawb not in (
            excel_hawb_set
        ):

            results.append({

                "HAWB": pdf_hawb,
                "Excel CW": "",
                "PDF CW": "",
                "Difference": "",
                "Result":
                    "HAWB NOT FOUND IN EXCEL",
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