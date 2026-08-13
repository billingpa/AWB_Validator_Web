import pandas as pd
import pdfplumber
import os
import re


# ============================================================
# CONFIGURATION
# ============================================================

MAX_HEADER_SEARCH_ROWS = 50

AWB_HEADER_ALIASES = {
    "awb",
    "hawb",
    "awbblno",
    "awbblno.",
    "awbbl",
    "awbhblno",
    "awbhbl",
    "hawbno",
    "hawbnumber",
    "awbno",
    "awbnumber",
}

CW_HEADER_ALIASES = {
    "cw",
    "chargeableweight",
    "chargeablewt",
    "cwt",
}


# ============================================================
# GENERAL TEXT NORMALIZATION
# ============================================================

def normalize_header(value):

    if value is None:
        return ""

    # Avoid problems with lists / Series / arrays
    if not isinstance(value, (str, int, float, bool)):
        try:
            if pd.isna(value):
                return ""
        except Exception:
            return ""

    text = str(value)

    # Normalize common Excel / PDF characters
    text = text.replace("\xa0", " ")
    text = text.replace("\n", " ")
    text = text.replace("\r", " ")
    text = text.replace("\t", " ")

    text = text.strip().upper()

    # Remove spaces and separators
    text = re.sub(r"[\s/_\\\-]+", "", text)

    # Remove ordinal / degree symbols
    text = text.replace("º", "")
    text = text.replace("°", "")

    # Remove punctuation
    text = re.sub(r"[^A-Z0-9]", "", text)

    return text.lower()


# ============================================================
# NORMALIZE HAWB / AWB
# ============================================================

def normalize_hawb(value):

    if value is None:
        return ""

    # IMPORTANT:
    # Prevent pandas "truth value of a Series is ambiguous"
    # error that happened previously.
    if isinstance(value, (pd.Series, pd.DataFrame, list, tuple, dict)):
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    text = str(value)

    text = text.strip().upper()

    # Remove Excel artifacts
    text = text.replace("\xa0", "")
    text = text.replace("\n", "")
    text = text.replace("\r", "")
    text = text.replace("\t", "")

    # Remove spaces and hyphens
    text = text.replace(" ", "")
    text = text.replace("-", "")

    return text


# ============================================================
# IDENTIFY AWB HEADER
# ============================================================

def is_awb_header(value):

    normalized = normalize_header(value)

    if not normalized:
        return False

    if normalized in AWB_HEADER_ALIASES:
        return True

    # Flexible detection
    if "awb" in normalized:
        return True

    if "hawb" in normalized:
        return True

    return False


# ============================================================
# IDENTIFY CW HEADER
# ============================================================

def is_cw_header(value):

    normalized = normalize_header(value)

    if not normalized:
        return False

    if normalized in CW_HEADER_ALIASES:
        return True

    if normalized == "cw":
        return True

    if "chargeableweight" in normalized:
        return True

    return False


# ============================================================
# DEBUG REPRESENTATION OF A ROW
# ============================================================

def row_preview(row):

    values = []

    for value in row:

        if value is None:
            continue

        try:

            if pd.isna(value):
                continue

        except Exception:
            pass

        text = str(value).strip()

        if text:
            values.append(text)

    return " | ".join(values[:20])


# ============================================================
# FIND HEADER
# ============================================================

def find_excel_header(excel_file):

    try:

        excel = pd.ExcelFile(excel_file)

    except Exception as e:

        raise Exception(
            f"Unable to open Excel file. "
            f"Original error: {str(e)}"
        )

    best_candidate = None

    diagnostic_rows = []

    # ========================================================
    # SEARCH EVERY SHEET
    # ========================================================

    for sheet_name in excel.sheet_names:

        try:

            df_preview = pd.read_excel(
                excel_file,
                sheet_name=sheet_name,
                header=None,
                nrows=MAX_HEADER_SEARCH_ROWS
            )

        except Exception as e:

            diagnostic_rows.append(
                f"Sheet '{sheet_name}' could not be read: {str(e)}"
            )

            continue

        for row_index in range(len(df_preview)):

            row = df_preview.iloc[row_index].tolist()

            awb_columns = []
            cw_columns = []

            for col_index, value in enumerate(row):

                if is_awb_header(value):
                    awb_columns.append(col_index)

                if is_cw_header(value):
                    cw_columns.append(col_index)

            # We need both AWB and CW
            if awb_columns and cw_columns:

                candidate = {
                    "sheet": sheet_name,
                    "header_row": row_index,
                    "awb_column": awb_columns[0],
                    "cw_column": cw_columns[0],
                    "preview": row_preview(row),
                }

                # Prefer candidates where AWB and CW
                # are reasonably close.
                distance = abs(
                    awb_columns[0] - cw_columns[0]
                )

                candidate["distance"] = distance

                if best_candidate is None:

                    best_candidate = candidate

                else:

                    # Prefer smaller distance
                    if distance < best_candidate["distance"]:

                        best_candidate = candidate

    # ========================================================
    # NO HEADER FOUND
    # ========================================================

    if best_candidate is None:

        sheet_list = ", ".join(
            excel.sheet_names
        )

        raise Exception(
            "\n"
            "Unable to locate Excel header row.\n\n"
            f"Sheets detected: {sheet_list}\n\n"
            "The program searched the first "
            f"{MAX_HEADER_SEARCH_ROWS} rows of every sheet "
            "for an AWB/HAWB/AWB-BL type header and a CW "
            "or Chargeable Weight header.\n\n"
            "Please verify that the Excel contains both "
            "columns in the same table."
        )

    print("")
    print("========================================")
    print("EXCEL HEADER DETECTED")
    print("========================================")
    print(
        "Sheet:",
        best_candidate["sheet"]
    )
    print(
        "Header row:",
        best_candidate["header_row"] + 1
    )
    print(
        "AWB column:",
        best_candidate["awb_column"] + 1
    )
    print(
        "CW column:",
        best_candidate["cw_column"] + 1
    )
    print(
        "Header:",
        best_candidate["preview"]
    )
    print("========================================")
    print("")

    return (
        best_candidate["sheet"],
        best_candidate["header_row"]
    )


# ============================================================
# READ EXCEL
# ============================================================

def read_excel_file(excel_file):

    sheet_name, header_row = find_excel_header(
        excel_file
    )

    df = pd.read_excel(
        excel_file,
        sheet_name=sheet_name,
        header=header_row
    )

    # ========================================================
    # CLEAN COLUMN NAMES
    # ========================================================

    cleaned_columns = []

    for column in df.columns:

        text = str(column)

        text = text.replace("\xa0", " ")
        text = text.replace("\n", " ")
        text = text.replace("\r", " ")
        text = text.replace("\t", " ")

        text = re.sub(
            r"\s+",
            " ",
            text
        )

        text = text.strip()

        cleaned_columns.append(text)

    df.columns = cleaned_columns

    # ========================================================
    # LOCATE AWB COLUMN
    # ========================================================

    awb_column = None

    for column in df.columns:

        if is_awb_header(column):

            awb_column = column
            break

    # ========================================================
    # LOCATE CW COLUMN
    # ========================================================

    cw_column = None

    for column in df.columns:

        if is_cw_header(column):

            cw_column = column
            break

    # ========================================================
    # ERROR IF AWB NOT FOUND
    # ========================================================

    if awb_column is None:

        available = [
            str(x)
            for x in df.columns
        ]

        raise Exception(
            "AWB column was not found after reading "
            f"the detected header row.\n\n"
            "Columns detected:\n"
            + " | ".join(available)
        )

    # ========================================================
    # ERROR IF CW NOT FOUND
    # ========================================================

    if cw_column is None:

        available = [
            str(x)
            for x in df.columns
        ]

        raise Exception(
            "CW column was not found after reading "
            f"the detected header row.\n\n"
            "Columns detected:\n"
            + " | ".join(available)
        )

    # ========================================================
    # RENAME TO STANDARD NAMES
    # ========================================================

    rename_map = {
        awb_column: "HAWB",
        cw_column: "CW",
    }

    df.rename(
        columns=rename_map,
        inplace=True
    )

    print(
        "AWB column detected as:",
        awb_column
    )

    print(
        "CW column detected as:",
        cw_column
    )

    print(
        "Excel rows:",
        len(df)
    )

    return df, sheet_name, header_row


# ============================================================
# EXTRACT HAWB FROM PDF FILE NAME
# ============================================================

def extract_hawb_from_filename(filename):

    filename_upper = filename.upper()

    # Remove things such as [1], [2], etc.
    filename_upper = re.sub(
        r"\[\d+\]",
        "",
        filename_upper
    )

    # --------------------------------------------------------
    # Common HAWB format
    # PTY0045653
    # SPTY0045648
    # --------------------------------------------------------

    match = re.search(
        r"\b[A-Z]{2,5}\d{5,}\b",
        filename_upper
    )

    if match:

        return normalize_hawb(
            match.group(0)
        )

    # --------------------------------------------------------
    # AWB format
    # 202-31647291
    # 992-10748500
    # --------------------------------------------------------

    match = re.search(
        r"\b\d{3}-\d{5,}\b",
        filename_upper
    )

    if match:

        return normalize_hawb(
            match.group(0)
        )

    # --------------------------------------------------------
    # AWB without hyphen
    # --------------------------------------------------------

    match = re.search(
        r"\b\d{8,12}\b",
        filename_upper
    )

    if match:

        return normalize_hawb(
            match.group(0)
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

                try:

                    text = page.extract_text()

                except Exception:

                    text = None

                if text:

                    pdf_text += text + "\n"

    except Exception as e:

        print(
            "Unable to read PDF:",
            pdf_path,
            str(e)
        )

        return None

    lines = pdf_text.splitlines()

    # ========================================================
    # ORIGINAL PROVIDER
    # ========================================================

    for line in lines:

        match = re.search(
            r"\d+(?:\.\d+)?\s+"
            r"\d+(?:\.\d+)?K?\s+"
            r"[A-Z]\s+"
            r"(\d+(?:\.\d+)?)\s+"
            r"\d+(?:\.\d+)?",
            line.upper()
        )

        if match:

            try:

                return float(
                    match.group(1)
                )

            except Exception:

                pass

    # ========================================================
    # CHARGEABLE WEIGHT LABEL
    # ========================================================

    for line in lines:

        upper_line = line.upper()

        if (
            "CHARGEABLE WEIGHT" in upper_line
            or "CHARGEABLE WT" in upper_line
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

    # ========================================================
    # FALLBACK:
    # SEARCH FOR "CW"
    # ========================================================

    for line in lines:

        upper_line = line.upper()

        if re.search(
            r"\bCW\b",
            upper_line
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
# BUILD PDF INDEX
# ============================================================

def build_pdf_index(input_folder):

    pdf_index = {}

    pdf_files = []

    for file in os.listdir(input_folder):

        if file.lower().endswith(".pdf"):

            pdf_files.append(file)

    print(
        "PDF files detected:",
        len(pdf_files)
    )

    for file in pdf_files:

        hawb = extract_hawb_from_filename(
            file
        )

        if hawb:

            pdf_index[hawb] = file

            print(
                "PDF indexed:",
                hawb,
                "->",
                file
            )

        else:

            print(
                "WARNING: Could not extract AWB from:",
                file
            )

    return pdf_index


# ============================================================
# VALIDATE AWB
# ============================================================

def validate_awb(
    input_folder,
    output_folder
):

    print("")
    print("========================================")
    print("STARTING AWB VALIDATION")
    print("========================================")

    # ========================================================
    # FIND EXCEL
    # ========================================================

    excel_file = None

    for file in os.listdir(input_folder):

        if file.lower().endswith(
            (".xlsx", ".xls")
        ):

            # Do not accidentally use the generated result
            if "AWB_Validation_Result" in file:

                continue

            excel_file = os.path.join(
                input_folder,
                file
            )

            break

    if excel_file is None:

        raise Exception(
            "No Excel file found in the uploaded files."
        )

    print(
        "Excel file:",
        excel_file
    )

    # ========================================================
    # READ EXCEL
    # ========================================================

    df_excel, sheet_name, header_row = read_excel_file(
        excel_file
    )

    print(
        "Using sheet:",
        sheet_name
    )

    print(
        "Using header row:",
        header_row + 1
    )

    # ========================================================
    # NORMALIZE HAWB
    # ========================================================

    df_excel["HAWB"] = (
        df_excel["HAWB"]
        .map(normalize_hawb)
    )

    # ========================================================
    # NORMALIZE CW
    # ========================================================

    # Handles values such as:
    # 651.50
    # "651.50"
    # "651,50"
    # "651.50 KG"

    def normalize_cw(value):

        if value is None:
            return None

        if isinstance(
            value,
            (pd.Series, pd.DataFrame, list, tuple, dict)
        ):
            return None

        try:

            if pd.isna(value):
                return None

        except Exception:

            pass

        text = str(value).strip()

        if not text:
            return None

        text = text.replace(
            ",",
            ""
        )

        match = re.search(
            r"-?\d+(?:\.\d+)?",
            text
        )

        if not match:
            return None

        try:

            return float(
                match.group(0)
            )

        except Exception:

            return None

    df_excel["CW"] = (
        df_excel["CW"]
        .map(normalize_cw)
    )

    # ========================================================
    # REMOVE EMPTY HAWB ROWS
    # ========================================================

    df_excel = df_excel[
        df_excel["HAWB"] != ""
    ].copy()

    print(
        "Valid Excel rows:",
        len(df_excel)
    )

    # ========================================================
    # CREATE PDF INDEX
    # ========================================================

    pdf_index = build_pdf_index(
        input_folder
    )

    print(
        "Indexed PDFs:",
        len(pdf_index)
    )

    # ========================================================
    # VALIDATE
    # ========================================================

    results = []

    for _, row in df_excel.iterrows():

        excel_hawb = row["HAWB"]
        excel_cw = row["CW"]

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

        pdf_file = pdf_index[
            excel_hawb
        ]

        pdf_path = os.path.join(
            input_folder,
            pdf_file
        )

        # ----------------------------------------------------
        # EXTRACT PDF CW
        # ----------------------------------------------------

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

        # ----------------------------------------------------
        # EXCEL CW MISSING
        # ----------------------------------------------------

        if excel_cw is None:

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
        # COMPARE
        # ----------------------------------------------------

        difference = round(
            abs(
                float(excel_cw)
                - float(pdf_cw)
            ),
            2
        )

        if difference <= 0.01:

            validation_result = "PASS"

        else:

            validation_result = "FAIL"

        results.append({

            "HAWB": excel_hawb,
            "Excel CW": excel_cw,
            "PDF CW": pdf_cw,
            "Difference": difference,
            "Result": validation_result,
            "PDF File": pdf_file

        })

    # ========================================================
    # PDFS NOT FOUND IN EXCEL
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
    print("========================================")
    print("VALIDATION COMPLETED")
    print("========================================")

    print(
        "Total results:",
        len(df_results)
    )

    if not df_results.empty:

        print(
            "PASS:",
            (
                df_results["Result"]
                == "PASS"
            ).sum()
        )

        print(
            "FAIL:",
            (
                df_results["Result"]
                == "FAIL"
            ).sum()
        )

        print(
            "PDF NOT FOUND:",
            (
                df_results["Result"]
                == "PDF NOT FOUND"
            ).sum()
        )

        print(
            "CW NOT FOUND IN PDF:",
            (
                df_results["Result"]
                == "CW NOT FOUND IN PDF"
            ).sum()
        )

    print(
        "Output:",
        output_file
    )

    print("========================================")

    return output_file