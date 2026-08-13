import os
import re
import glob
import pandas as pd
import pdfplumber
from openpyxl import load_workbook


# ============================================================
# CONFIGURATION
# ============================================================

TOLERANCE = 0.05


# ============================================================
# NORMALIZATION
# ============================================================

def normalize_text(value):
    """
    Generic text normalization.
    """

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    text = str(value)

    text = (
        text
        .replace("\xa0", " ")
        .replace("\n", " ")
        .replace("\r", " ")
        .replace("\t", " ")
    )

    text = re.sub(r"\s+", " ", text)

    return text.strip().upper()


def normalize_awb(value):
    """
    Normalize AWB / HAWB.

    Examples:

    J560881
    J-560881
    J 560881

    -> J560881

    PTY0045653
    PTY-0045653

    -> PTY0045653

    810-42903125

    -> 81042903125
    """

    text = normalize_text(value)

    if not text:
        return ""

    text = re.sub(r"[^A-Z0-9]", "", text)

    return text


def normalize_number(value):
    """
    Convert Excel numeric values into float.
    """

    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()

    if not text:
        return None

    text = (
        text
        .replace(",", "")
        .replace("\xa0", "")
        .strip()
    )

    try:
        return float(text)
    except ValueError:
        return None


# ============================================================
# HEADER NORMALIZATION
# ============================================================

def normalize_header(value):
    """
    Normalize an Excel header so different variations
    can be recognized.

    Examples:

    AWB
    AWB/ BL Nº
    AWB / BL No
    HAWB
    CW
    Chargeable Weight
    """

    text = normalize_text(value)

    if not text:
        return ""

    text = (
        text
        .replace("№", "NO")
        .replace("º", "O")
    )

    # Remove punctuation but preserve letters/numbers.
    text = re.sub(r"[^A-Z0-9]", "", text)

    return text


# ============================================================
# HEADER IDENTIFICATION
# ============================================================

def is_awb_header(value):
    """
    Determines whether a column is an AWB/HAWB column.

    IMPORTANT:
    MAWB is intentionally NOT accepted.
    """

    h = normalize_header(value)

    if not h:
        return False

    # Explicitly reject MAWB.
    if h == "MAWB":
        return False

    if "MAWB" in h:
        return False

    accepted = {
        "AWB",
        "HAWB",
        "AWBBLNO",
        "AWBBLNO",
        "AWBBL",
        "AWBNUMBER",
        "HAWBNUMBER",
        "HAWBNO",
        "AWBNO",
    }

    if h in accepted:
        return True

    # Variations such as:
    # AWB/BL Nº
    # AWB / BL No
    # HAWB No.
    if h.startswith("AWB") and "MAWB" not in h:
        return True

    if h.startswith("HAWB"):
        return True

    return False


def is_cw_header(value):
    """
    Determines whether a column represents Chargeable Weight.
    """

    h = normalize_header(value)

    if not h:
        return False

    accepted = {
        "CW",
        "CWT",
        "CHARGEABLEWEIGHT",
        "CHARGEABLEWT",
        "CHARGEWEIGHT",
        "WEIGHTCW",
    }

    if h in accepted:
        return True

    # Avoid confusing GW with CW.
    if h.startswith("CW") and "CARRIER" not in h:
        return True

    if "CHARGEABLE" in h and "WEIGHT" in h:
        return True

    return False


def find_excel_header(excel_file):
    """
    Search the workbook for the real header row.

    The header does NOT need to be on row 1 or row 2.

    The algorithm looks for a row containing:

        AWB / HAWB
        +
        CW

    MAWB is NOT accepted as the AWB field.
    """

    try:
        raw = pd.read_excel(
            excel_file,
            header=None,
            engine="openpyxl"
        )
    except Exception as e:
        raise Exception(
            f"Unable to read Excel file: {str(e)}"
        )

    best_row = None
    best_score = -1

    max_rows = min(len(raw), 100)

    for row_idx in range(max_rows):

        row = raw.iloc[row_idx]

        awb_found = False
        cw_found = False

        for value in row:

            if is_awb_header(value):
                awb_found = True

            if is_cw_header(value):
                cw_found = True

        score = 0

        if awb_found:
            score += 10

        if cw_found:
            score += 10

        # Strong preference for rows containing both.
        if awb_found and cw_found:
            score += 50

        if score > best_score:

            best_score = score
            best_row = row_idx

    if best_row is None or best_score < 20:

        # Create diagnostic information.
        diagnostic = []

        for row_idx in range(min(len(raw), 20)):

            values = [
                normalize_text(v)
                for v in raw.iloc[row_idx].tolist()
                if normalize_text(v)
            ]

            if values:
                diagnostic.append(
                    f"Row {row_idx + 1}: {values[:15]}"
                )

        raise Exception(
            "Unable to locate Excel header row.\n\n"
            "The Excel must contain an AWB/HAWB column and a CW column.\n\n"
            + "\n".join(diagnostic)
        )

    return best_row


# ============================================================
# READ EXCEL
# ============================================================

def read_excel_file(excel_file):

    header_row = find_excel_header(excel_file)

    df = pd.read_excel(
        excel_file,
        header=header_row,
        engine="openpyxl"
    )

    # Normalize column names.
    new_columns = []

    for column in df.columns:

        text = normalize_text(column)

        new_columns.append(text)

    df.columns = new_columns

    # --------------------------------------------------------
    # FIND AWB COLUMN
    # --------------------------------------------------------

    awb_column = None

    for column in df.columns:

        if is_awb_header(column):

            awb_column = column
            break

    if awb_column is None:

        raise Exception(
            "AWB/HAWB column could not be identified."
        )

    # --------------------------------------------------------
    # FIND CW COLUMN
    # --------------------------------------------------------

    cw_column = None

    for column in df.columns:

        if is_cw_header(column):

            cw_column = column
            break

    if cw_column is None:

        raise Exception(
            "CW column could not be identified."
        )

    return df, header_row, awb_column, cw_column


# ============================================================
# PDF FILE INDEX
# ============================================================

def create_pdf_index(input_folder):

    pdf_files = glob.glob(
        os.path.join(
            input_folder,
            "*.pdf"
        )
    )

    index = {}

    for pdf_path in pdf_files:

        filename = os.path.basename(pdf_path)

        stem = os.path.splitext(filename)[0]

        normalized_filename = normalize_awb(stem)

        if normalized_filename:

            index.setdefault(
                normalized_filename,
                []
            ).append(pdf_path)

    return pdf_files, index


# ============================================================
# FIND PDF
# ============================================================

def find_pdf_for_awb(
    awb,
    pdf_files,
    pdf_index
):

    normalized_awb = normalize_awb(awb)

    if not normalized_awb:
        return None

    # --------------------------------------------------------
    # 1. Exact normalized filename match
    # --------------------------------------------------------

    if normalized_awb in pdf_index:

        candidates = pdf_index[normalized_awb]

        if candidates:
            return candidates[0]

    # --------------------------------------------------------
    # 2. Search AWB inside filename
    #
    # This supports filenames such as:
    #
    # Original 2 - (for Consignee) - J560881
    #
    # Copy 5 - (Extra Copy) - HAWB No_ J560885
    #
    # Original - PTY0045653
    #
    # --------------------------------------------------------

    for pdf_path in pdf_files:

        filename = os.path.basename(pdf_path)

        normalized_filename = normalize_awb(
            os.path.splitext(filename)[0]
        )

        if normalized_awb in normalized_filename:

            return pdf_path

    return None


# ============================================================
# EXTRACT TEXT FROM PDF
# ============================================================

def extract_pdf_text(pdf_path):

    text_parts = []

    try:

        with pdfplumber.open(pdf_path) as pdf:

            for page in pdf.pages:

                text = page.extract_text()

                if text:
                    text_parts.append(text)

    except Exception:

        return ""

    return "\n".join(text_parts)


# ============================================================
# FIND NUMBER NEAR CW
# ============================================================

def extract_cw_from_text(text):

    if not text:
        return None

    # Normalize text.
    clean = (
        text
        .replace("\xa0", " ")
        .replace(",", "")
    )

    # --------------------------------------------------------
    # Patterns for Chargeable Weight
    # --------------------------------------------------------

    patterns = [

        # Chargeable Weight: 123.45
        r"CHARGEABLE\s+WEIGHT\s*[:\-]?\s*(\d+(?:\.\d+)?)",

        # Chargeable Wt: 123.45
        r"CHARGEABLE\s+WT\.?\s*[:\-]?\s*(\d+(?:\.\d+)?)",

        # C.W.: 123.45
        r"\bC\.?\s*W\.?\s*[:\-]?\s*(\d+(?:\.\d+)?)",

        # CW: 123.45
        r"\bCW\s*[:\-]?\s*(\d+(?:\.\d+)?)",

        # CWT: 123.45
        r"\bCWT\s*[:\-]?\s*(\d+(?:\.\d+)?)",

        # Chargeable Weight 123.45 KG
        r"CHARGEABLE\s+WEIGHT\s+(\d+(?:\.\d+)?)\s*(?:KG)?",

    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            clean,
            flags=re.IGNORECASE
        )

        if match:

            try:
                return float(match.group(1))
            except ValueError:
                pass

    # --------------------------------------------------------
    # Second method:
    # Search line-by-line
    # --------------------------------------------------------

    lines = clean.splitlines()

    for i, line in enumerate(lines):

        normalized_line = normalize_text(line)

        if (
            "CHARGEABLE WEIGHT" in normalized_line
            or normalized_line.startswith("CW")
            or normalized_line.startswith("CWT")
        ):

            numbers = re.findall(
                r"\d+(?:\.\d+)?",
                line
            )

            if numbers:

                try:
                    return float(numbers[-1])
                except ValueError:
                    pass

    return None


# ============================================================
# FALLBACK PDF SEARCH
# ============================================================

def search_awb_inside_pdf(
    awb,
    pdf_files
):

    normalized_awb = normalize_awb(awb)

    if not normalized_awb:
        return None, None

    for pdf_path in pdf_files:

        text = extract_pdf_text(pdf_path)

        if not text:
            continue

        normalized_text = normalize_awb(text)

        if normalized_awb in normalized_text:

            return pdf_path, text

    return None, None


# ============================================================
# VALIDATION
# ============================================================

def validate_awb(
    input_folder,
    output_folder
):

    # --------------------------------------------------------
    # CREATE OUTPUT FOLDER
    # --------------------------------------------------------

    os.makedirs(
        output_folder,
        exist_ok=True
    )

    # --------------------------------------------------------
    # FIND EXCEL
    # --------------------------------------------------------

    excel_files = glob.glob(
        os.path.join(
            input_folder,
            "*.xlsx"
        )
    )

    if not excel_files:

        raise Exception(
            "No Excel file was found."
        )

    # Ignore previous output files if any.
    excel_files = [
        f for f in excel_files
        if "validation_result" not in os.path.basename(f).lower()
    ]

    if not excel_files:

        raise Exception(
            "No valid Excel input file was found."
        )

    excel_file = excel_files[0]

    # --------------------------------------------------------
    # READ EXCEL
    # --------------------------------------------------------

    df_excel, header_row, awb_column, cw_column = read_excel_file(
        excel_file
    )

    # --------------------------------------------------------
    # PDF INDEX
    # --------------------------------------------------------

    pdf_files, pdf_index = create_pdf_index(
        input_folder
    )

    # --------------------------------------------------------
    # VALIDATION RESULTS
    # --------------------------------------------------------

    results = []

    # --------------------------------------------------------
    # PROCESS EACH EXCEL ROW
    # --------------------------------------------------------

    for _, row in df_excel.iterrows():

        awb_value = row.get(
            awb_column
        )

        cw_value = row.get(
            cw_column
        )

        awb = normalize_awb(
            awb_value
        )

        excel_cw = normalize_number(
            cw_value
        )

        # Skip completely empty rows.
        if not awb:

            continue

        pdf_path = find_pdf_for_awb(
            awb,
            pdf_files,
            pdf_index
        )

        pdf_cw = None
        difference = None
        result = "PDF NOT FOUND"

        # ----------------------------------------------------
        # PDF FOUND
        # ----------------------------------------------------

        if pdf_path:

            pdf_text = extract_pdf_text(
                pdf_path
            )

            pdf_cw = extract_cw_from_text(
                pdf_text
            )

            # ------------------------------------------------
            # If CW wasn't found, try searching the PDF text
            # again using the AWB.
            # ------------------------------------------------

            if pdf_cw is None:

                pdf_cw = extract_cw_from_text(
                    pdf_text
                )

            # ------------------------------------------------
            # Compare
            # ------------------------------------------------

            if pdf_cw is not None and excel_cw is not None:

                difference = round(
                    pdf_cw - excel_cw,
                    2
                )

                if abs(difference) <= TOLERANCE:

                    result = "PASS"

                else:

                    result = "FAIL"

            elif pdf_cw is None:

                result = "CW NOT FOUND IN PDF"

            else:

                result = "EXCEL CW NOT FOUND"

        # ----------------------------------------------------
        # OUTPUT ROW
        # ----------------------------------------------------

        results.append(
            {
                "AWB": awb,
                "Excel CW": excel_cw,
                "PDF CW": pdf_cw,
                "Difference": difference,
                "Result": result,
                "PDF File": (
                    os.path.basename(pdf_path)
                    if pdf_path
                    else ""
                )
            }
        )

    # --------------------------------------------------------
    # CREATE RESULT DATAFRAME
    # --------------------------------------------------------

    result_df = pd.DataFrame(
        results
    )

    # --------------------------------------------------------
    # OUTPUT FILE
    # --------------------------------------------------------

    output_file = os.path.join(
        output_folder,
        "AWB_Validation_Result.xlsx"
    )

    # --------------------------------------------------------
    # WRITE EXCEL
    # --------------------------------------------------------

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        result_df.to_excel(
            writer,
            index=False,
            sheet_name="Validation"
        )

    return output_file