import pandas as pd
import pdfplumber
import os
import re


# =====================================================
# NORMALIZE VALUE
# =====================================================

def normalize_text(value):

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    return (
        str(value)
        .strip()
        .upper()
    )


# =====================================================
# NORMALIZE AWB / HAWB
# =====================================================

def normalize_hawb(value):

    value = normalize_text(value)

    if not value:
        return ""

    # Remove spaces, hyphens and common separators
    value = re.sub(
        r"[\s\-_\/\\]+",
        "",
        value
    )

    # Remove non-printable characters
    value = "".join(
        char for char in value
        if char.isprintable()
    )

    return value


# =====================================================
# NORMALIZE PDF FILENAME FOR MATCHING
# =====================================================

def normalize_filename_for_matching(filename):

    filename = normalize_text(filename)

    # Remove extension
    filename = re.sub(
        r"\.PDF$",
        "",
        filename,
        flags=re.IGNORECASE
    )

    # Normalize exactly the same way as AWB
    filename = re.sub(
        r"[\s\-_\/\\]+",
        "",
        filename
    )

    return filename


# =====================================================
# FIND EXCEL HEADER ROW
# =====================================================

def find_excel_header(excel_file):

    # Read more rows because some supplier files
    # contain title rows before the real header.
    df = pd.read_excel(
        excel_file,
        header=None,
        nrows=30
    )

    for i in range(len(df)):

        raw_row = df.iloc[i].tolist()

        row = []

        for value in raw_row:

            value = normalize_text(value)

            # Normalize header spacing / separators
            value = re.sub(
                r"[\s\-_\/]+",
                " ",
                value
            )

            value = value.strip()

            row.append(value)

        # ---------------------------------------------
        # Detect AWB-type header
        # ---------------------------------------------

        awb_found = False
        cw_found = False

        for cell in row:

            cell_upper = cell.upper()

            # Possible AWB headers
            if (
                cell_upper == "AWB"
                or cell_upper == "HAWB"
                or "AWB BL" in cell_upper
                or "AWB/ BL" in cell_upper
                or "AWB HBL" in cell_upper
                or "HAWB NO" in cell_upper
                or "AWB NO" in cell_upper
            ):

                awb_found = True

            # CW header
            if (
                cell_upper == "CW"
                or cell_upper.startswith("CW ")
                or "CHARGEABLE WEIGHT" in cell_upper
            ):

                cw_found = True

        if awb_found and cw_found:

            return i

    raise Exception(
        "Unable to locate Excel header row. "
        "The file must contain an AWB/HAWB column and a CW column."
    )


# =====================================================
# FIND COLUMN
# =====================================================

def find_column(columns, column_type):

    normalized_columns = []

    for column in columns:

        original = str(column)

        normalized = normalize_text(
            column
        )

        normalized = re.sub(
            r"[\s\-_\/]+",
            " ",
            normalized
        )

        normalized = normalized.strip()

        normalized_columns.append(
            (original, normalized)
        )

    # =================================================
    # AWB COLUMN
    # =================================================

    if column_type == "AWB":

        # Priority 1 - exact AWB
        for original, normalized in normalized_columns:

            if normalized == "AWB":

                return original

        # Priority 2 - exact HAWB
        for original, normalized in normalized_columns:

            if normalized == "HAWB":

                return original

        # Priority 3 - AWB / BL
        for original, normalized in normalized_columns:

            if (
                "AWB BL" in normalized
                or "AWB/ BL" in normalized
                or "AWB HBL" in normalized
            ):

                return original

        # Priority 4 - AWB No
        for original, normalized in normalized_columns:

            if (
                normalized.startswith("AWB NO")
                or normalized.startswith("HAWB NO")
            ):

                return original

        return None

    # =================================================
    # CW COLUMN
    # =================================================

    if column_type == "CW":

        # Exact CW first
        for original, normalized in normalized_columns:

            if normalized == "CW":

                return original

        # Chargeable Weight
        for original, normalized in normalized_columns:

            if "CHARGEABLE WEIGHT" in normalized:

                return original

        return None

    return None


# =====================================================
# READ EXCEL
# =====================================================

def read_excel_file(excel_file):

    header_row = find_excel_header(
        excel_file
    )

    df = pd.read_excel(
        excel_file,
        header=header_row
    )

    # Clean column names
    cleaned_columns = []

    for column in df.columns:

        column_name = normalize_text(
            column
        )

        # Replace non-breaking spaces
        column_name = column_name.replace(
            "\xa0",
            " "
        )

        # Collapse spaces
        column_name = re.sub(
            r"\s+",
            " ",
            column_name
        ).strip()

        cleaned_columns.append(
            column_name
        )

    df.columns = cleaned_columns

    # Find AWB
    awb_column = find_column(
        df.columns,
        "AWB"
    )

    if awb_column is None:

        raise Exception(
            "AWB/HAWB column not found in Excel."
        )

    # Find CW
    cw_column = find_column(
        df.columns,
        "CW"
    )

    if cw_column is None:

        raise Exception(
            "CW column not found in Excel."
        )

    # Rename dynamically detected columns
    rename_map = {
        awb_column: "HAWB",
        cw_column: "CW"
    }

    df.rename(
        columns=rename_map,
        inplace=True
    )

    # Normalize AWB
    df["HAWB"] = df["HAWB"].apply(
        normalize_hawb
    )

    # Convert CW
    df["CW"] = pd.to_numeric(
        df["CW"],
        errors="coerce"
    )

    # Remove completely empty AWBs
    df = df[
        df["HAWB"] != ""
    ].copy()

    return df, header_row


# =====================================================
# FIND PDF FOR AWB
# =====================================================

def find_pdf_for_hawb(
    hawb,
    pdf_files
):

    normalized_hawb = normalize_hawb(
        hawb
    )

    if not normalized_hawb:

        return None

    # ---------------------------------------------
    # First pass:
    # Exact normalized AWB contained in filename
    # ---------------------------------------------

    for pdf_file in pdf_files:

        normalized_filename = (
            normalize_filename_for_matching(
                pdf_file
            )
        )

        if normalized_hawb in normalized_filename:

            return pdf_file

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

    except Exception:

        return None

    if not pdf_text:

        return None

    lines = pdf_text.split("\n")

    # =================================================
    # ORIGINAL PROVIDER
    # =================================================

    for line in lines:

        match = re.search(
            r'\d+\s+\d+(?:\.\d+)?K\s+[A-Z]\s+'
            r'(\d+(?:\.\d+)?)\s+'
            r'\d+(?:\.\d+)?',
            line
        )

        if match:

            try:

                return float(
                    match.group(1)
                )

            except (
                ValueError,
                TypeError
            ):

                pass

    # =================================================
    # PROVIDER WITH "CHARGEABLE WEIGHT"
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

                except (
                    ValueError,
                    TypeError
                ):

                    pass

    # =================================================
    # ADDITIONAL CW PATTERNS
    # =================================================

    for line in lines:

        upper_line = line.upper()

        # Common labels
        if (
            "C.W." in upper_line
            or "CHG WT" in upper_line
            or "CHARGE WT" in upper_line
            or "CHARGEABLE WT" in upper_line
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

                except (
                    ValueError,
                    TypeError
                ):

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

        if file.lower().endswith(
            ".xlsx"
        ):

            # Ignore generated result file
            if file.startswith(
                "AWB_Validation_Result"
            ):

                continue

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
    # READ EXCEL
    # =================================================

    df_excel, header_row = read_excel_file(
        excel_file
    )

    print(
        f"Header detected on row: {header_row}"
    )

    print(
        f"Excel rows detected: {len(df_excel)}"
    )

    print(
        f"AWB values detected: "
        f"{df_excel['HAWB'].head(10).tolist()}"
    )

    # =================================================
    # FIND ALL PDF FILES
    # =================================================

    pdf_files = []

    for file in os.listdir(input_folder):

        if file.lower().endswith(
            ".pdf"
        ):

            pdf_files.append(
                file
            )

    print(
        f"PDF files found: {len(pdf_files)}"
    )

    # =================================================
    # VALIDATE
    # =================================================

    results = []

    matched_pdf_files = set()

    for _, row in df_excel.iterrows():

        excel_hawb = row["HAWB"]
        excel_cw = row["CW"]

        # ---------------------------------------------
        # Find PDF by direct AWB search in filename
        # ---------------------------------------------

        pdf_file = find_pdf_for_hawb(
            excel_hawb,
            pdf_files
        )

        # ---------------------------------------------
        # PDF NOT FOUND
        # ---------------------------------------------

        if pdf_file is None:

            results.append({

                "HAWB": excel_hawb,
                "Excel CW": excel_cw,
                "PDF CW": "",
                "Difference": "",
                "Result": "PDF NOT FOUND",
                "PDF File": ""

            })

            continue

        matched_pdf_files.add(
            pdf_file
        )

        pdf_path = os.path.join(
            input_folder,
            pdf_file
        )

        # ---------------------------------------------
        # Extract PDF CW
        # ---------------------------------------------

        pdf_cw = extract_pdf_cw(
            pdf_path
        )

        # ---------------------------------------------
        # CW NOT FOUND
        # ---------------------------------------------

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

        # ---------------------------------------------
        # Compare CW
        # ---------------------------------------------

        try:

            difference = round(
                abs(
                    float(excel_cw) -
                    float(pdf_cw)
                ),
                2
            )

        except (
            ValueError,
            TypeError
        ):

            results.append({

                "HAWB": excel_hawb,
                "Excel CW": excel_cw,
                "PDF CW": pdf_cw,
                "Difference": "",
                "Result": "INVALID EXCEL CW",
                "PDF File": pdf_file

            })

            continue

        # ---------------------------------------------
        # Result
        # ---------------------------------------------

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
    # EXTRA PDFS
    # =================================================

    excel_hawb_set = set(
        df_excel["HAWB"]
    )

    for pdf_file in pdf_files:

        # Find whether this PDF contains
        # an AWB that exists in Excel.
        normalized_filename = (
            normalize_filename_for_matching(
                pdf_file
            )
        )

        found_in_excel = False

        for excel_hawb in excel_hawb_set:

            if (
                excel_hawb
                and
                excel_hawb in normalized_filename
            ):

                found_in_excel = True
                break

        if not found_in_excel:

            results.append({

                "HAWB": "",
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

    # =================================================
    # SUMMARY
    # =================================================

    total = len(results)

    pass_count = sum(
        1
        for result in results
        if result["Result"] == "PASS"
    )

    fail_count = sum(
        1
        for result in results
        if result["Result"] == "FAIL"
    )

    pdf_not_found_count = sum(
        1
        for result in results
        if result["Result"] == "PDF NOT FOUND"
    )

    cw_not_found_count = sum(
        1
        for result in results
        if result["Result"] == "CW NOT FOUND IN PDF"
    )

    print("")
    print(
        "=============================="
    )
    print(
        "VALIDATION SUMMARY"
    )
    print(
        "=============================="
    )
    print(
        f"Total results: {total}"
    )
    print(
        f"PASS: {pass_count}"
    )
    print(
        f"FAIL: {fail_count}"
    )
    print(
        f"PDF NOT FOUND: "
        f"{pdf_not_found_count}"
    )
    print(
        f"CW NOT FOUND IN PDF: "
        f"{cw_not_found_count}"
    )
    print(
        "=============================="
    )

    print(
        "Validation completed:",
        output_file
    )

    return output_file