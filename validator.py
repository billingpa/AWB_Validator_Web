import pandas as pd
import pdfplumber
import os
import re


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
# FIND EXCEL HEADER ROW
# =====================================================

def find_excel_header(excel_file):

    df = pd.read_excel(
        excel_file,
        header=None,
        nrows=10
    )

    for i in range(len(df)):

        row = [
            str(x).strip()
            for x in df.iloc[i].tolist()
        ]

        if (
            "HAWB" in row
            and "CW" in row
        ):

            return i

        if (
            "AWB/ BL Nº" in row
            and "CW" in row
        ):

            return i

    raise Exception(
        "Unable to locate Excel header row."
    )


# =====================================================
# EXTRACT HAWB FROM PDF FILE NAME
# =====================================================

def extract_hawb_from_filename(filename):

    filename = re.sub(
        r'\[\d+\]',
        '',
        filename
    )

    match = re.search(
        r'([A-Z]{3}\d{7}|[A-Z]\d{5,}|[0-9]{3}-?[0-9]{5,})',
        filename.upper()
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

    except Exception:

        return None

    lines = pdf_text.split("\n")

    # ORIGINAL PROVIDER

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

            except:

                pass

    # SVC PROVIDER

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

                except:

                    pass

    return None
def validate_awb(
    input_folder,
    output_folder
):

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

    print(f"Excel found: {excel_file}")

    # DETECT HEADER

    header_row = find_excel_header(
        excel_file
    )

    print(
        f"Header detected on row: {header_row}"
    )

    # READ EXCEL

    df_excel = pd.read_excel(
        excel_file,
        header=header_row
    )

    if "AWB/ BL Nº" in df_excel.columns:

        df_excel.rename(
            columns={
                "AWB/ BL Nº": "HAWB"
            },
            inplace=True
        )

        print(
            "Provider detected: SVC"
        )

    else:

        print(
            "Provider detected: ORIGINAL"
        )

    if "HAWB" not in df_excel.columns:

        raise Exception(
            "HAWB column not found."
        )

    if "CW" not in df_excel.columns:

        raise Exception(
            "CW column not found."
        )

    df_excel["HAWB"] = df_excel["HAWB"].apply(
        normalize_hawb
    )

    df_excel["CW"] = pd.to_numeric(
        df_excel["CW"],
        errors="coerce"
    )

    # CREATE PDF INDEX

    pdf_index = {}

    for file in os.listdir(input_folder):

        if file.lower().endswith(".pdf"):

            hawb = extract_hawb_from_filename(
                file
            )

            if hawb:

                pdf_index[hawb] = file

    print(
        f"PDFs found: {len(pdf_index)}"
    )

    results = []

    # VALIDATE EXCEL ROWS

    for _, row in df_excel.iterrows():

        excel_hawb = row["HAWB"]
        excel_cw = row["CW"]

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

        pdf_file = pdf_index[excel_hawb]

        pdf_path = os.path.join(
            input_folder,
            pdf_file
        )

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

        difference = round(
            abs(
                float(excel_cw) - pdf_cw
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

    # EXTRA PDFS

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

    # SAVE RESULT

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