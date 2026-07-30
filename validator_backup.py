def validate_awb(
        input_folder,
        output_folder
    ):
    import pandas as pd
    import pdfplumber
    import os
    import re

    INPUT_FOLDER = input_folder
    OUTPUT_FOLDER = output_folder


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
    # FIND EXCEL FILE
    # =====================================================

    excel_file = None

    for file in os.listdir(INPUT_FOLDER):

        if file.lower().endswith(".xlsx"):

            excel_file = os.path.join(INPUT_FOLDER, file)
            break


    if not excel_file:

        raise Exception(
            "No Excel file found inside input folder."
        )


    print(f"Excel found: {excel_file}")


    # =====================================================
    # READ EXCEL
    # =====================================================

    df_excel = pd.read_excel(
        excel_file,
        header=1
    )


    df_excel["HAWB"] = df_excel["HAWB"].apply(
        normalize_hawb
    )


    df_excel["CW"] = pd.to_numeric(
        df_excel["CW"],
        errors="coerce"
    )


    # =====================================================
    # PROCESS PDFS
    # =====================================================

    results = []


    pdf_files = [
        f for f in os.listdir(INPUT_FOLDER)
        if f.lower().endswith(".pdf")
    ]


    for pdf_file in pdf_files:


        print(f"Reading: {pdf_file}")


        # ==========================================
        # GET HAWB FROM FILE NAME
        # ==========================================

        hawb_from_filename = None


        # Accept:
        # J123456
        # A12345678
        # 810-42902856
        # 81042902856

        match = re.search(
            r'([A-Z]\d{5,}|[0-9]{3}-?[0-9]{5,})',
            pdf_file.upper()
        )


        if match:

            hawb_from_filename = normalize_hawb(
                match.group(1)
            )


        if not hawb_from_filename:


            results.append({

                "HAWB": "",
                "Excel CW": "",
                "PDF CW": "",
                "Difference": "",
                "Result": "HAWB NOT FOUND IN FILE NAME",
                "PDF File": pdf_file

            })


            continue



        pdf_path = os.path.join(
            INPUT_FOLDER,
            pdf_file
        )


        pdf_text = ""


        # ==========================================
        # READ PDF
        # ==========================================

        try:


            with pdfplumber.open(pdf_path) as pdf:


                for page in pdf.pages:


                    text = page.extract_text()


                    if text:

                        pdf_text += text + "\n"



        except Exception as e:


            results.append({

                "HAWB": hawb_from_filename,
                "Excel CW": "",
                "PDF CW": "",
                "Difference": "",
                "Result": f"PDF ERROR: {str(e)}",
                "PDF File": pdf_file

            })


            continue



        # ==========================================
        # EXTRACT CHARGEABLE WEIGHT
        # ==========================================

        pdf_cw = None


        lines = pdf_text.split("\n")


        for line in lines:


            match = re.search(

                r'\d+\s+\d+(?:\.\d+)?K\s+[A-Z]\s+(\d+(?:\.\d+)?)\s+\d+(?:\.\d+)?',

                line

            )


            if match:


                try:

                    pdf_cw = float(
                        match.group(1)
                    )

                    break


                except:

                    pass



        # ==========================================
        # FIND HAWB IN EXCEL
        # ==========================================

        match_excel = df_excel[
            df_excel["HAWB"] == hawb_from_filename
        ]



        if match_excel.empty:


            results.append({

                "HAWB": hawb_from_filename,
                "Excel CW": "",
                "PDF CW": pdf_cw,
                "Difference": "",
                "Result": "HAWB NOT FOUND IN EXCEL",
                "PDF File": pdf_file

            })


            continue



        excel_cw = float(
            match_excel.iloc[0]["CW"]
        )



        # ==========================================
        # COMPARE CW
        # ==========================================

        if pdf_cw is None:


            results.append({

                "HAWB": hawb_from_filename,
                "Excel CW": excel_cw,
                "PDF CW": "",
                "Difference": "",
                "Result": "CW NOT FOUND IN PDF",
                "PDF File": pdf_file

            })


            continue



        difference = round(
            abs(excel_cw - pdf_cw),
            2
        )



        if difference <= 0.01:

            result = "PASS"

        else:

            result = "FAIL"



        results.append({

            "HAWB": hawb_from_filename,
            "Excel CW": excel_cw,
            "PDF CW": pdf_cw,
            "Difference": difference,
            "Result": result,
            "PDF File": pdf_file

        })



    # =====================================================
    # SAVE RESULTS
    # =====================================================

    os.makedirs(
        OUTPUT_FOLDER,
        exist_ok=True
    )


    output_file = os.path.join(
        OUTPUT_FOLDER,
        "AWB_Validation_Result.xlsx"
    )


    df_results = pd.DataFrame(results)


    df_results.to_excel(
        output_file,
        index=False
    )



    return output_file