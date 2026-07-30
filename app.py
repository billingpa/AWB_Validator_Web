import streamlit as st
import os
import tempfile

from validator import validate_awb


if "reset_counter" not in st.session_state:
    st.session_state.reset_counter = 0

st.title("Samsung SDS - AWB Validator SDSPA")

if st.button("🔄 Reset / New Batch"):

    st.session_state.reset_counter += 1

    st.rerun()

st.write(
    "Upload the Excel master file and PDF invoices for validation."
)


excel_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx"],
    key=f"excel_{st.session_state.reset_counter}"
)


pdf_files = st.file_uploader(
    "Upload PDF files",
    type=["pdf"],
    accept_multiple_files=True,
    key=f"pdf_{st.session_state.reset_counter}"
)


if st.button("▶ Run Validation"):


    if excel_file is None:

        st.warning(
            "Please upload the Excel file."
        )

    elif len(pdf_files) == 0:

        st.warning(
            "Please upload PDF files."
        )

    else:

        with tempfile.TemporaryDirectory() as temp_folder:


            input_folder = os.path.join(
                temp_folder,
                "input"
            )

            output_folder = os.path.join(
                temp_folder,
                "output"
            )


            os.makedirs(input_folder)
            os.makedirs(output_folder)


            excel_path = os.path.join(
                input_folder,
                excel_file.name
            )


            with open(excel_path, "wb") as f:

                f.write(
                    excel_file.getbuffer()
                )


            for pdf in pdf_files:


                pdf_path = os.path.join(
                    input_folder,
                    pdf.name
                )


                with open(pdf_path, "wb") as f:

                    f.write(
                        pdf.getbuffer()
                    )


            result = validate_awb(
                input_folder,
                output_folder
            )


            st.success(
                "Validation completed!"
            )


            with open(result, "rb") as file:

                st.download_button(
                    label="Download Validation Result",
                    data=file,
                    file_name="AWB_Validation_Result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )