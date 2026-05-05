import streamlit as st
import os
from excel_parser.parser_v3 import parse_file

st.set_page_config(page_title="Excel Parser", layout="wide")

st.title("Excel → JSON Parser (EMEA)")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    os.makedirs("cache", exist_ok=True)
    os.makedirs("parsed", exist_ok=True)

    file_path = os.path.join("cache", uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.read())

    st.success(f"File saved to {file_path}")

    if st.button("Parse file"):
        try:
            output_path = os.path.join("parsed", "emea.json")
            result = parse_file(file_path, sheet_name="EMEA", output_path=output_path)

            st.success("Parsing completed")
            st.write(f"Output saved to {output_path}")

            st.json(result[:10])

            with open(output_path, "rb") as f:
                st.download_button("Download JSON", f, file_name="emea.json")

        except Exception as e:
            st.error(str(e))
