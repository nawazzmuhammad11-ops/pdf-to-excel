import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="PDF to Excel Converter", page_icon="📄")

st.title("📄 PDF to Excel Converter")
st.write("Apni PDF file upload karein aur usay Excel spreadsheet mein convert karein.")

uploaded_file = st.file_uploader("PDF File Choose Karein", type=["pdf"])

if uploaded_file is not None:
    st.success("File Successfully Upload Ho Gayi!")
    
    if st.button("Excel Mein Convert Karein"):
        with st.spinner("PDF Processing Ho Rahi Hai..."):
            all_tables = []
            all_text_lines = []
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        df = pd.DataFrame(table)
                        all_tables.append(df)
                    
                    text = page.extract_text()
                    if text:
                        for line in text.split('\n'):
                            if line.strip():
                                all_text_lines.append([line.strip()])

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                if all_tables:
                    combined_df = pd.concat(all_tables, ignore_index=True)
                    combined_df.to_excel(writer, index=False, header=False, sheet_name="Tables")
                    st.success("Table Structure Detect Ho Gaya!")
                elif all_text_lines:
                    text_df = pd.DataFrame(all_text_lines, columns=["Document Content"])
                    text_df.to_excel(writer, index=False, sheet_name="Text Content")
                    st.success("Text Document Successfully Convert Ho Gaya!")
            
            output.seek(0)
            
            if all_tables or all_text_lines:
                st.download_button(
                    label="📥 Excel File Download Karein",
                    data=output,
                    file_name="converted_document.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("PDF se koi text ya data extract nahi ho saka.")
