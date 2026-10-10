import streamlit as st
import pdfplumber
import pandas as pd
import io
import pytesseract
from pdf2image import convert_from_bytes

st.set_page_config(page_title="Advanced PDF to Excel Converter", page_icon="📄")

st.title("📄 Advanced PDF to Excel Converter")
st.write("Ek se zyada PDFs upload karein aur unhe Excel mein convert karein (Scanned documents bhi supported hain).")

# Multiple files uploader
uploaded_files = st.file_uploader("PDF Files Choose Karein (Multiple Supported)", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    st.success(f"{len(uploaded_files)} File(s) Successfully Upload Ho Gayi!")
    
    if st.button("Sabhi Files Ko Convert Karein"):
        with st.spinner("PDFs Processing Ho Rahi Hain..."):
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for idx, uploaded_file in enumerate(uploaded_files):
                    file_name = f"File_{idx+1}"
                    all_tables = []
                    all_text_lines = []
                    
                    # 1. Standard Extraction (Structured Tables & Text)
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

                    # 2. OCR Fallback for Scanned Images/PDFs if no text or tables found
                    if not all_tables and not all_text_lines:
                        try:
                            images = convert_from_bytes(uploaded_file.read())
                            for img in images:
                                ocr_text = pytesseract.image_to_string(img)
                                if ocr_text:
                                    for line in ocr_text.split('\n'):
                                        if line.strip():
                                            all_text_lines.append([line.strip()])
                        except Exception as e:
                            st.warning(f"{uploaded_file.name} ke OCR mein dikkat aayi: {e}")

                    # Save result to separate Excel sheet per file
                    if all_tables:
                        combined_df = pd.concat(all_tables, ignore_index=True)
                        combined_df.to_excel(writer, index=False, header=False, sheet_name=f"{file_name}_Tables")
                    elif all_text_lines:
                        text_df = pd.DataFrame(all_text_lines, columns=["Document Content"])
                        text_df.to_excel(writer, index=False, sheet_name=f"{file_name}_Text")
                    else:
                        empty_df = pd.DataFrame(["Data extracted nahi ho saka"], columns=["Error"])
                        empty_df.to_excel(writer, index=False, sheet_name=f"{file_name}_Empty")

            output.seek(0)
            
            st.download_button(
                label="📥 Convert Shuda Excel File Download Karein",
                data=output,
                file_name="converted_documents.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
