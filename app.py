import streamlit as st
import pandas as pd
import pdfplumber
import io

st.set_page_config(page_title="PDF to Excel Converter", layout="centered")

st.title("📄 PDF to Excel Converter")
st.write("Apni PDF file upload karein aur usay Excel spreadsheet mein convert karein.")

uploaded_file = st.file_uploader("PDF File Choose Karein", type=["pdf"])

if uploaded_file is not None:
    st.success("File Successfully Upload Ho Gayi!")
    
    if st.button("Excel Mein Convert Karein"):
        with st.spinner("Processing... Behar meherbani intazar karein."):
            all_tables = []
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    extracted = page.extract_tables()
                    for table in extracted:
                        df = pd.DataFrame(table)
                        all_tables.append(df)
            
            if all_tables:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for i, df in enumerate(all_tables):
                        df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False, header=False)
                
                output.seek(0)
                
                st.download_button(
                    label="📥 Download Excel File",
                    data=output,
                    file_name="converted_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Is PDF file mein koi table nahi mila. Meherbani karke structured table wali PDF try karein.")

st.markdown("---")
st.caption("Powered by Streamlit | App Status: Active")
