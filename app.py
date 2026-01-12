import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import zipfile
import re

st.set_page_config(page_title="Universal DOCX Generator", layout="centered")

# ===================== SIDEBAR =====================
st.sidebar.title("Excel Rules")

st.sidebar.markdown("""
                    
### Mandatory Columns
- **1st Column:** ECode
- **2nd Column:** Name

These two columns are compulsory and must be in this exact order.
""")

# ---------- SAMPLE EXCEL ----------
sample_df = pd.DataFrame({
    "ECode": ["EMP001"],
    "Name": ["Neeraj Balodi"],
    "Designation": ["Manager"],
    "Department": ["IT"],
    "Month": ["Jan 2025"],
    "Amount": [5000],
    "Remarks": ["Best Performer"]
})

sample_buffer = BytesIO()
sample_df.to_excel(sample_buffer, index=False, engine="openpyxl")
sample_buffer.seek(0)

st.sidebar.download_button(
    label="⬇️ Download Sample Excel",
    data=sample_buffer,
    file_name="Sample_Employee_Data.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ===================== MAIN ==================================
st.title("Universal DOCX Template Generator")
st.write("Excel headers automatically replace matching placeholders in DOCX")

template_file = st.file_uploader("Upload DOCX Template", type=["docx"])
excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])

# ===================== PLACEHOLDER ENGINE =====================
def replace_placeholders(doc, data):
    pattern = re.compile(r"\{\s*(.*?)\s*\}")

    # -------- PARAGRAPHS --------
    for para in doc.paragraphs:
        full_text = "".join(run.text for run in para.runs)
        if not full_text:
            continue

        updated_text = full_text
        matches = pattern.findall(full_text)

        for match in matches:
            key = match.strip()
            if key in data:
                updated_text = updated_text.replace(
                    f"{{{match}}}", str(data[key])
                )

        if updated_text != full_text:
            for run in para.runs:
                run.text = ""
            para.add_run(updated_text)

    # -------- TABLES --------
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    full_text = "".join(run.text for run in para.runs)
                    if not full_text:
                        continue

                    updated_text = full_text
                    matches = pattern.findall(full_text)

                    for match in matches:
                        key = match.strip()
                        if key in data:
                            updated_text = updated_text.replace(
                                f"{{{match}}}", str(data[key])
                            )

                    if updated_text != full_text:
                        for run in para.runs:
                            run.text = ""
                        para.add_run(updated_text)

# ===================== PROCESS =====================
if template_file and excel_file:
    df = pd.read_excel(excel_file)

    # ---------- VALIDATION ----------
    if df.shape[1] < 2:
        st.error("❌ Excel must contain at least ECode and Name columns.")
        st.stop()

    if df.columns[0].strip() != "ECode":
        st.error("❌ First column must be exactly 'ECode'.")
        st.stop()

    if df.columns[1].strip() != "Name":
        st.error("❌ Second column must be exactly 'Name'.")
        st.stop()

    st.success("✅ Excel structure validated successfully")

    st.subheader("📊 Detected Excel Headers")
    st.write(list(df.columns))

    if st.button("Generate Documents (ZIP)"):
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for _, row in df.iterrows():
                doc = Document(template_file)

                row_data = {
                    col: "" if pd.isna(row[col]) else row[col]
                    for col in df.columns
                }

                replace_placeholders(doc, row_data)

                doc_buffer = BytesIO()
                doc.save(doc_buffer)
                doc_buffer.seek(0)

                filename = f"{row['ECode']}_{row['Name']}.docx"
                zip_file.writestr(filename, doc_buffer.read())

        zip_buffer.seek(0)

        st.download_button(
            label="Download Generated Documents (ZIP)",
            data=zip_buffer,
            file_name="Generated_Documents.zip",
            mime="application/zip"
        )
