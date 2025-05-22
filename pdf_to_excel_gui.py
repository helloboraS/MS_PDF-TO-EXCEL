
import streamlit as st
import pandas as pd
import pdfplumber
import re
import tempfile
import os

def extract_data_from_pdf(pdf_path):
    raw_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            raw_text += page.extract_text() + "\n"

    lines = raw_text.splitlines()
    records = []
    record = {}

    pos_pattern = re.compile(r"^\d{2,3}\s+OT\d+")
    model_pattern = re.compile(r"^803\d{6,}")
    eccn_pattern = re.compile(r"(EAR99|5A992\.c)")

    for line in lines:
        if pos_pattern.match(line):
            if record:
                records.append(record)
                record = {}
            parts = line.split()
            if len(parts) >= 8:
                record["Pos"] = parts[0]
                record["PO No"] = parts[1]
                record["SAP Order No"] = parts[2]
                record["Part Number"] = parts[3]
                record["Part Description"] = " ".join(parts[4:-4])
                record["Quantity"] = parts[-4]
                record["Country of Origin"] = parts[-3]
                record["Ship Qty"] = parts[-2]
                record["Unit Price"] = parts[-1]
        elif model_pattern.match(line):
            parts = line.split()
            if len(parts) >= 6:
                record["Model No"] = parts[0]
                record["HTS Code"] = parts[1]
                record["HTS Description"] = " ".join(parts[2:-3])
                record["Price UOM"] = parts[-3]
                record["Extended Price"] = parts[-2]
                eccn_match = eccn_pattern.search(line)
                record["ECCN"] = eccn_match.group(0) if eccn_match else ""

    if record:
        records.append(record)

    return pd.DataFrame(records)

st.set_page_config(page_title="PDF 항목 추출기", layout="centered")
st.title("📄 PDF → Excel 항목 추출기")

uploaded_file = st.file_uploader("PDF 파일을 업로드하세요", type=["pdf"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
        tmp_file.write(uploaded_file.read())
        temp_pdf_path = tmp_file.name

    with st.spinner("PDF에서 항목 추출 중..."):
        try:
            df = extract_data_from_pdf(temp_pdf_path)
            os.remove(temp_pdf_path)

            st.success("✅ 추출 완료! 아래에서 미리보기를 확인하세요.")
            st.dataframe(df)

            # 파일 다운로드 링크 생성
            excel_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            df.to_excel(excel_file.name, index=False)
            with open(excel_file.name, "rb") as f:
                st.download_button(
                    label="📥 엑셀 파일 다운로드",
                    data=f,
                    file_name="extracted_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"❌ 오류 발생: {e}")
