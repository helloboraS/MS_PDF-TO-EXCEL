import streamlit as st
import pandas as pd
import pdfplumber
import tempfile
import os
import re

def extract_data_from_pdf(pdf_path):
    records = []
    current_record = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            lines = page.extract_text().split("\n")
            for line in lines:
                # 기본 정보 라인: POS부터 시작하는 줄
                if re.match(r"^\d{2,3}\s+OT", line):
                    parts = line.split()
                    current_record = {
                        "PO No": parts[1],
                        "SAP Order No": parts[2],
                        "Part Number": parts[3],
                        "Part Description": " ".join(parts[4:-5]),
                        "Country of Origin": parts[-5],  # 위치 수정
                        "Ship Qty": parts[-4],            # 위치 수정
                        "Price UOM": parts[-3],
                        "Unit Price": parts[-2],
                        "Extended Price": parts[-1],
                        "Model No": "",
                        "HTS Code": "",
                        "HTS Description": ""
                    }
                    records.append(current_record)

                # 세부 정보 라인: Model No, HTS Code 등
                elif re.match(r"^\d{4}\s+\d{8,10}", line):
                    parts = line.split()
                    if len(parts) >= 3 and records:
                        records[-1]["Model No"] = parts[0]
                        records[-1]["HTS Code"] = parts[1]
                        records[-1]["HTS Description"] = " ".join(parts[2:])

    column_order = [
        "PO No", "SAP Order No", "Part Number", "Part Description",
        "Ship Qty", "Price UOM", "Unit Price", "Extended Price",
        "Model No", "HTS Code", "Country of Origin", "HTS Description"
    ]

    df = pd.DataFrame(records)
    for col in column_order:
        if col not in df.columns:
            df[col] = ""
    df = df[column_order]
    return df

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
