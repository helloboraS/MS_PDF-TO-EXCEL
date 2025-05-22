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
                if re.match(r"^\d{2,3}\s+OT", line):
                    parts = line.split()
                    current_record = {
                        "PO No": parts[1],
                        "SAP Order No": parts[2],
                        "Part Number": parts[3],
                        "Part Description": " ".join(parts[4:-6]),
                        "Model No": parts[-6],
                        "Country of Origin": parts[-5],
                        "Ship Qty": parts[-4],
                        "Price UOM": parts[-3],
                        "Unit Price": parts[-2],
                        "Extended Price": parts[-1],
                        "HTS Code": "",
                        "HTS Description": ""
                    }
                    records.append(current_record)

                elif re.match(r"^\d{10}\s+\d{8,10}\s+", line):
                    parts = line.split()
                    if len(parts) >= 3 and records:
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

uploaded_files = st.file_uploader("PDF 파일을 하나 이상 업로드하세요", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    with st.spinner("PDF에서 항목 추출 중..."):
        all_data = {}
        try:
            for uploaded_file in uploaded_files:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                    tmp_file.write(uploaded_file.read())
                    temp_pdf_path = tmp_file.name

                df = extract_data_from_pdf(temp_pdf_path)
                os.remove(temp_pdf_path)
                sheet_name = os.path.splitext(uploaded_file.name)[0][:31]  # Excel 시트명 제한 고려
                all_data[sheet_name] = df

            st.success("✅ 모든 PDF에서 추출 완료! 아래에서 미리보기를 확인하세요.")
            for name, df in all_data.items():
                st.subheader(f"📄 {name}")
                st.dataframe(df)

            excel_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            with pd.ExcelWriter(excel_file.name, engine="openpyxl") as writer:
                for name, df in all_data.items():
                    df.to_excel(writer, sheet_name=name, index=False)

            with open(excel_file.name, "rb") as f:
                st.download_button(
                    label="📥 모든 시트 포함 엑셀 파일 다운로드",
                    data=f,
                    file_name="multiple_extracted_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"❌ 오류 발생: {e}")
