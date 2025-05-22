import streamlit as st
import pandas as pd
import pdfplumber
import tempfile
import os

def extract_data_from_pdf(pdf_path):
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    # Skip invalid or header rows
                    if row and row[0] and row[0].isdigit():
                        # Clean row and ensure it's the right length
                        clean_row = [cell.strip() if cell else "" for cell in row]
                        if len(clean_row) >= 15:
                            record = {
                                "PO No": clean_row[1],
                                "SAP Order No": clean_row[2],
                                "Part Number": clean_row[3],
                                "Part Description": clean_row[4],
                                "Ship Qty": clean_row[11],
                                "Price UOM": clean_row[12],
                                "Unit Price": clean_row[13],
                                "Extended Price": clean_row[14],
                                "Model No": "",
                                "HTS Code": "",
                                "Country of Origin": clean_row[10],
                                "HTS Description": ""
                            }
                            records.append(record)
                    elif row and len(row) >= 6 and row[0].startswith("8"):
                        if records:
                            records[-1]["Model No"] = row[0]
                            records[-1]["HTS Code"] = row[1]
                            records[-1]["HTS Description"] = row[2]
    
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