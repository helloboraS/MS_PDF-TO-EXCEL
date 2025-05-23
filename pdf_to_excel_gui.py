import streamlit as st
import pandas as pd
import pdfplumber
import tempfile
import os
import re

def extract_format_a(pdf_path):
    records = []
    current_record = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            lines = page.extract_text().split("\n")
            for line in lines:
                parts = line.split()
                if len(parts) >= 12 and parts[2].isdigit() and parts[-4].isdigit():
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
                elif len(parts) >= 3 and parts[0].isdigit() and parts[1].isdigit():
                    if records:
                        records[-1]["HTS Code"] = parts[1]
                        records[-1]["HTS Description"] = " ".join(parts[2:])

    df = pd.DataFrame(records)
    column_order = [
        "PO No", "SAP Order No", "Part Number", "Part Description",
        "Ship Qty", "Price UOM", "Unit Price", "Extended Price",
        "Model No", "HTS Code", "Country of Origin", "HTS Description"
    ]
    for col in column_order:
        if col not in df.columns:
            df[col] = ""
    return df[column_order]

def extract_format_b(pdf_path):
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            lines = page.extract_text().split("\n")
            for i in range(len(lines) - 2):
                line = lines[i].strip()
                model_line = lines[i + 1].strip()
                desc_line = lines[i + 2].strip()
                if re.match(r"^\d+\s+\d{10}\s+\S+\s+\S+\s+\S+\s+\d+\s+[A-Z]{2}\s+\d+\s+\d+\s+EA\s+\d+$", line):
                    parts = line.split()
                    model_parts = model_line.split()
                    record = {
                        "Delivery No.": parts[1],
                        "Manufacturer Part No.": parts[2],
                        "Model No": model_parts[0] if model_parts else "",
                        "Microsoft Part No.": parts[3],
                        "HTS Code": parts[6],  # 원래 Country of Origin 값
                        "Country of Origin": parts[7],  # 원래 Ship Qty 값
                        "Ship Qty": parts[8],  # 원래 Unit Price 값
                        "Unit Price": parts[9],
                        "Price UOM": parts[10],
                        "Extended Price": parts[11],
                        "Part Description": desc_line
                    }
                    records.append(record)

    df = pd.DataFrame(records)
    column_order = [
        "Delivery No.", "Manufacturer Part No.", "Model No",
        "Microsoft Part No.", "HTS Code", "Country of Origin",
        "Ship Qty", "Unit Price", "Price UOM", "Extended Price",
        "Part Description"
    ]
    for col in column_order:
        if col not in df.columns:
            df[col] = ""
    return df[column_order]

# Streamlit 앱 UI
st.set_page_config(page_title="PDF 항목 추출기", layout="wide")
st.title("📄 PDF → Excel 항목 추출기")

tab1, tab2 = st.tabs(["📘 MS1056", "📗 MS1279-PAYMENTS"])

with tab1:
    uploaded_files_a = st.file_uploader("[MS1056] PDF 파일을 하나 이상 업로드하세요", type=["pdf"], accept_multiple_files=True, key="a")
    if uploaded_files_a:
        with st.spinner("PDF에서 항목 추출 중..."):
            all_data = {}
            try:
                for uploaded_file in uploaded_files_a:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                        tmp_file.write(uploaded_file.read())
                        temp_pdf_path = tmp_file.name

                    df = extract_format_a(temp_pdf_path)
                    os.remove(temp_pdf_path)
                    sheet_name = os.path.splitext(uploaded_file.name)[0][:31]
                    all_data[sheet_name] = df

                st.success("✅ MS1056 PDF 추출 완료")
                for name, df in all_data.items():
                    st.subheader(f"📄 {name}")
                    st.dataframe(df)

                excel_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                with pd.ExcelWriter(excel_file.name, engine="openpyxl") as writer:
                    for name, df in all_data.items():
                        df.to_excel(writer, sheet_name=name, index=False)

                with open(excel_file.name, "rb") as f:
                    st.download_button(
                        label="📥 MS1056 엑셀 다운로드",
                        data=f,
                        file_name="ms1056_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"❌ 오류 발생: {e}")

with tab2:
    uploaded_files_b = st.file_uploader("[MS1279-PAYMENTS] PDF 파일을 하나 이상 업로드하세요", type=["pdf"], accept_multiple_files=True, key="b")
    if uploaded_files_b:
        all_data = {}
        st.subheader("🔍 미리보기 결과")
        try:
            for uploaded_file in uploaded_files_b:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                    tmp_file.write(uploaded_file.read())
                    temp_pdf_path = tmp_file.name

                df = extract_format_b(temp_pdf_path)
                os.remove(temp_pdf_path)
                sheet_name = os.path.splitext(uploaded_file.name)[0][:31]
                all_data[sheet_name] = df

                st.write(f"📄 {sheet_name}")
                st.dataframe(df)

            if all_data:
                excel_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                with pd.ExcelWriter(excel_file.name, engine="openpyxl") as writer:
                    for name, df in all_data.items():
                        df.to_excel(writer, sheet_name=name, index=False)

                with open(excel_file.name, "rb") as f:
                    st.download_button(
                        label="📥 MS1279-PAYMENTS 엑셀 다운로드",
                        data=f,
                        file_name="ms1279_payments_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"❌ 오류 발생: {e}")
