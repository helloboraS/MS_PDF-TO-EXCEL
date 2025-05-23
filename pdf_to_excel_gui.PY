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
            for i in range(len(lines)):
                line = lines[i].strip()
                match = re.match(r"^(\d{10})\s+([\w\-]+)\s+(\w+)\s+(MSF-\d+)\s+(\d+)\s+(CN|TH|US|SG|KR)\s+(\d+)\s+(\d+)\s+EA\s+(\d+)", line)
                if match and i + 1 < len(lines):
                    records.append({
                        "Delivery No.": match.group(1),
                        "Manufacturer Part No.": match.group(2),
                        "Model No": match.group(3),
                        "Microsoft Part No.": match.group(4),
                        "HTS Code": match.group(5),
                        "Country of Origin": match.group(6),
                        "Ship Qty": match.group(7),
                        "Unit Price": match.group(8),
                        "Price UOM": "EA",
                        "Extended Price": match.group(9),
                        "Part Description": lines[i + 1].strip()
                    })

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

st.set_page_config(page_title="PDF 항목 추출기", layout="wide")
st.title("\ud83d\udcc4 PDF \u2192 Excel \ud56d\ubaa9 \ucd94\ucd9c\uae30")

tab1, tab2 = st.tabs(["\ud83d\udcd8 MS1056", "\ud83d\udcd7 MS1279-PAYMENTS"])

with tab1:
    uploaded_files_a = st.file_uploader("[MS1056] PDF \ud30c\uc77c\uc744 \ud558\ub098 \uc774\uc0c1 \uc5c5\ub85c\ub4dc\ud558\uc138\uc694", type=["pdf"], accept_multiple_files=True, key="a")
    if uploaded_files_a:
        with st.spinner("PDF\uc5d0\uc11c \ud56d\ubaa9 \ucd94\ucd9c \uc911..."):
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

                st.success("\u2705 MS1056 PDF \ucd94\ucd9c \uc644\ub8cc")
                for name, df in all_data.items():
                    st.subheader(f"\ud83d\udcc4 {name}")
                    st.dataframe(df)

                excel_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                with pd.ExcelWriter(excel_file.name, engine="openpyxl") as writer:
                    for name, df in all_data.items():
                        df.to_excel(writer, sheet_name=name, index=False)

                with open(excel_file.name, "rb") as f:
                    st.download_button(
                        label="\ud83d\udcc5 MS1056 \uc5d8\uc140 \ub2e4\uc6b4\ub85c\ub4dc",
                        data=f,
                        file_name="ms1056_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"\u274c \uc624\ub958 \ubc1c\uc0dd: {e}")

with tab2:
    uploaded_files_b = st.file_uploader("[MS1279-PAYMENTS] PDF \ud30c\uc77c\uc744 \ud558\ub098 \uc774\uc0c1 \uc5c5\ub85c\ub4dc\ud558\uc138\uc694", type=["pdf"], accept_multiple_files=True, key="b")
    if uploaded_files_b:
        all_data = {}
        st.subheader("\ud83d\udd0d \ubbf8\ub9ac\ubcf4\uae30 \uacb0\uacfc")
        try:
            for uploaded_file in uploaded_files_b:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                    tmp_file.write(uploaded_file.read())
                    temp_pdf_path = tmp_file.name

                df = extract_format_b(temp_pdf_path)
                os.remove(temp_pdf_path)
                sheet_name = os.path.splitext(uploaded_file.name)[0][:31]
                all_data[sheet_name] = df

                st.write(f"\ud83d\udcc4 {sheet_name}")
                st.dataframe(df)

            if all_data:
                excel_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                with pd.ExcelWriter(excel_file.name, engine="openpyxl") as writer:
                    for name, df in all_data.items():
                        df.to_excel(writer, sheet_name=name, index=False)

                with open(excel_file.name, "rb") as f:
                    st.download_button(
                        label="\ud83d\udcc5 MS1279-PAYMENTS \uc5d8\uc140 \ub2e4\uc6b4\ub85c\ub4dc",
                        data=f,
                        file_name="ms1279_payments_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"\u274c \uc624\ub958 \ubc1c\uc0dd: {e}")
