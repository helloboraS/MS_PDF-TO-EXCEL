import streamlit as st
import pandas as pd
import pdfplumber
import tempfile
import os

def extract_format_a(pdf_path):
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            lines = page.extract_text().split("\n")
            for line in lines:
                parts = line.split()
                if len(parts) >= 12 and parts[2].isdigit() and parts[-4].isdigit():
                    record = {
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
                    records.append(record)
                elif len(parts) >= 3 and parts[0].isdigit() and parts[1].isdigit():
                    if records:
                        records[-1]["HTS Code"] = parts[1]
                        records[-1]["HTS Description"] = " ".join(parts[2:])
    return pd.DataFrame(records)

def extract_format_b(pdf_path):
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            lines = page.extract_text().split("\n")
            i = 0
            while i < len(lines) - 1:
                line1 = lines[i].strip().split()
                line2 = lines[i + 1].strip().split()

                if not (line1 and line2):
                    i += 1
                    continue
                try:
                    delivery_no = line1[1]
                    msf_index = next(j for j, p in enumerate(line1) if p.startswith("MSF-"))
                    manufacturer_part_no = " ".join(line1[2:msf_index])
                    ms_part_no = line1[msf_index]

                    model_no = line2[2] if len(line2) > 2 else "NA"

                    hts_code = line1[msf_index + 2]
                    country = line1[msf_index + 3]
                    ship_qty = line1[msf_index + 4]
                    unit_price = line1[msf_index + 5]
                    price_uom = line1[msf_index + 6]
                    ext_price = line1[msf_index + 7]

                    desc_start_index = 3 if len(line2) > 3 else None
                    desc_raw = " ".join(line2[desc_start_index:]) if desc_start_index else ""
                    desc_clean = desc_raw.replace("NEW NLR", "").strip()

                    record = {
                        "Delivery No.": delivery_no,
                        "Manufacturer Part No.": manufacturer_part_no,
                        "Model No": model_no,
                        "Microsoft Part No.": ms_part_no,
                        "HTS Code": hts_code,
                        "Country of Origin": country,
                        "Ship Qty": ship_qty,
                        "Unit Price": unit_price,
                        "Price UOM": price_uom,
                        "Extended Price": ext_price,
                        "Part Description": desc_clean
                    }
                    records.append(record)
                    i += 2
                except Exception:
                    i += 1
    return pd.DataFrame(records)

st.set_page_config(page_title="MS Helper", layout="wide")
st.title("Microsoft Helper ♥")

tab1, tab2, tab3, tab4 = st.tabs(["📘 MS1056", "📗 MS1279-PAYMENTS", "📒 MS1279-MASTER 비교", "📕 MS1279-WESCO"])

with tab1:
    uploaded_files_a = st.file_uploader("MS1056 PDF Upload", type=["pdf"], accept_multiple_files=True, key="a")
    if uploaded_files_a:
        all_data = {}
        for uploaded_file in uploaded_files_a:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                tmp_file.write(uploaded_file.read())
                temp_pdf_path = tmp_file.name
            df = extract_format_a(temp_pdf_path)
            os.remove(temp_pdf_path)
            sheet_name = os.path.splitext(uploaded_file.name)[0][:31]
            all_data[sheet_name] = df
            st.subheader(f"{sheet_name}")
            st.dataframe(df)

        if all_data:
            excel_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            with pd.ExcelWriter(excel_file.name, engine="openpyxl") as writer:
                # MASTER DESC + ORIGIN 열 추가용 매핑 생성
                part_desc_map = {}
                if "master_df" in st.session_state:
                    master_df = st.session_state["master_df"]
                    master_df["Microsoft Part No."] = master_df["Microsoft Part No."].astype(str).str.strip()
                    part_desc_map = master_df.set_index("Microsoft Part No.")["Part Description"].to_dict()

                for name, df in all_data.items():
                    df["MASTER DESC + ORIGIN"] = df.apply(
                        lambda row: 
                            part_desc_map.get(row["Microsoft Part No."], "") +
                            (" MODEL: " + row["Model No"] if row["Model No"] != "NA" else "") +
                            " ORIGIN: " + row["Country of Origin"],
                        axis=1
                    )
                    df.to_excel(writer, sheet_name=name, index=False)

            with open(excel_file.name, "rb") as f:
                st.download_button(
                    label="📥 MS1056 엑셀 다운로드",
                    data=f,
                    file_name="ms1056_data.xlsx"
                )

with tab2:
    uploaded_files_b = st.file_uploader("MS1279 PDF Upload", type=["pdf"], accept_multiple_files=True, key="b")
    if uploaded_files_b:
        all_data = {}
        for uploaded_file in uploaded_files_b:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                tmp_file.write(uploaded_file.read())
                temp_pdf_path = tmp_file.name
            df = extract_format_b(temp_pdf_path)
            os.remove(temp_pdf_path)
            sheet_name = os.path.splitext(uploaded_file.name)[0][:31]
            all_data[sheet_name] = df
            st.subheader(f"{sheet_name}")
            st.dataframe(df)
        if all_data:
            excel_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            with pd.ExcelWriter(excel_file.name, engine="openpyxl") as writer:
                for name, df in all_data.items():
                    df.to_excel(writer, sheet_name=name, index=False)
            with open(excel_file.name, "rb") as f:
                st.download_button(
                    label="📥 MS1279 엑셀 다운로드",
                    data=f,
                    file_name="ms1279_data.xlsx"
                )
