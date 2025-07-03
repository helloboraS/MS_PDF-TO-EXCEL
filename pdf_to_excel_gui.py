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
                    msf_index = next(j for j, p in enumerate(line1) if p.startswith("MSF-"))
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

tab1, tab2, tab3, tab4, tab5 = st.tabs(['📘 MS1056', '📗 MS1279-PAYMENTS', '📒 MS1279-MASTER 비교', '📕 MS1279-WESCO', '📙 HS코드 비교'])

with tab1:
    st.write("탭1: 기존 구현된 MS1056 기능")

with tab2:
    st.write("탭2: 기존 구현된 MS1279-PAYMENTS 기능")

with tab3:
    st.write("탭3: 기존 구현된 MASTER 비교 기능")

with tab4:
    st.write("탭4: 기존 구현된 WESCO 기능")

with tab5:
    st.subheader("Microsoft Part No.와 INV HS 코드 비교")

    part_no_input = st.text_area("Microsoft Part No. 입력 (쉼표로 구분)", placeholder="예: MSF-000001, MSF-000002")
    inv_hs_input = st.text_area("INV HS 입력 (쉼표로 구분)", placeholder="예: 8473301000, 8473309000")

    if st.button("비교 실행"):
        if not os.path.exists("MASTER_MS5673.xlsx"):
            st.error("MASTER_MS5673.xlsx 파일이 없습니다.")
        else:
            master_df = pd.read_excel("MASTER_MS5673.xlsx")
            part_nos = [x.strip() for x in part_no_input.split(",")]
            inv_hs_codes = [x.strip().replace("-", "") for x in inv_hs_input.split(",")]

            result = []
            for pno, inv_hs in zip(part_nos, inv_hs_codes):
                row = master_df[master_df["Microsoft Part No."].astype(str).str.strip() == pno]
                hs_code = row["HS Code"].values[0] if not row.empty else ""
                hs_code = str(hs_code).zfill(10)
                match_10 = "O" if inv_hs[:10] == hs_code[:10] else "X"
                match_6 = "O" if inv_hs[:6] == hs_code[:6] else "X"
                result.append({
                    "Microsoft Part No.": pno,
                    "INV HS": inv_hs,
                    "MASTER HS Code": hs_code,
                    "HS10_MATCH": match_10,
                    "HS6_MATCH": match_6
                })
            st.dataframe(pd.DataFrame(result))
