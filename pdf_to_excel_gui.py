
import streamlit as st
import pandas as pd
import pdfplumber
import tempfile
import os

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

st.set_page_config(page_title="MS PAYMENT 신고서", layout="wide")
st.title("📗 MS1279-PAYMENTS")

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
            merged_df = pd.concat(all_data.values(), ignore_index=True)
            filtered_df = pd.DataFrame({
                "HS CODE": merged_df["HTS Code"],
                "DESC + ORIGIN": merged_df.apply(
                    lambda row: row["Part Description"]
                    + (" MODEL: " + row["Model No"] if row["Model No"] != "NA" else "")
                    + " ORIGIN: " + row["Country of Origin"], axis=1),
                "PART NO.": "PART NO: " + merged_df["Microsoft Part No."] + " (" + merged_df["Manufacturer Part No."] + ")",
                "Q'TY": merged_df["Ship Qty"],
                "UOM": merged_df["Price UOM"],
                "UNIT PRICE": merged_df["Unit Price"],
                "TOTAL AMOUNT": merged_df["Extended Price"],
                "PART NO. FULL": merged_df["Microsoft Part No."] + " (" + merged_df["Manufacturer Part No."] + ")",
                "Model No": merged_df["Model No"]
            })

            # MASTER-DES 추가
            master_df = st.session_state.get("master_df")
            if master_df is not None:
                master_part_desc_map = master_df.set_index("Microsoft Part No.")["Part Description"].to_dict()
                filtered_df["MASTER-DES"] = merged_df.apply(
                    lambda row: master_part_desc_map.get(row["Microsoft Part No."], "미확인")
                    + (" MODEL: " + row["Model No"] if row["Model No"] != "NA" else "")
                    + " ORIGIN: " + row["Country of Origin"], axis=1
                )
            else:
                filtered_df["MASTER-DES"] = "미확인"

            filtered_df.to_excel(writer, sheet_name="신고서용", index=False)

        with open(excel_file.name, "rb") as f:
            st.download_button(
                label="📥 MS1279-PAYMENTS 엑셀 다운로드",
                data=f,
                file_name="ms1279_payments_data.xlsx"
            )
