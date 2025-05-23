import streamlit as st
import pandas as pd
import pdfplumber
import tempfile
import os
import re

def extract_format_b(pdf_path):
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            lines = page.extract_text().split("\n")
            for line in lines:
                parts = line.strip().split()
                if len(parts) >= 11 and parts[0].isdigit() and parts[1].isdigit():
                    try:
                        delivery_no = parts[1]
                        manufacturer_part_no = parts[2]
                        # Model No 추정
                        if parts[3].startswith("MSF-"):
                            model_no = "NA"
                            ms_part_no = parts[3]
                            hts_code = parts[5]
                            country = parts[6]
                            ship_qty = parts[7]
                            unit_price = parts[8]
                            price_uom = parts[9]
                            ext_price = parts[10]
                            desc_start_index = 11
                        else:
                            model_no = parts[3]
                            ms_part_no = parts[4]
                            hts_code = parts[6]
                            country = parts[7]
                            ship_qty = parts[8]
                            unit_price = parts[9]
                            price_uom = parts[10]
                            ext_price = parts[11]
                            desc_start_index = 12

                        desc_raw = " ".join(parts[desc_start_index:]) if len(parts) > desc_start_index else ""
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
                    except Exception:
                        continue
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

# Streamlit 앱 UI 구성
st.set_page_config(page_title="PDF 항목 추출기", layout="wide")
st.title("📄 PDF → Excel 항목 추출기")

tab1, tab2 = st.tabs(["📘 MS1056", "📗 MS1279-PAYMENTS"])

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
                        "PART NO. FULL": merged_df["Microsoft Part No."] + " (" + merged_df["Manufacturer Part No."] + ")"
                    })
                    filtered_df.to_excel(writer, sheet_name="신고서용", index=False)

                with open(excel_file.name, "rb") as f:
                    st.download_button(
                        label="📥 MS1279-PAYMENTS 엑셀 다운로드",
                        data=f,
                        file_name="ms1279_payments_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"❌ 오류 발생: {e}")
