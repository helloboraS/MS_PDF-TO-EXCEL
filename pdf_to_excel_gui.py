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
            for line in lines:
                parts = line.strip().split()
                if len(parts) < 11 or not (parts[0].isdigit() and parts[1].isdigit()):
                    continue
                try:
                    delivery_no = parts[1]

                    # MSF- 시작하는 index 탐색
                    msf_index = next(i for i, p in enumerate(parts) if p.startswith("MSF-"))
                    ms_part_no = parts[msf_index]

                    # Model No 결정
                    if msf_index > 3:
                        model_no = parts[msf_index - 1]
                        manufacturer_part_no = " ".join(parts[2:msf_index - 1])
                    else:
                        model_no = "NA"
                        manufacturer_part_no = " ".join(parts[2:msf_index])

                    # 나머지 필드
                    hts_code = parts[msf_index + 2]
                    country = parts[msf_index + 3]
                    ship_qty = parts[msf_index + 4]
                    unit_price = parts[msf_index + 5]
                    price_uom = parts[msf_index + 6]
                    ext_price = parts[msf_index + 7]

                    desc_start_index = msf_index + 8
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

# Streamlit UI
st.set_page_config(page_title="PDF 항목 추출기", layout="wide")
st.title("📄 PDF → Excel 항목 추출기")

tab2 = st.tabs(["📗 MS1279-PAYMENTS"])[0]

with tab2:
    uploaded_files_b = st.file_uploader("[MS1279-PAYMENTS] PDF 파일 업로드", type=["pdf"], accept_multiple_files=True)
    if uploaded_files_b:
        all_data = {}
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
