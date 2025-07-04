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
                                                                    for name, df in all_data.items():
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
                                                                                                        filtered_df.to_excel(writer, sheet_name="신고서용", index=False)
                                                                                                        with open(excel_file.name, "rb") as f:
                                                                                                            st.download_button(
                                                                                                            label="📥 MS1279-PAYMENTS 엑셀 다운로드",
                                                                                                            data=f,
                                                                                                            file_name="ms1279_payments_data.xlsx"
                                                                                                            )

                                                                                                            with tab3:


                                                                                                                # st.header("📒 마스터 데이터 비교")

                                                                                                                if "master_df" not in st.session_state:
                                                                                                                    if not os.path.exists("MASTER_MS5673.xlsx"):
                                                                                                                        st.warning("⚠️ MASTER_MS5673.xlsx 파일이 현재 디렉토리에 존재하지 않습니다.")
                                                                                                                        if os.path.exists("MASTER_MS5673.xlsx"):
                                                                                                                            st.session_state["master_df"] = pd.read_excel("MASTER_MS5673.xlsx")

                                                                                                                            uploaded_excel = st.file_uploader("엑셀업로드", type=["xlsx"], key="compare_excel")

                                                                                                                            master_df = st.session_state.get("master_df")

                                                                                                                            def clean_code(code):
                                                                                                                                return str(code).strip().replace("-", "")

                                                                                                                            def fix_hscode(code):
                                                                                                                                try:
                                                                                                                                    code_str = str(code)
                                                                                                                                    if code_str.endswith(".0"):
                                                                                                                                        code_str = code_str[:-2]
                                                                                                                                        return code_str.zfill(10)
                                                                                                                                    except:
                                                                                                                                        return ""        

                                                                                                                                    if uploaded_excel and master_df is not None:
                                                                                                                                        input_df = pd.read_excel(uploaded_excel)
                                                                                                                                        master_df = master_df.rename(columns=lambda x: x.strip())
                                                                                                                                        input_df = input_df.rename(columns=lambda x: x.strip())

                                                                                                                                        input_df["Microsoft Part No."] = input_df["Microsoft Part No."].astype(str).str.strip()
                                                                                                                                        master_df["Microsoft Part No."] = master_df["Microsoft Part No."].astype(str).str.strip()

                                                                                                                                        merged = input_df.merge(master_df, how="left", on="Microsoft Part No.")
                                                                                                                                        merged["INV HS"] = merged["INV HS"].apply(clean_code)



                                                                                                                                        merged["HS Code"] = merged["HS Code"].apply(clean_code).apply(fix_hscode)

                                                                                                                                        merged["HS10_MATCH"] = merged.apply(lambda row: "O" if row["INV HS"][:10] == row["HS Code"][:10] else "X", axis=1)
                                                                                                                                        merged["HS6_MATCH"] = merged.apply(lambda row: "O" if row["INV HS"][:6] == row["HS Code"][:6] else "X", axis=1)

                                                                                                                                        merged["전파"] = merged["전파인증번호"].apply(lambda x: "O" if pd.notna(x) and str(x).strip() else "X")
                                                                                                                                        merged["전기"] = merged["전기인증번호"].apply(lambda x: "O" if pd.notna(x) and str(x).strip() else "X")

                                                                                                                                        final_df = merged.copy()

                                                                                                                                        invoice_sheet = pd.DataFrame({
                                                                                                                                        "HS Code": final_df["HS Code"],
                                                                                                                                        "Part Description": final_df["Part Description"] + ' ORIGIN:' + final_df["원산지"],
                                                                                                                                        "Microsoft Part No.": "PART NO: " + final_df["Microsoft Part No."],
                                                                                                                                        "수량": final_df["수량"],
                                                                                                                                        "단위": final_df["단위"],
                                                                                                                                        "단가": final_df["단가"],
                                                                                                                                        "금액": final_df["금액"],
                                                                                                                                        "Microsoft Part No. (원본)": final_df["Microsoft Part No."],  # ← 추가된 열
                                                                                                                                        "전파": final_df["전파"],
                                                                                                                                        "전기": final_df["전기"],
                                                                                                                                        "요건비대상사유": final_df["요건비대상"]
                                                                                                                                        })

                                                                                                                                        radio_req = (
                                                                                                                                        final_df.groupby(["HS Code", "원산지", "모델명", "전파인증번호"], as_index=False)
                                                                                                                                        .agg({"수량": "sum"})
                                                                                                                                        )

                                                                                                                                        safety_req = (
                                                                                                                                        final_df.groupby(["기관", "HS Code", "원산지", "모델명", "전기인증번호", "정격전압"], as_index=False)
                                                                                                                                        .agg({"수량": "sum"})
                                                                                                                                        )

                                                                                                                                        to_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                                                                                                                                        with pd.ExcelWriter(to_excel.name, engine="openpyxl") as writer:
                                                                                                                                            final_df.drop(columns=["무역거래처상호"], errors="ignore").to_excel(writer, index=False, sheet_name="비교결과")
                                                                                                                                            invoice_sheet.to_excel(writer, index=False, sheet_name="신고서")
                                                                                                                                            radio_req.to_excel(writer, index=False, sheet_name="전파요건")
                                                                                                                                            safety_req.to_excel(writer, index=False, sheet_name="전안요건")

                                                                                                                                            with open(to_excel.name, "rb") as f:
                                                                                                                                                st.download_button(
                                                                                                                                                label="📥 비교 결과 엑셀 다운로드",
                                                                                                                                                data=f,
                                                                                                                                                file_name="MS5673_신고.xlsx"
                                                                                                                                                )

                                                                                                                                                elif master_df is not None:
                                                                                                                                                    st.markdown("---")
                                                                                                                                                    #st.subheader("🔍 단일 Microsoft Part No. 수기 비교")

                                                                                                                                                    if "compare_results" not in st.session_state:
                                                                                                                                                        st.session_state.compare_results = []

                                                                                                                                                        with st.form("manual_compare_form"):
                                                                                                                                                            part_no_input = st.text_input("Microsoft Part No. 입력")
                                                                                                                                                            inv_hs_input = st.text_input("INV HS Code 입력")
                                                                                                                                                            submitted = st.form_submit_button("비교하기")

                                                                                                                                                            def clean_hs(code):
                                                                                                                                                                try:
                                                                                                                                                                    code = str(code).strip().replace("-", "")
                                                                                                                                                                    if code.endswith(".0"):
                                                                                                                                                                        code = code[:-2]
                                                                                                                                                                        return code.zfill(10)
                                                                                                                                                                    except:
                                                                                                                                                                        return ""

                                                                                                                                                                    if submitted and part_no_input:
                                                                                                                                                                        part_no = part_no_input.strip()
                                                                                                                                                                        inv_hs = inv_hs_input.strip().replace("-", "")

                                                                                                                                                                        row = master_df[master_df["Microsoft Part No."] == part_no]

                                                                                                                                                                        if not row.empty:
                                                                                                                                                                            hs_code = clean_hs(row.iloc[0]["HS Code"])
                                                                                                                                                                            desc = row.iloc[0].get("Part Description", "")
                                                                                                                                                                            result = {
                                                                                                                                                                            "Microsoft Part No.": part_no,
                                                                                                                                                                            "INV HS": inv_hs,
                                                                                                                                                                            "MASTER HS": hs_code,
                                                                                                                                                                            "HS6_MATCH": "O" if inv_hs[:6] == hs_code[:6] else "X",
                                                                                                                                                                            "HS10_MATCH": "O" if inv_hs[:10] == hs_code[:10] else "X",
                                                                                                                                                                            "Part Description": desc,
                                                                                                                                                                            "모델명": row.iloc[0].get("모델명", ""),
                                                                                                                                                                            "전파인증번호": row.iloc[0].get("전파인증번호", ""),
                                                                                                                                                                            "전기인증번호": row.iloc[0].get("전기인증번호", ""),
                                                                                                                                                                            "기관": row.iloc[0].get("기관", ""),
                                                                                                                                                                            "정격전압": row.iloc[0].get("정격전압", ""),
                                                                                                                                                                            "요건비대상": row.iloc[0].get("요건비대상", ""),
                                                                                                                                                                            "REMARK": row.iloc[0].get("REMARK", "")
                                                                                                                                                                            }
                                                                                                                                                                            else:
                                                                                                                                                                                result = {
                                                                                                                                                                                "Microsoft Part No.": part_no,
                                                                                                                                                                                "INV HS": inv_hs,
                                                                                                                                                                                "MASTER HS": "N/A",
                                                                                                                                                                                "HS6_MATCH": "X",
                                                                                                                                                                                "HS10_MATCH": "X",
                                                                                                                                                                                "Part Description": "⚠️ 마스터에 없음"
                                                                                                                                                                                }

                                                                                                                                                                                st.session_state.compare_results.append(result)

                                                                                                                                                                                if st.session_state.get("compare_results"):
                                                                                                                                                                                    st.dataframe(pd.DataFrame(st.session_state.compare_results))
                                                                                                                                                                                    if master_df is None:
                                                                                                                                                                                        st.warning("⚠️ 마스터 파일이 없습니다. 최초 1회 업로드가 필요합니다.")




                                                                                                                                                                                        with tab4:
                                                                                                                                                                                            #st.header("📕 MS1279-WESCO 인보이스 추출 (Item No + Description 매칭)")
                                                                                                                                                                                            uploaded_file = st.file_uploader("WESCO 인보이스 PDF 업로드", type=["pdf"], key="wesco_bbox_descmerge")
                                                                                                                                                                                            if uploaded_file and "master_df" in st.session_state:
                                                                                                                                                                                                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                                                                                                                                                                                                    tmp_file.write(uploaded_file.read())
                                                                                                                                                                                                    temp_pdf_path = tmp_file.name

                                                                                                                                                                                                    import collections
                                                                                                                                                                                                    import re

                                                                                                                                                                                                    def group_words_by_line(words, y_tolerance=3):
                                                                                                                                                                                                        lines = collections.defaultdict(list)
                                                                                                                                                                                                        for word in words:
                                                                                                                                                                                                            y_center = (word["top"] + word["bottom"]) / 2
                                                                                                                                                                                                            matched = False
                                                                                                                                                                                                            for key in lines:
                                                                                                                                                                                                                if abs(key - y_center) <= y_tolerance:
                                                                                                                                                                                                                    lines[key].append(word)
                                                                                                                                                                                                                    matched = True
                                                                                                                                                                                                                    break
                                                                                                                                                                                                                if not matched:
                                                                                                                                                                                                                    lines[y_center].append(word)
                                                                                                                                                                                                                    return lines

                                                                                                                                                                                                                def clean_code(text):
                                                                                                                                                                                                                    return re.sub(r"[^A-Za-z0-9]", "", str(text)).upper()

                                                                                                                                                                                                                extracted_rows = []

                                                                                                                                                                                                                with pdfplumber.open(temp_pdf_path) as pdf:
                                                                                                                                                                                                                    for page in pdf.pages:
                                                                                                                                                                                                                        words = page.extract_words(use_text_flow=True, keep_blank_chars=True)
                                                                                                                                                                                                                        lines = group_words_by_line(words)

                                                                                                                                                                                                                        for _, line_words in sorted(lines.items()):
                                                                                                                                                                                                                            line_words.sort(key=lambda w: w["x0"])
                                                                                                                                                                                                                            text_line = [w["text"] for w in line_words]
                                                                                                                                                                                                                            digit_count = sum(1 for t in text_line if any(c.isdigit() for c in t))
                                                                                                                                                                                                                            if digit_count >= 2 and len(text_line) >= 6:
                                                                                                                                                                                                                                extracted_rows.append(text_line)

                                                                                                                                                                                                                                os.remove(temp_pdf_path)

                                                                                                                                                                                                                                if extracted_rows:
                                                                                                                                                                                                                                    headers = [
                                                                                                                                                                                                                                    "Item Number", "Description", "Ordered Qty",
                                                                                                                                                                                                                                    "Shipped Qty", "UM", "Unit Price", "Per", "Amount"
                                                                                                                                                                                                                                    ]
                                                                                                                                                                                                                                    norm_rows = [row + [""] * (8 - len(row)) for row in extracted_rows if len(row) <= 8]
                                                                                                                                                                                                                                    wesco_df = pd.DataFrame(norm_rows, columns=headers)

                                                                                                                                                                                                                                    # 정제
                                                                                                                                                                                                                                    wesco_df["clean_item"] = wesco_df["Item Number"].apply(clean_code)
                                                                                                                                                                                                                                    wesco_df["clean_desc"] = wesco_df["Description"].apply(clean_code)

                                                                                                                                                                                                                                    master_df = st.session_state["master_df"].copy()
                                                                                                                                                                                                                                    master_df["clean_code"] = master_df["Microsoft Part No."].apply(clean_code)
                                                                                                                                                                                                                                    master_df["clean_desc"] = master_df["Part Description"].apply(clean_code)

                                                                                                                                                                                                                                    # 병합 (1차: item 기준)
                                                                                                                                                                                                                                    columns_to_pull = [
                                                                                                                                                                                                                                    "Microsoft Part No.", "Part Description", "HS Code", "요건비대상", "clean_code", "clean_desc"
                                                                                                                                                                                                                                    ]
                                                                                                                                                                                                                                    merged_by_item = wesco_df.merge(master_df[columns_to_pull], left_on="clean_item", right_on="clean_code", how="left", suffixes=("", "_item"))

                                                                                                                                                                                                                                    # 병합 (2차: desc 기준)
                                                                                                                                                                                                                                    merged_by_desc = wesco_df.merge(master_df[columns_to_pull], left_on="clean_desc", right_on="clean_desc", how="left", suffixes=("", "_desc"))

                                                                                                                                                                                                                                    # 보완: item 병합 우선, 없으면 desc 병합 결과 채움
                                                                                                                                                                                                                                    final = merged_by_item.copy()
                                                                                                                                                                                                                                    for col in ["Microsoft Part No.", "Part Description", "HS Code", "요건비대상"]:
                                                                                                                                                                                                                                        final[col] = final[col].fillna(merged_by_desc[col])

                                                                                                                                                                                                                                        # 누락된 항목 처리
                                                                                                                                                                                                                                        final["Microsoft Part No."] = final["Microsoft Part No."].fillna("신규코드")


                                                                                                                                                                                                                                        # 원산지 추출
                                                                                                                                                                                                                                        coo_lines = [w["text"] for w in words if "COO:" in w["text"] or "Origin:" in w["text"]]
                                                                                                                                                                                                                                        origin_map = {}
                                                                                                                                                                                                                                        current_origin = ""
                                                                                                                                                                                                                                        for line in coo_lines:
                                                                                                                                                                                                                                            if "COO:" in line:
                                                                                                                                                                                                                                                match = re.search(r"COO:\s*(\S+)", line)
                                                                                                                                                                                                                                                if match:
                                                                                                                                                                                                                                                    current_origin = match.group(1)
                                                                                                                                                                                                                                                    elif "Origin:" in line and not current_origin:
                                                                                                                                                                                                                                                        match = re.search(r"Origin:\s*(\S+)", line)
                                                                                                                                                                                                                                                        if match:
                                                                                                                                                                                                                                                            current_origin = match.group(1)
                                                                                                                                                                                                                                                            final["Country of Origin"] = current_origin

                                                                                                                                                                                                                                                            final["Part Description"] = final["Part Description"].fillna(final["Description"])
                                                                                                                                                                                                                                                            # 줄 단위 origin 추출 (끝까지 탐색)
                                                                                                                                                                                                                                                            origin_map = {}
                                                                                                                                                                                                                                                            item_list = wesco_df["Item Number"].dropna().unique().tolist()

                                                                                                                                                                                                                                                            lines_by_page = []  # ← 기존 페이지 처리 중 수집된 라인들 사용

                                                                                                                                                                                                                                                            for page in pdf.pages:
                                                                                                                                                                                                                                                                lines_by_page.extend(page.extract_text().split("\n"))

                                                                                                                                                                                                                                                                for idx, line in enumerate(lines_by_page):
                                                                                                                                                                                                                                                                    for item in item_list:
                                                                                                                                                                                                                                                                        if item.strip() in line:
                                                                                                                                                                                                                                                                            # 이 아이템 아래 모든 줄에서 origin 찾기
                                                                                                                                                                                                                                                                            origin_val = "미확인"
                                                                                                                                                                                                                                                                            for next_line in lines_by_page[idx:]:  # ← 끝까지 검색
                                                                                                                                                                                                                                                                            match = re.search(r"(?:COO|Origin):\s*(\S+)", next_line)
                                                                                                                                                                                                                                                                            if match:
                                                                                                                                                                                                                                                                                origin_val = match.group(1)
                                                                                                                                                                                                                                                                                break
                                                                                                                                                                                                                                                                            origin_map[item] = origin_val



                                                                                                                                                                                                                                                                            def extract_export_code_map(lines_by_page, item_list):
                                                                                                                                                                                                                                                                                """
                                                                                                                                                                                                                                                                                각 Item 기준으로 Export Code or HS Code 값을 찾아주는 함수
                                                                                                                                                                                                                                                                                """

                                                                                                                                                                                                                                                                                # origin_map을 final에 적용
                                                                                                                                                                                                                                                                                final["Country of Origin"] = final["Item Number"].map(origin_map).fillna("미확인")

                                                                                                                                                                                                                                                                                # 원산지 2자리 코드로 변환
                                                                                                                                                                                                                                                                                origin_abbrev = {
                                                                                                                                                                                                                                                                                "china": "CN",
                                                                                                                                                                                                                                                                                "cn": "CN",
                                                                                                                                                                                                                                                                                "vietnam": "VN",
                                                                                                                                                                                                                                                                                "vn": "VN",
                                                                                                                                                                                                                                                                                "korea": "KR",
                                                                                                                                                                                                                                                                                "kr": "KR",
                                                                                                                                                                                                                                                                                "taiwan": "TW",
                                                                                                                                                                                                                                                                                "tw": "TW",
                                                                                                                                                                                                                                                                                "thailand": "TH",
                                                                                                                                                                                                                                                                                "th": "TH"
                                                                                                                                                                                                                                                                                }
                                                                                                                                                                                                                                                                                final["Country of Origin"] = final["Country of Origin"].str.lower().map(origin_abbrev).fillna("미확인")

                                                                                                                                                                                                                                                                                st.dataframe(final[[
                                                                                                                                                                                                                                                                                "Item Number", "Microsoft Part No.", "Part Description",
                                                                                                                                                                                                                                                                                "Ordered Qty", "Shipped Qty", "UM", "Unit Price", "Amount",
                                                                                                                                                                                                                                                                                "HS Code", "요건비대상"
                                                                                                                                                                                                                                                                                ]])
                                                                                                                                                                                                                                                                                # 최종 저장 열 명시적으로 지정 → clean_ 열 완전 제외
                                                                                                                                                                                                                                                                                columns_to_export = [
                                                                                                                                                                                                                                                                                "Item Number", "Microsoft Part No.", "Part Description",
                                                                                                                                                                                                                                                                                "Ordered Qty", "Shipped Qty", "UM", "Unit Price", "Amount",
                                                                                                                                                                                                                                                                                "HS Code", "요건비대상", "Country of Origin", "Export Code"
                                                                                                                                                                                                                                                                                ]
                                                                                                                                                                                                                                                                                final_to_export = final[columns_to_export]

                                                                                                                                                                                                                                                                                # 엑셀 저장 - Export Code 포함 확인
                                                                                                                                                                                                                                                                                print("Export Code 열 포함 여부:", "Export Code" in final_to_export.columns)

                                                                                                                                                                                                                                                                                excel_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                                                                                                                                                                                                                                                                                with pd.ExcelWriter(excel_file.name, engine="openpyxl") as writer:
                                                                                                                                                                                                                                                                                    final_to_export.to_excel(writer, index=False, sheet_name="WESCO_MERGED")

                                                                                                                                                                                                                                                                                    invoice_sheet = pd.DataFrame({
                                                                                                                                                                                                                                                                                    "HS Code": final["HS Code"],
                                                                                                                                                                                                                                                                                    "Part Description": final["Part Description"] + ' ORIGIN:' + final["Country of Origin"],
                                                                                                                                                                                                                                                                                    "Microsoft Part No.": "ITEM NO: " + final["Microsoft Part No."],
                                                                                                                                                                                                                                                                                    "수량": final["Shipped Qty"],
                                                                                                                                                                                                                                                                                    "단위": final["UM"],
                                                                                                                                                                                                                                                                                    "단가": final["Unit Price"],
                                                                                                                                                                                                                                                                                    "금액": final["Amount"],
                                                                                                                                                                                                                                                                                    "Microsoft Part No. (원본)": final["Microsoft Part No."]
                                                                                                                                                                                                                                                                                    })

                                                                                                                                                                                                                                                                                    invoice_sheet.to_excel(writer, index=False, sheet_name="신고서")

                                                                                                                                                                                                                                                                                    with open(excel_file.name, "rb") as f:
                                                                                                                                                                                                                                                                                        st.download_button(
                                                                                                                                                                                                                                                                                        label="엑셀 다운로드",
                                                                                                                                                                                                                                                                                        data=f,
                                                                                                                                                                                                                                                                                        file_name="wesco_invoice.xlsx"
                                                                                                                                                                                                                                                                                        )
                                                                                                                                                                                                                                                                                        else:
                                                                                                                                                                                                                                                                                            st.warning("유효한 데이터를 추출할 수 없습니다.")
                                                                                                                                                                                                                                                                                            elif "master_df" not in st.session_state:
                                                                                                                                                                                                                                                                                                st.warning("MASTER_MS5673.xlsx 파일이 로드되지 않았습니다. 먼저 마스터 파일을 탭3에서 업로드하세요.")