
import streamlit as st
import pandas as pd
import tempfile
import os

st.set_page_config(page_title="PDF 항목 추출기 + 마스터 비교", layout="wide")
st.title("📄 PDF → Excel 항목 추출기 + 마스터 데이터 비교")

tab1, tab2, tab3 = st.tabs(["📘 MS1056", "📗 MS1279-PAYMENTS", "📒 마스터 비교"])

with tab3:
    st.header("📒 마스터 데이터 비교")

    uploaded_excel = st.file_uploader("📥 비교 대상 엑셀 업로드 (Microsoft Part No., 원산지, 수량, 단위, 단가, 금액, INV HS 포함)", type=["xlsx"], key="compare_excel")
    master_file = st.file_uploader("📘 마스터 파일 업로드 (전체 데이터 포함)", type=["xlsx"], key="master_excel")

    if uploaded_excel and master_file:
        input_df = pd.read_excel(uploaded_excel)
        master_df = pd.read_excel(master_file)

        # 필요한 컬럼만 추출하고 이름 통일
        master_df = master_df.rename(columns=lambda x: x.strip())
        input_df = input_df.rename(columns=lambda x: x.strip())

        # 병합
        merged = input_df.merge(master_df, how="left", on="Microsoft Part No.")

        # HS CODE 비교
        merged["HS10_MATCH"] = merged.apply(
            lambda row: "O" if str(row.get("INV HS", "")).replace("-", "")[:10] == str(row.get("HS CODE", "")).replace("-", "")[:10] else "X", axis=1
        )
        merged["HS6_MATCH"] = merged.apply(
            lambda row: "O" if str(row.get("INV HS", "")).replace("-", "")[:6] == str(row.get("HS CODE", "")).replace("-", "")[:6] else "X", axis=1
        )

        # 원하는 컬럼 순서대로 정리
        columns_to_show = [
            "Microsoft Part No.", "원산지", "수량", "단위", "단가", "금액", "INV HS",
            "Part Description", "HS CODE", "모델명", "전파인증번호", "전기인증번호", "기관", "정격전압", "요건비대상사유", "REMARK",
            "HS10_MATCH", "HS6_MATCH"
        ]
        final_df = merged[[col for col in columns_to_show if col in merged.columns]]

        st.subheader("🔍 비교 결과 미리보기")
        st.dataframe(final_df)

        to_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        final_df.to_excel(to_excel.name, index=False)

        with open(to_excel.name, "rb") as f:
            st.download_button(
                label="📥 비교 결과 엑셀 다운로드",
                data=f,
                file_name="master_compare_result.xlsx"
            )
