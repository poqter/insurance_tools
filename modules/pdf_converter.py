import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO


def clean_dataframe(df: pd.DataFrame, remove_empty_rows=True, remove_empty_cols=True):
    """
    PDF에서 추출된 표 데이터를 기본 정리하는 함수
    """
    df = df.copy()

    # 모든 값을 문자열 기준으로 정리
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)

    # 완전히 빈 문자열을 결측값으로 변환
    df = df.replace("", pd.NA)

    if remove_empty_rows:
        df = df.dropna(how="all")

    if remove_empty_cols:
        df = df.dropna(axis=1, how="all")

    return df


def tables_to_excel(all_tables, merged_df=None, save_mode="merged_and_each"):
    """
    추출된 표 목록을 엑셀 파일로 변환
    """
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if save_mode in ["merged_only", "merged_and_each"] and merged_df is not None:
            merged_df.to_excel(writer, index=False, sheet_name="전체표합치기")

        if save_mode in ["each_only", "merged_and_each"]:
            for i, item in enumerate(all_tables, start=1):
                df = item["df"]
                page_num = item["page"]
                table_num = item["table"]

                sheet_name = f"p{page_num}_table{table_num}"
                df.to_excel(writer, index=False, sheet_name=sheet_name[:31])

    output.seek(0)
    return output


def run():
    st.title("📄 PDF 표 엑셀 변환기")
    st.caption("PDF 파일 안의 표를 자동으로 추출하여 하나의 엑셀 파일로 변환합니다.")

    st.divider()

    uploaded_files = st.file_uploader(
        "PDF 파일을 업로드하세요",
        type=["pdf"],
        accept_multiple_files=True
    )

    st.subheader("⚙️ 변환 옵션")

    col1, col2, col3 = st.columns(3)

    with col1:
        add_file_name = st.checkbox("파일명 열 추가", value=True)
        add_page_info = st.checkbox("페이지/표번호 열 추가", value=True)

    with col2:
        remove_empty_rows = st.checkbox("빈 행 제거", value=True)
        remove_empty_cols = st.checkbox("빈 열 제거", value=True)

    with col3:
        first_row_as_header = st.checkbox("첫 행을 제목행으로 사용", value=False)
        show_each_table = st.checkbox("표별 미리보기 표시", value=False)

    save_mode_label = st.radio(
        "엑셀 저장 방식",
        [
            "전체 합친 시트 + 표별 개별 시트",
            "전체 합친 시트만",
            "표별 개별 시트만",
        ],
        horizontal=True
    )

    save_mode_map = {
        "전체 합친 시트 + 표별 개별 시트": "merged_and_each",
        "전체 합친 시트만": "merged_only",
        "표별 개별 시트만": "each_only",
    }

    save_mode = save_mode_map[save_mode_label]

    st.divider()

    if not uploaded_files:
        st.info("PDF 파일을 업로드하면 변환을 시작할 수 있습니다.")
        return

    if st.button("🚀 PDF 표 추출 시작", type="primary"):
        all_tables = []

        with st.spinner("PDF에서 표를 추출하는 중입니다..."):
            for uploaded_file in uploaded_files:
                file_name = uploaded_file.name

                try:
                    with pdfplumber.open(uploaded_file) as pdf:
                        for page_num, page in enumerate(pdf.pages, start=1):
                            tables = page.extract_tables()

                            for table_num, table in enumerate(tables, start=1):
                                if not table:
                                    continue

                                df = pd.DataFrame(table)

                                df = clean_dataframe(
                                    df,
                                    remove_empty_rows=remove_empty_rows,
                                    remove_empty_cols=remove_empty_cols
                                )

                                if df.empty:
                                    continue

                                # 첫 행을 컬럼명으로 사용
                                if first_row_as_header and len(df) > 1:
                                    df.columns = df.iloc[0]
                                    df = df.iloc[1:].reset_index(drop=True)

                                # 출처 정보 추가
                                if add_page_info:
                                    df.insert(0, "표번호", table_num)
                                    df.insert(0, "페이지", page_num)

                                if add_file_name:
                                    df.insert(0, "파일명", file_name)

                                all_tables.append(
                                    {
                                        "file": file_name,
                                        "page": page_num,
                                        "table": table_num,
                                        "df": df
                                    }
                                )

                except Exception as e:
                    st.error(f"'{file_name}' 처리 중 오류가 발생했습니다: {e}")

        if not all_tables:
            st.warning(
                "추출된 표가 없습니다. PDF가 스캔본 이미지이거나 표 구조가 복잡한 파일일 수 있습니다."
            )
            return

        merged_df = pd.concat(
            [item["df"] for item in all_tables],
            ignore_index=True
        )

        st.success(f"총 {len(all_tables)}개의 표를 추출했습니다.")
        st.write(f"전체 병합 데이터: **{len(merged_df):,}행 × {len(merged_df.columns):,}열**")

        st.subheader("👀 전체 병합 미리보기")
        st.dataframe(merged_df, use_container_width=True)

        if show_each_table:
            st.subheader("📌 표별 미리보기")

            for i, item in enumerate(all_tables, start=1):
                with st.expander(
                    f"{i}. {item['file']} / {item['page']}페이지 / 표 {item['table']}"
                ):
                    st.dataframe(item["df"], use_container_width=True)

        excel_file = tables_to_excel(
            all_tables=all_tables,
            merged_df=merged_df,
            save_mode=save_mode
        )

        st.download_button(
            label="📥 엑셀 파일 다운로드",
            data=excel_file,
            file_name="pdf_table_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )