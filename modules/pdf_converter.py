import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


# -----------------------------
# 엑셀 저장 불가 문자 제거
# -----------------------------
def clean_excel_text(value):
    """
    엑셀 저장 시 오류를 일으키는 숨은 제어문자를 제거합니다.
    PDF에서 추출된 데이터에는 눈에 보이지 않는 특수문자가 섞일 수 있습니다.
    """
    if isinstance(value, str):
        value = ILLEGAL_CHARACTERS_RE.sub("", value)
        value = value.strip()
    return value


# -----------------------------
# DataFrame 정리 함수
# -----------------------------
def clean_dataframe(df: pd.DataFrame, remove_empty_rows=True, remove_empty_cols=True):
    """
    PDF에서 추출된 표 데이터를 기본 정리하는 함수입니다.

    처리 내용:
    1. 엑셀 저장 불가 문자 제거
    2. 문자열 앞뒤 공백 제거
    3. 빈 문자열을 결측값으로 변환
    4. 빈 행 제거
    5. 빈 열 제거
    """
    df = df.copy()

    # 모든 셀의 엑셀 금지 문자 제거
    df = df.map(clean_excel_text)

    # 컬럼명에도 금지 문자가 들어갈 수 있으므로 정리
    df.columns = [
        clean_excel_text(col) if isinstance(col, str) else col
        for col in df.columns
    ]

    # 완전히 빈 문자열을 결측값으로 변환
    df = df.replace("", pd.NA)

    if remove_empty_rows:
        df = df.dropna(how="all")

    if remove_empty_cols:
        df = df.dropna(axis=1, how="all")

    return df


# -----------------------------
# 안전한 시트명 생성
# -----------------------------
def safe_sheet_name(name: str):
    """
    엑셀 시트명에서 사용할 수 없는 문자를 제거하고,
    최대 31자 제한에 맞춥니다.
    """
    invalid_chars = ["\\", "/", "*", "?", ":", "[", "]"]

    for char in invalid_chars:
        name = name.replace(char, "_")

    name = name.strip()

    if not name:
        name = "Sheet"

    return name[:31]


# -----------------------------
# 추출된 표를 엑셀 파일로 변환
# -----------------------------
def tables_to_excel(all_tables, merged_df=None, save_mode="merged_and_each"):
    """
    추출된 표 목록을 엑셀 파일로 변환합니다.

    save_mode:
    - merged_and_each: 전체 합친 시트 + 표별 개별 시트
    - merged_only: 전체 합친 시트만
    - each_only: 표별 개별 시트만
    """
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 전체 병합 시트 저장
        if save_mode in ["merged_only", "merged_and_each"] and merged_df is not None:
            merged_df = clean_dataframe(merged_df)
            merged_df.to_excel(
                writer,
                index=False,
                sheet_name=safe_sheet_name("전체표합치기")
            )

        # 표별 개별 시트 저장
        if save_mode in ["each_only", "merged_and_each"]:
            used_sheet_names = set()

            for i, item in enumerate(all_tables, start=1):
                df = clean_dataframe(item["df"])

                page_num = item["page"]
                table_num = item["table"]

                base_sheet_name = safe_sheet_name(f"p{page_num}_table{table_num}")

                # 같은 이름의 시트가 중복되지 않도록 처리
                sheet_name = base_sheet_name
                count = 1

                while sheet_name in used_sheet_names:
                    suffix = f"_{count}"
                    sheet_name = safe_sheet_name(base_sheet_name[:31 - len(suffix)] + suffix)
                    count += 1

                used_sheet_names.add(sheet_name)

                df.to_excel(
                    writer,
                    index=False,
                    sheet_name=sheet_name
                )

    output.seek(0)
    return output


# -----------------------------
# Streamlit 실행 함수
# -----------------------------
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
        add_file_name = st.checkbox("파일명 열 추가", value=False)
        add_page_info = st.checkbox("페이지/표번호 열 추가", value=False)

    with col2:
        remove_empty_rows = st.checkbox("빈 행 제거", value=False)
        remove_empty_cols = st.checkbox("빈 열 제거", value=False)

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
                                    header_row = [
                                        clean_excel_text(col)
                                        for col in df.iloc[0].tolist()
                                    ]

                                    df.columns = header_row
                                    df = df.iloc[1:].reset_index(drop=True)

                                    # 제목행 변경 이후 다시 한 번 정리
                                    df = clean_dataframe(
                                        df,
                                        remove_empty_rows=remove_empty_rows,
                                        remove_empty_cols=remove_empty_cols
                                    )

                                # 출처 정보 추가
                                if add_page_info:
                                    df.insert(0, "표번호", table_num)
                                    df.insert(0, "페이지", page_num)

                                if add_file_name:
                                    df.insert(0, "파일명", clean_excel_text(file_name))

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

        # 전체 표 병합
        try:
            merged_df = pd.concat(
                [item["df"] for item in all_tables],
                ignore_index=True
            )

            # 병합 후 최종 정리
            merged_df = clean_dataframe(
                merged_df,
                remove_empty_rows=remove_empty_rows,
                remove_empty_cols=remove_empty_cols
            )

        except Exception as e:
            st.error(f"표 병합 중 오류가 발생했습니다: {e}")
            return

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

        try:
            excel_file = tables_to_excel(
                all_tables=all_tables,
                merged_df=merged_df,
                save_mode=save_mode
            )

        except Exception as e:
            st.error(f"엑셀 파일 생성 중 오류가 발생했습니다: {e}")
            return

        st.download_button(
            label="📥 엑셀 파일 다운로드",
            data=excel_file,
            file_name="pdf_table_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )