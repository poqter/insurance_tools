import hashlib
import re
from datetime import datetime
from io import BytesIO

import openpyxl
import streamlit as st
from openpyxl.utils import get_column_letter


def make_input_signature(
    main_bytes: bytes,
    template_bytes: bytes,
    template_mode: str,
    start_row: int,
    end_row: int,
) -> str:
    digest = hashlib.sha256()
    digest.update(main_bytes)
    digest.update(template_bytes)
    digest.update(f"{template_mode}:{start_row}:{end_row}".encode("utf-8"))
    return digest.hexdigest()


def build_analysis_file(
    main_bytes: bytes,
    template_bytes: bytes,
    start_row: int,
    end_row: int,
) -> tuple[bytes, str, str]:
    """기존 엑셀 복사·가공 로직을 실행하고 결과 바이트, 파일명, 고객명을 반환합니다."""
    main_wb = openpyxl.load_workbook(BytesIO(main_bytes), data_only=True)

    required_sheets = ["계약사항", "상품별보장내용"]
    missing_sheets = [sheet for sheet in required_sheets if sheet not in main_wb.sheetnames]
    if missing_sheets:
        raise ValueError("필수 시트 없음:" + ",".join(missing_sheets))

    try:
        print_wb = openpyxl.load_workbook(BytesIO(template_bytes))
    except Exception as exc:
        raise ValueError("결과 양식 파일을 열 수 없습니다.") from exc

    main_ws1 = main_wb["계약사항"]
    main_ws2 = main_wb["상품별보장내용"]
    print_ws = print_wb.active

    for idx in range(27):
        print_ws.cell(row=10, column=4 + idx).value = main_ws1[f"J{9 + idx}"].value

    for row_offset, col in enumerate(["K", "L"]):
        for idx in range(27):
            print_ws.cell(row=8 + row_offset, column=4 + idx).value = main_ws1[f"{col}{9 + idx}"].value

    for col in range(6, 30):
        raw_value = main_ws2.cell(row=7, column=col).value
        if raw_value is not None:
            number = re.sub(r"[^\d]", "", str(raw_value))
            print_ws.cell(row=7, column=col - 2).value = int(number) if number else ""

    for row in range(2, 7):
        for col in range(6, 30):
            print_ws.cell(row=row, column=col - 2).value = main_ws2.cell(row=row, column=col).value

    for row in range(start_row, end_row + 1):
        for col in range(6, 30):
            print_ws.cell(row=row + 3, column=col - 2).value = main_ws2.cell(row=row, column=col).value

    name_prefix = str(main_ws1["B2"].value or "고객")[:3]
    detail_text = main_ws1["D2"].value or ""
    print_ws["A1"] = f"{name_prefix}님의 기존 보험 보장 분석 {detail_text}"

    def get_real_last_row(ws):
        for row in range(ws.max_row, 0, -1):
            if any(cell.value not in [None, ""] for cell in ws[row]):
                return row
        return 1

    def get_real_last_col(ws):
        for col in range(ws.max_column, 0, -1):
            col_letter = get_column_letter(col)
            if any(
                ws[f"{col_letter}{row}"].value not in [None, ""]
                for row in range(1, ws.max_row + 1)
            ):
                return col
        return 1

    real_last_row = get_real_last_row(print_ws)
    real_last_col = get_real_last_col(print_ws)
    last_col_letter = get_column_letter(real_last_col)
    print_ws.print_area = f"A1:{last_col_letter}{real_last_row}"

    today_str = datetime.today().strftime("%Y%m%d")
    filename = f"{name_prefix}님의_보장분석엑셀_{today_str}.xlsx"
    output_excel = BytesIO()
    print_wb.save(output_excel)
    output_excel.seek(0)
    return output_excel.getvalue(), filename, name_prefix


def run():
    st.title("📊 보장 분석 도우미")
    st.caption("보험사 보장분석 자료를 고객용 엑셀 양식으로 변환합니다.")

    with st.expander("사용 방법 안내"):
        st.markdown(
            """
            1. 한화라이프랩에서 내려받은 **컨설팅보장분석.xlsx** 파일을 업로드합니다.
            2. 기본 양식 또는 개인 양식을 선택합니다.
            3. **보장 분석 시작**을 누른 뒤 결과 파일을 다운로드합니다.

            - 지원 파일: `.xlsx`
            - 개인 양식을 사용할 때만 복사 행 범위를 변경할 수 있습니다.
            """
        )
        st.caption("버전 v1.3.0 · 제작 박병선 팀장")

    try:
        with open("print.xlsx", "rb") as file:
            default_template_data = file.read()
        default_template_error = None
    except Exception as exc:
        default_template_data = b""
        default_template_error = str(exc)

    st.markdown("### 1. 원본 보장분석 파일")
    uploaded_main = st.file_uploader(
        "컨설팅보장분석.xlsx 파일을 업로드하세요",
        type=["xlsx"],
        key="analyzer_main_file",
    )

    st.markdown("### 2. 결과 양식")
    template_mode = st.radio(
        "사용할 결과 양식을 선택하세요",
        ["기본 양식 사용", "개인 양식 사용"],
        horizontal=True,
        key="analyzer_template_mode",
    )

    st.caption("기본 양식을 수정해 개인 양식으로 사용할 수 있습니다.")
    st.download_button(
        "기본 양식 다운로드",
        data=default_template_data,
        file_name="print.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=not bool(default_template_data),
        use_container_width=True,
    )

    uploaded_template = None
    start_row = 9
    end_row = 45

    if template_mode == "기본 양식 사용":
        if default_template_error:
            st.error("기본 양식 파일을 찾을 수 없습니다. 프로젝트의 print.xlsx 파일을 확인해 주세요.")
        else:
            st.info("기본 내장 양식을 사용합니다.")
    else:
        uploaded_template = st.file_uploader(
            "개인용 보장분석 양식을 업로드하세요",
            type=["xlsx"],
            key="analyzer_template_file",
        )
        with st.expander("개인 양식 고급 설정"):
            row_col1, row_col2 = st.columns(2)
            with row_col1:
                start_row = int(
                    st.number_input(
                        "복사 시작 행",
                        min_value=1,
                        max_value=100,
                        value=9,
                        step=1,
                        key="analyzer_start_row",
                    )
                )
            with row_col2:
                end_row = int(
                    st.number_input(
                        "복사 종료 행",
                        min_value=1,
                        max_value=100,
                        value=45,
                        step=1,
                        key="analyzer_end_row",
                    )
                )

    row_range_valid = end_row > start_row
    if not row_range_valid:
        st.error("복사 종료 행은 시작 행보다 크게 입력해 주세요.")

    main_bytes = uploaded_main.getvalue() if uploaded_main else b""
    if template_mode == "기본 양식 사용":
        template_bytes = default_template_data
        template_ready = bool(default_template_data)
    else:
        template_bytes = uploaded_template.getvalue() if uploaded_template else b""
        template_ready = uploaded_template is not None

    ready_to_analyze = bool(uploaded_main) and template_ready and row_range_valid
    current_signature = (
        make_input_signature(main_bytes, template_bytes, template_mode, start_row, end_row)
        if ready_to_analyze
        else None
    )

    st.markdown("### 3. 분석 실행")
    if st.button(
        "보장 분석 시작",
        type="primary",
        disabled=not ready_to_analyze,
        use_container_width=True,
        key="analyzer_run",
    ):
        st.session_state.pop("analyzer_result", None)
        st.session_state.pop("analyzer_error", None)
        try:
            with st.spinner("보장분석 파일을 처리하고 있습니다..."):
                result_bytes, filename, customer_name = build_analysis_file(
                    main_bytes,
                    template_bytes,
                    start_row,
                    end_row,
                )
            st.session_state["analyzer_result"] = {
                "signature": current_signature,
                "bytes": result_bytes,
                "filename": filename,
                "customer_name": customer_name,
                "template_name": template_mode.replace(" 사용", ""),
            }
        except ValueError as exc:
            message = str(exc)
            if message.startswith("필수 시트 없음:"):
                missing = message.split(":", 1)[1].replace(",", ", ")
                friendly = f"업로드한 원본 파일에서 필수 시트({missing})를 찾을 수 없습니다."
            else:
                friendly = message
            st.session_state["analyzer_error"] = {
                "signature": current_signature,
                "message": friendly,
                "detail": repr(exc),
            }
        except Exception as exc:
            st.session_state["analyzer_error"] = {
                "signature": current_signature,
                "message": "파일 처리 중 문제가 발생했습니다. 원본 파일과 결과 양식을 다시 확인해 주세요.",
                "detail": repr(exc),
            }

    error = st.session_state.get("analyzer_error")
    if error and error.get("signature") == current_signature:
        st.error(error["message"])
        with st.expander("오류 상세 보기"):
            st.code(error["detail"])

    result = st.session_state.get("analyzer_result")
    if result and result.get("signature") == current_signature:
        st.divider()
        st.success("보장 분석이 완료되었습니다.")
        result_col1, result_col2 = st.columns(2)
        with result_col1:
            st.metric("고객명", result["customer_name"])
        with result_col2:
            st.metric("적용 양식", result["template_name"])
        st.caption(f"결과 파일: {result['filename']}")
        st.download_button(
            "결과 엑셀 다운로드",
            data=result["bytes"],
            file_name=result["filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
            key="analyzer_download_result",
        )


if __name__ == "__main__":
    run()
