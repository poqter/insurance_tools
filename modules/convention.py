import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
import os

# ✅ 한 줄로 제어: True면 썸머 기준 노출/계산 포함
SHOW_SUMMER = True

def run():
    st.set_page_config(page_title="보험 계약 환산기", layout="wide")
    st.title("📊 보험 계약 실적 환산기 (컨벤션{} 기준)".format(" & 썸머" if SHOW_SUMMER else ""))

    # 👉 사이드바 안내
    with st.sidebar:
        st.header("🧭 사용 방법")
        st.markdown(
            """
            **🖥️ 한화라이프랩 전산**  
            **- 📂 계약관리**  
            **- 📑 보유계약 장기**  
            **- ⏱️ 기간 설정**  
            **- 💾 엑셀 다운로드 후 파일 첨부하면 됩니다.**
            """
        )

    uploaded_file = st.file_uploader("📂 계약 목록 Excel 파일 업로드 (.xlsx)", type=["xlsx"])

    if not uploaded_file:
        st.info("📤 계약 목록 Excel 파일(.xlsx)을 업로드해주세요.")
        return

    base_filename = os.path.splitext(uploaded_file.name)[0]
    download_filename = f"{base_filename}_환산결과.xlsx"

    # 1) 필요한 컬럼만 로드 (✅ 수금자명 추가)
    columns_needed = ["수금자명", "계약일", "보험사", "상품명", "납입기간", "초회보험료", "쉐어율", "납입방법", "상품군2", "계약상태"]
    df = pd.read_excel(uploaded_file, usecols=columns_needed)

    # 제외목록 안전 초기화
    excluded_df = pd.DataFrame()

    # '일시납 / 연금성·저축성 / 철회·해약' 제외
    if {"납입방법", "상품군2", "계약상태"}.issubset(df.columns):
        before_count = len(df)

        df["납입방법"] = df["납입방법"].astype(str).str.strip()
        df["상품군2"] = df["상품군2"].astype(str).str.strip()
        df["계약상태"] = df["계약상태"].astype(str).str.strip()

        is_lumpsum = df["납입방법"].str.contains("일시납")
        is_savings = df["상품군2"].str.contains("연금성|저축성")
        is_cancelled = df["계약상태"].str.contains("철회|해약")

        is_excluded = is_lumpsum | is_savings | is_cancelled
        excluded_df = df[is_excluded].copy()
        df = df[~is_excluded].copy()

        excluded_count = before_count - len(df)
        if excluded_count > 0:
            st.warning(f"⚠️ 제외된 계약 {excluded_count}건 (일시납 / 연금성·저축성 / 철회|해약 계약)이 계산에서 제외되었습니다.")

    # 2) 내부 컬럼명 정규화
    df.rename(columns={"계약일": "계약일자", "초회보험료": "보험료"}, inplace=True)

    # 3) 필수 항목 체크 (✅ 수금자명 포함)
    required_columns = {"수금자명", "계약일자", "보험사", "상품명", "납입기간", "보험료", "쉐어율"}
    if not required_columns.issubset(df.columns):
        st.error("❌ 업로드된 파일에 다음 항목이 모두 포함되어 있어야 합니다:\n" + ", ".join(sorted(required_columns)))
        st.stop()

    if df["쉐어율"].isnull().any():
        st.error("❌ '쉐어율'에 빈 값이 포함되어 있습니다. 모든 행에 값을 입력해주세요.")
        st.stop()

    # 4) 환산율 분류 (컨벤션 & 썸머 계산)
    def classify(row):
        보험사원본 = str(row["보험사"])
        납기 = int(row["납입기간"])

        # 보험사 분류
        if 보험사원본 == "한화생명":
            보험사 = "한화생명"
        elif "생명" in 보험사원본 or 보험사원본 in ["신한라이프"]:
            보험사 = "기타생보"
        elif 보험사원본 in ["한화손보", "삼성화재", "흥국화재", "KB손보"]:
            보험사 = 보험사원본
        elif any(x in 보험사원본 for x in ["손해", "화재", "손보", "해상"]):
            보험사 = "기타손보"
        else:
            보험사 = 보험사원본

        # 컨벤션 기준
        if 보험사 == "한화생명":
            conv_rate = 150
        elif 보험사 in ["한화손보", "삼성화재", "흥국화재", "KB손보"]:
            conv_rate = 250
        elif 보험사 == "기타손보":
            conv_rate = 200
        elif 보험사 == "기타생보":
            conv_rate = 100 if 납기 >= 10 else 50
        else:
            conv_rate = 0

        # 썸머 기준
        if 보험사 == "한화생명":
            summ_rate = 150 if 납기 >= 10 else 100
        elif 보험사 == "기타생보":
            summ_rate = 100 if 납기 >= 10 else 30
        elif 보험사 in ["한화손보", "삼성화재", "흥국화재", "KB손보"]:
            summ_rate = 200 if 납기 >= 10 else 100
        elif 보험사 == "기타손보":
            summ_rate = 100 if 납기 >= 10 else 50
        else:
            summ_rate = 0

        return pd.Series([conv_rate, summ_rate])

    df[["컨벤션율", "썸머율"]] = df.apply(classify, axis=1)

    # 쉐어율 정규화
    df["쉐어율"] = df["쉐어율"].apply(lambda x: float(str(x).replace('%','')) if pd.notnull(x) else x)

    # 실적보험료 (필요시 쉐어율 반영 주석 해제)
    df["실적보험료"] = df["보험료"]  # * df["쉐어율"] / 100

    # 환산금액 계산
    df["컨벤션환산금액"] = df["실적보험료"] * df["컨벤션율"] / 100
    df["썸머환산금액"] = df["실적보험료"] * df["썸머율"] / 100

    # 날짜 유효성 체크
    df["계약일자_raw"] = pd.to_datetime(df["계약일자"], errors="coerce")
    invalid_dates = df[df["계약일자_raw"].isna()]
    if not invalid_dates.empty:
        st.warning(f"⚠️ {len(invalid_dates)}건의 계약일자가 날짜로 인식되지 않았습니다. 엑셀에서 '2025-07-23'처럼 정확한 형식으로 입력해주세요.")

    # ✅ 수금자명 목록 및 화면 필터
    collectors = ["전체"] + sorted(df["수금자명"].astype(str).unique().tolist())
    selected_collector = st.selectbox("👤 수금자명 선택", collectors, index=0)

    if selected_collector != "전체":
        show_df = df[df["수금자명"].astype(str) == selected_collector].copy()
    else:
        show_df = df.copy()

    # 화면 표시용 포맷팅
    def to_styled(dfin: pd.DataFrame) -> pd.DataFrame:
        _styled = dfin.copy()
        _styled["계약일자"] = pd.to_datetime(_styled["계약일자"], errors="coerce").dt.strftime("%Y-%m-%d")
        _styled["납입기간"] = _styled["납입기간"].astype(str) + "년"
        _styled["보험료"] = _styled["보험료"].map("{:,.0f} 원".format)
        _styled["쉐어율"] = _styled["쉐어율"].astype(str) + " %"
        _styled["실적보험료"] = _styled["실적보험료"].map("{:,.0f} 원".format)
        _styled["컨벤션율"] = _styled["컨벤션율"].astype(str) + " %"
        if SHOW_SUMMER:
            _styled["썸머율"] = _styled["썸머율"].astype(str) + " %"
        _styled["컨벤션환산금액"] = _styled["컨벤션환산금액"].map("{:,.0f} 원".format)
        if SHOW_SUMMER:
            _styled["썸머환산금액"] = _styled["썸머환산금액"].map("{:,.0f} 원".format)

        base_cols = ["수금자명", "계약일자", "보험사", "상품명", "납입기간", "보험료",
                     "컨벤션율"]
        if SHOW_SUMMER:
            base_cols += ["썸머율"]
        base_cols += ["실적보험료", "컨벤션환산금액"]
        if SHOW_SUMMER:
            base_cols += ["썸머환산금액"]
        return _styled[base_cols]

    # 합계 함수
    def sums(dfin: pd.DataFrame):
        performance_sum = dfin["실적보험료"].sum()
        convention_sum  = dfin["컨벤션환산금액"].sum()
        summer_sum      = dfin["썸머환산금액"].sum() if SHOW_SUMMER else 0
        return performance_sum, convention_sum, summer_sum

    # 목표/갭
    CONV_TARGET = 1_500_000
    SUMM_TARGET = 3_000_000

    # ── 화면 표시(선택된 수금자 기준) ────────────────────────────
    disp_styled = to_styled(show_df)
    st.subheader(f"📄 {'전체' if selected_collector=='전체' else selected_collector} 환산 결과")
    st.dataframe(disp_styled, use_container_width=True)

    perf_sum, conv_sum, summ_sum = sums(show_df)
    conv_gap = conv_sum - CONV_TARGET
    summ_gap = (summ_sum - SUMM_TARGET) if SHOW_SUMMER else 0

    st.subheader("📈 총합")
    st.markdown(
        f"""
        <div style='border: 2px solid #1f77b4; border-radius: 10px; padding: 20px; background-color: #f7faff; margin-bottom: 20px;'>
            <h4 style='color:#1f77b4;'>📈 총합 요약</h4>
            <p><strong>▶ 실적보험료 합계:</strong> {perf_sum:,.0f} 원</p>
            <p><strong>▶ 컨벤션 기준 합계:</strong> {conv_sum:,.0f} 원</p>
            {f"<p><strong>▶ 썸머 기준 합계:</strong> {summ_sum:,.0f} 원</p>" if SHOW_SUMMER else ""}
        </div>
        """,
        unsafe_allow_html=True
    )

    def gap_box(title, amount):
        if amount > 0:
            color = "#e6f4ea"; text_color = "#0c6b2c"; symbol = f"+{amount:,.0f} 원 초과"
        elif amount < 0:
            color = "#fdecea"; text_color = "#b80000"; symbol = f"{amount:,.0f} 원 부족"
        else:
            color = "#f3f3f3"; text_color = "#000000"; symbol = "기준 달성"
        return f"""
        <div style='border: 1px solid {text_color}; border-radius: 8px; background-color: {color}; padding: 12px; margin: 10px 0;'>
            <strong style='color:{text_color};'>{title}: {symbol}</strong>
        </div>
        """

    st.markdown(gap_box("컨벤션 목표 대비", conv_gap), unsafe_allow_html=True)
    if SHOW_SUMMER:
        st.markdown(gap_box("썸머 목표 대비", summ_gap), unsafe_allow_html=True)

    # ── 요약 테이블(모든 수금자 집계) ────────────────────────────
    st.subheader("🧮 수금자명별 요약")
    group = df.groupby("수금자명", dropna=False).agg(
        실적보험료합계=("실적보험료", "sum"),
        컨벤션합계=("컨벤션환산금액", "sum"),
        썸머합계=("썸머환산금액", "sum") if SHOW_SUMMER else ("실적보험료", "sum") # dummy when False
    ).reset_index()
    if not SHOW_SUMMER:
        group.drop(columns=["썸머합계"], inplace=True)

    # 목표 대비 컬럼
    group["컨벤션_갭"] = group["컨벤션합계"] - CONV_TARGET
    if SHOW_SUMMER:
        group["썸머_갭"] = group["썸머합계"] - SUMM_TARGET

    # 화면용 포맷
    disp_group = group.copy()
    for col in ["실적보험료합계", "컨벤션합계", "컨벤션_갭"]:
        disp_group[col] = disp_group[col].map("{:,.0f} 원".format)
    if SHOW_SUMMER:
        for col in ["썸머합계", "썸머_갭"]:
            disp_group[col] = disp_group[col].map("{:,.0f} 원".format)
    st.dataframe(disp_group, use_container_width=True)

    # ── 엑셀 출력: 요약 시트 + 수금자명별 시트 ───────────────────
    def write_table(ws, df_for_sheet: pd.DataFrame):
        # df_for_sheet는 이미 문자열 포맷팅된 테이블(헤더 포함)
        for r_idx, row in enumerate(dataframe_to_rows(df_for_sheet, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.alignment = Alignment(horizontal="center", vertical="center")
        end_col_letter = ws.cell(row=1, column=df_for_sheet.shape[1]).column_letter
        end_row = ws.max_row
        table = Table(displayName=f"tbl_{ws.title.replace(' ', '_')}", ref=f"A1:{end_col_letter}{end_row}")
        style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        table.tableStyleInfo = style
        ws.add_table(table)
        for column_cells in ws.columns:
            max_len = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
            ws.column_dimensions[column_cells[0].column_letter].width = max_len + 10

    def sums_and_gaps_box(ws, dfin_numeric: pd.DataFrame):
        # 합계/갭 행 삽입 (현재 시트의 테이블 아래)
        performance_sum, convention_sum, summer_sum = sums(dfin_numeric)
        headers = {ws.cell(row=1, column=i).value: i for i in range(1, ws.max_column + 1)}

        # 합계 라벨 위치: 컨벤션율 컬럼 아래에 '총 합계' 표기
        sum_row = ws.max_row + 2
        if "컨벤션율" in headers:  # 포맷팅된 시트엔 있음
            ws.cell(row=sum_row, column=headers["컨벤션율"], value="총 합계").alignment = Alignment(horizontal="center")

        # 금액 기입
        if "실적보험료" in headers:
            ws.cell(row=sum_row, column=headers["실적보험료"], value=f"{performance_sum:,.0f} 원").alignment = Alignment(horizontal="center")
            ws.cell(row=sum_row, column=headers["실적보험료"]).font = Font(bold=True)
        if "컨벤션환산금액" in headers:
            ws.cell(row=sum_row, column=headers["컨벤션환산금액"], value=f"{convention_sum:,.0f} 원").alignment = Alignment(horizontal="center")
            ws.cell(row=sum_row, column=headers["컨벤션환산금액"]).font = Font(bold=True)
        if SHOW_SUMMER and "썸머환산금액" in headers:
            ws.cell(row=sum_row, column=headers["썸머환산금액"], value=f"{summer_sum:,.0f} 원").alignment = Alignment(horizontal="center")
            ws.cell(row=sum_row, column=headers["썸머환산금액"]).font = Font(bold=True)

        # 갭
        def get_gap_style(amount):
            if amount > 0:
                return f"+{amount:,.0f} 원 초과", "008000"
            elif amount < 0:
                return f"{amount:,.0f} 원 부족", "FF0000"
            else:
                return "기준 달성", "000000"

        conv_gap = convention_sum - CONV_TARGET
        result_row = sum_row + 2
        if "컨벤션환산금액" in headers and "실적보험료" in headers:
            ws.cell(row=result_row, column=headers["컨벤션환산금액"], value="컨벤션 기준 대비").alignment = Alignment(horizontal="center")
            ct, cc = get_gap_style(conv_gap)
            ws.cell(row=result_row, column=headers["실적보험료"], value=ct).alignment = Alignment(horizontal="center")
            ws.cell(row=result_row, column=headers["실적보험료"]).font = Font(bold=True, color=cc)

        if SHOW_SUMMER and "썸머환산금액" in headers and "실적보험료" in headers:
            summ_gap = summer_sum - SUMM_TARGET
            ws.cell(row=result_row + 1, column=headers["컨벤션환산금액"], value="썸머 기준 대비").alignment = Alignment(horizontal="center")
            stt, stc = get_gap_style(summ_gap)
            ws.cell(row=result_row + 1, column=headers["실적보험료"], value=stt).alignment = Alignment(horizontal="center")
            ws.cell(row=result_row + 1, column=headers["실적보험료"]).font = Font(bold=True, color=stc)

    # 엑셀 워크북 생성
    wb = Workbook()
    # 요약 시트
    ws_summary = wb.active
    ws_summary.title = "요약"

    # 요약 시트 표(숫자→문자 포맷팅)
    summary_disp = group.copy()
    summary_disp_fmt = summary_disp.copy()
    money_cols = ["실적보험료합계", "컨벤션합계", "컨벤션_갭"]
    if SHOW_SUMMER:
        money_cols += ["썸머합계", "썸머_갭"]
    for col in money_cols:
        summary_disp_fmt[col] = summary_disp_fmt[col].map(lambda x: f"{x:,.0f} 원")

    write_table(ws_summary, summary_disp_fmt)

    # 수금자명별 상세 시트
    for collector in sorted(df["수금자명"].astype(str).unique().tolist()):
        sub = df[df["수금자명"].astype(str) == collector].copy()
        # 화면용과 동일 포맷
        styled_sub = to_styled(sub)

        ws = wb.create_sheet(title=collector[:31])  # 시트명 31자 제한
        write_table(ws, styled_sub)
        # 합계/갭 표기(원본 숫자 기반)
        sums_and_gaps_box(ws, sub)

    # 엑셀 저장
    excel_output = BytesIO()
    wb.save(excel_output)
    excel_output.seek(0)

    st.download_button(
        label="📥 환산 결과 엑셀 다운로드 (요약 + 수금자별 시트)",
        data=excel_output,
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # 제외된 계약(있다면) 화면 하단에 안내
    if not excluded_df.empty:
        st.subheader("🚫 제외된 계약 목록")
        excluded_display = excluded_df[["수금자명","계약일","보험사","상품명","납입기간","초회보험료","납입방법"]].copy()
        excluded_display.columns = ["수금자명","계약일","보험사","상품명","납입기간","보험료","납입방법"]
        st.dataframe(excluded_display, use_container_width=True)

        st.markdown("📝 **제외 계약별 사유:**")
        for _, row in excluded_df.iterrows():
            상품명 = row.get("상품명", "")
            사유들 = []
            if isinstance(row.get("납입방법", ""), str) and "일시납" in row["납입방법"]:
                사유들.append("일시납")
            if isinstance(row.get("상품군2", ""), str) and ("연금성" in row["상품군2"] or "저축성" in row["상품군2"]):
                사유들.append("연금/저축성")
            if isinstance(row.get("계약상태", ""), str) and "철회" in row["계약상태"]:
                사유들.append("철회")
            if isinstance(row.get("계약상태", ""), str) and "해약" in row["계약상태"]:
                사유들.append("해약")
            사유_텍스트 = " / ".join(사유들) if 사유들 else "제외 조건 미상"
            st.markdown(f"- ({상품명}) → 제외사유: {사유_텍스트}")
