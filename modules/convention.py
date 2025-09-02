import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
import os

# ✅ 한 줄로 제어: True로 바꾸면 썸머 기준 즉시 복원
SHOW_SUMMER = True

def run():
    st.set_page_config(page_title="보험 계약 환산기", layout="wide")
    st.title("📊 보험 계약 실적 환산기 (컨벤션{} 기준)".format(" & 썸머" if SHOW_SUMMER else ""))

        # 👉 여기부터 사이드바 안내 추가
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
    # 👉 여기까지

    uploaded_file = st.file_uploader("📂 계약 목록 Excel 파일 업로드 (.xlsx)", type=["xlsx"])

    if uploaded_file:
        base_filename = os.path.splitext(uploaded_file.name)[0]
        download_filename = f"{base_filename}_환산결과.xlsx"

        # 1) 필요한 컬럼만 로드
        columns_needed = ["계약일", "보험사", "상품명", "납입기간", "초회보험료", "쉐어율", "납입방법", "상품군2", "계약상태"]
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
        df.rename(columns={
            "계약일": "계약일자",
            "초회보험료": "보험료"
        }, inplace=True)

        # 제외된 계약 목록(있다면) 표시
        if not excluded_df.empty:
            st.subheader("🚫 제외된 계약 목록")
            excluded_display = excluded_df[["계약일", "보험사", "상품명", "납입기간", "초회보험료", "납입방법"]]
            excluded_display.columns = ["계약일", "보험사", "상품명", "납입기간", "보험료", "납입방법"]
            st.dataframe(excluded_display)

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

        # 3) 필수 항목 체크
        required_columns = {"계약일자", "보험사", "상품명", "납입기간", "보험료", "쉐어율"}
        if not required_columns.issubset(df.columns):
            st.error("❌ 업로드된 파일에 다음 항목이 모두 포함되어 있어야 합니다:\n" + ", ".join(required_columns))
            st.stop()

        if df["쉐어율"].isnull().any():
            st.error("❌ '쉐어율'에 빈 값이 포함되어 있습니다. 모든 행에 값을 입력해주세요.")
            st.stop()

        # 4) 환산율 분류 (컨벤션 & 썸머 둘 다 계산하지만, 노출은 플래그로 제어)
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
                conv_rate = 120
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

        # 합계
        performance_sum = df["실적보험료"].sum()
        convention_sum = df["컨벤션환산금액"].sum()
        summer_sum = df["썸머환산금액"].sum() if SHOW_SUMMER else 0

        # 화면 표시용 스타일링
        styled_df = df.copy()
        styled_df["계약일자"] = pd.to_datetime(styled_df["계약일자"], errors="coerce")
        invalid_dates = styled_df[styled_df["계약일자"].isna()]
        if not invalid_dates.empty:
            st.warning(f"⚠️ {len(invalid_dates)}건의 계약일자가 날짜로 인식되지 않았습니다. 엑셀에서 '2025-07-23'처럼 정확한 형식으로 입력해주세요.")

        styled_df["계약일자"] = styled_df["계약일자"].dt.strftime("%Y-%m-%d")
        styled_df["납입기간"] = styled_df["납입기간"].astype(str) + "년"
        styled_df["보험료"] = styled_df["보험료"].map("{:,.0f} 원".format)
        styled_df["쉐어율"] = styled_df["쉐어율"].astype(str) + " %"
        styled_df["실적보험료"] = styled_df["실적보험료"].map("{:,.0f} 원".format)
        styled_df["컨벤션율"] = styled_df["컨벤션율"].astype(str) + " %"
        if SHOW_SUMMER:
            styled_df["썸머율"] = styled_df["썸머율"].astype(str) + " %"
        styled_df["컨벤션환산금액"] = styled_df["컨벤션환산금액"].map("{:,.0f} 원".format)
        if SHOW_SUMMER:
            styled_df["썸머환산금액"] = styled_df["썸머환산금액"].map("{:,.0f} 원".format)

        # ✅ 컬럼 순서 (플래그에 따라 동적 구성)
        base_cols = ["계약일자", "보험사", "상품명", "납입기간", "보험료", #"쉐어율", 
                     "컨벤션율"]
        base_cols += (["썸머율"] if SHOW_SUMMER else [])
        base_cols += ["실적보험료", "컨벤션환산금액"]
        base_cols += (["썸머환산금액"] if SHOW_SUMMER else [])
        styled_df = styled_df[base_cols]

        # --- 엑셀 출력 ---
        wb = Workbook()
        ws = wb.active
        ws.title = "환산결과"

        for r_idx, row in enumerate(dataframe_to_rows(styled_df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.alignment = Alignment(horizontal="center", vertical="center")

        end_col_letter = ws.cell(row=1, column=styled_df.shape[1]).column_letter
        end_row = ws.max_row
        table = Table(displayName="환산결과표", ref=f"A1:{end_col_letter}{end_row}")
        style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        table.tableStyleInfo = style
        ws.add_table(table)

        for column_cells in ws.columns:
            max_len = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
            ws.column_dimensions[column_cells[0].column_letter].width = max_len + 10

        # 총합 행(동적 위치)
        sum_row = ws.max_row + 2
        # 헤더 인덱스 매핑
        headers = {ws.cell(row=1, column=i).value: i for i in range(1, ws.max_column + 1)}
        ws.cell(row=sum_row, column=headers["컨벤션율"], value="총 합계").alignment = Alignment(horizontal="center")
        ws.cell(row=sum_row, column=headers["실적보험료"], value="{:,.0f} 원".format(performance_sum)).alignment = Alignment(horizontal="center")
        ws.cell(row=sum_row, column=headers["컨벤션환산금액"], value="{:,.0f} 원".format(convention_sum)).alignment = Alignment(horizontal="center")
        if SHOW_SUMMER:
            ws.cell(row=sum_row, column=headers["썸머환산금액"], value="{:,.0f} 원".format(summer_sum)).alignment = Alignment(horizontal="center")

        # 굵게 처리
        for name in ["컨벤션율", "실적보험료", "컨벤션환산금액"] + (["썸머환산금액"] if SHOW_SUMMER else []):
            ws.cell(row=sum_row, column=headers[name]).font = Font(bold=True)

        # 목표/갭(컨벤션은 항상 노출, 썸머는 조건부)
        convention_target = 1_500_000
        summer_target = 3_000_000

        convention_gap = convention_sum - convention_target
        summer_gap = summer_sum - summer_target if SHOW_SUMMER else 0

        result_row = sum_row + 2
        def get_gap_style(amount):
            if amount > 0:
                return f"+{amount:,.0f} 원 초과", "008000"
            elif amount < 0:
                return f"{amount:,.0f} 원 부족", "FF0000"
            else:
                return "기준 달성", "000000"

        # 컨벤션 갭
        ws.cell(row=result_row, column=headers["컨벤션환산금액"], value="컨벤션 기준 대비").alignment = Alignment(horizontal="center")
        ct, cc = get_gap_style(convention_gap)
        ws.cell(row=result_row, column=headers["실적보험료"], value=ct).alignment = Alignment(horizontal="center")
        ws.cell(row=result_row, column=headers["실적보험료"]).font = Font(bold=True, color=cc)

        # 썸머 갭(선택)
        if SHOW_SUMMER:
            ws.cell(row=result_row + 1, column=headers["컨벤션환산금액"], value="썸머 기준 대비").alignment = Alignment(horizontal="center")
            stt, stc = get_gap_style(summer_gap)
            ws.cell(row=result_row + 1, column=headers["실적보험료"], value=stt).alignment = Alignment(horizontal="center")
            ws.cell(row=result_row + 1, column=headers["실적보험료"]).font = Font(bold=True, color=stc)

        # 다운로드 버퍼
        excel_output = BytesIO()
        wb.save(excel_output)
        excel_output.seek(0)

        # --- 화면 출력 ---
        st.subheader("📄 환산 결과 요약")
        st.dataframe(styled_df)

        st.subheader("📈 총합")
        # 요약 카드(썸머는 조건부 문구)
        st.markdown("""
        <div style='
            border: 2px solid #1f77b4;
            border-radius: 10px;
            padding: 20px;
            background-color: #f7faff;
            margin-bottom: 20px;
        '>
            <h4 style='color:#1f77b4;'>📈 총합 요약</h4>
            <p><strong>▶ 실적보험료 합계:</strong> {:,.0f} 원</p>
            <p><strong>▶ 컨벤션 기준 합계:</strong> {:,.0f} 원</p>
            {}
        </div>
        """.format(
            performance_sum,
            convention_sum,
            f"<p><strong>▶ 썸머 기준 합계:</strong> {summer_sum:,.0f} 원</p>" if SHOW_SUMMER else ""
        ), unsafe_allow_html=True)

        # 갭 박스
        def gap_box(title, amount):
            if amount > 0:
                color = "#e6f4ea"; text_color = "#0c6b2c"; symbol = f"+{amount:,.0f} 원 초과"
            elif amount < 0:
                color = "#fdecea"; text_color = "#b80000"; symbol = f"{amount:,.0f} 원 부족"
            else:
                color = "#f3f3f3"; text_color = "#000000"; symbol = "기준 달성"
            return f"""
            <div style='
                border: 1px solid {text_color};
                border-radius: 8px;
                background-color: {color};
                padding: 12px;
                margin: 10px 0;
            '>
                <strong style='color:{text_color};'>{title}: {symbol}</strong>
            </div>
            """

        st.markdown(gap_box("컨벤션 목표 대비", convention_gap), unsafe_allow_html=True)
        if SHOW_SUMMER:
            st.markdown(gap_box("썸머 목표 대비", summer_gap), unsafe_allow_html=True)

        st.download_button(
            label="📥 환산 결과 엑셀 다운로드",
            data=excel_output,
            file_name=download_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("📤 계약 목록 Excel 파일(.xlsx)을 업로드해주세요.")
