import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
import os, re
import numpy as np

# ── 전역 상수 ────────────────────────────────────────────────
TABLE_SEQ = 0

# ✅ 컨벤션 목표(일반/더블/트리플)
CONV_TARGETS = [
    ("일반", 1_800_000),
    ("더블", 3_600_000),
    ("트리플", 5_400_000),
]

# (유지) 썸머 목표 필요 시 사용
SUMM_TARGET = 3_000_000

# ✅ 필수 조건
MIN_COUNT = 5
HANWHA_MIN_PREMIUM = 30_000  # 한화생명 가동 3만원 이상 1건 필수


# ── 유틸 ────────────────────────────────────────────────────
def unique_sheet_name(wb, base, limit=31):
    name = str(base)[:limit] if base else "Sheet"
    if name not in wb.sheetnames:
        return name
    i = 2
    while True:
        suffix = f"_{i}"
        trunc = limit - len(suffix)
        cand = f"{name[:trunc]}{suffix}"
        if cand not in wb.sheetnames:
            return cand
        i += 1


def header_idx(ws, name, default=None):
    for i in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=i).value == name:
            return i
    return default


def safe_table_name(base: str) -> str:
    name = re.sub(r"[^A-Za-z0-9_]", "_", base)
    if not re.match(r"^[A-Za-z_]", name):
        name = f"tbl_{name}"
    return name[:254]


def autosize_columns_full(ws, padding=10):
    for column_cells in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = max_len + padding


def mark(ok: bool) -> str:
    return "✅" if ok else "❌"


def check_requirements(dfin: pd.DataFrame):
    count_ok = len(dfin) >= MIN_COUNT
    hanwha_ok = (
        (dfin["보험사"].astype(str).str.strip() == "한화생명")
        & (pd.to_numeric(dfin["보험료"], errors="coerce").fillna(0) >= HANWHA_MIN_PREMIUM)
    ).any()
    return count_ok, hanwha_ok


# ── 데이터 준비 단계 ─────────────────────────────────────────
def load_df(uploaded_file: BytesIO) -> pd.DataFrame:
    columns_needed = [
        "수금자명", "계약일", "보험사", "상품명", "납입기간",
        "초회보험료", "쉐어율", "납입방법", "상품군2", "계약상태"
    ]
    return pd.read_excel(uploaded_file, usecols=columns_needed)


def exclude_contracts(df: pd.DataFrame):
    excluded_df = pd.DataFrame()

    if {"납입방법", "상품군2", "계약상태"}.issubset(df.columns):
        tmp = df.copy()
        tmp["납입방법"] = tmp["납입방법"].astype(str).str.strip()
        tmp["상품군2"] = tmp["상품군2"].astype(str).str.strip()
        tmp["계약상태"] = tmp["계약상태"].astype(str).str.strip()

        is_lumpsum = tmp["납입방법"].str.contains("일시납", na=False)
        is_savings = tmp["상품군2"].str.contains("연금성|저축성", na=False)
        is_cancelled = tmp["계약상태"].str.contains("철회|해약|실효", na=False)

        is_excluded = is_lumpsum | is_savings | is_cancelled
        excluded_df = tmp[is_excluded].copy()
        df_valid = tmp[~is_excluded].copy()
        return df_valid, excluded_df

    return df.copy(), excluded_df


def build_excluded_with_reason(exdf: pd.DataFrame) -> pd.DataFrame:
    base_cols = ["수금자명", "계약일자", "보험사", "상품명", "납입기간", "보험료", "납입방법", "제외사유"]
    if exdf is None or exdf.empty:
        return pd.DataFrame(columns=base_cols)

    tmp = exdf.copy()

    def reason_row(row):
        r = []
        if "일시납" in str(row.get("납입방법", "")):
            r.append("일시납")
        if ("연금성" in str(row.get("상품군2", ""))) or ("저축성" in str(row.get("상품군2", ""))):
            r.append("연금/저축성")

        status = str(row.get("계약상태", ""))
        if "철회" in status:
            r.append("철회")
        if "해약" in status:
            r.append("해약")
        if "실효" in status:
            r.append("실효")

        return " / ".join(r) if r else "제외 조건 미상"

    tmp["제외사유"] = tmp.apply(reason_row, axis=1)

    tmp_disp = tmp[["수금자명","계약일","보험사","상품명","납입기간","초회보험료","납입방법","제외사유"]].copy()
    tmp_disp.rename(columns={"계약일":"계약일자","초회보험료":"보험료"}, inplace=True)

    tmp_disp["계약일자"] = pd.to_datetime(tmp_disp["계약일자"], errors="coerce").dt.strftime("%Y-%m-%d")
    tmp_disp["납입기간"] = tmp_disp["납입기간"].apply(lambda x: f"{int(float(x))}년" if pd.notnull(x) else "")
    tmp_disp["보험료"] = tmp_disp["보험료"].map(lambda x: "{:,.0f} 원".format(x) if pd.notnull(x) else "")
    return tmp_disp[base_cols]


def compute_rates_and_amounts(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.rename(columns={"계약일": "계약일자", "초회보험료": "보험료"}, inplace=True)

    df["납입기간_num"] = pd.to_numeric(df["납입기간"], errors="coerce").fillna(0).astype(int)
    ins = df["보험사"].astype(str).str.strip()

    is_hanhwa_life = ins.eq("한화생명")

    is_db_nonlife = ins.str.contains("DB", na=False) & ins.str.contains("손", na=False)
    is_heungkuk = ins.str.contains("흥국", na=False) & ins.str.contains("화재", na=False)

    is_kb_nonlife = ins.str.contains("KB", na=False) & ins.str.contains("손", na=False)
    is_hanhwa_nonlife = ins.str.contains("한화", na=False) & ins.str.contains("손", na=False) & (~is_hanhwa_life)

    is_nonlife_generic = ins.str.contains("손해|손보|화재|해상", regex=True, na=False)
    is_life_other = (~is_hanhwa_life) & (ins.str.contains("생명", na=False) | ins.isin(["신한라이프"]))

    conv_rate = np.select(
        [
            is_hanhwa_life,
            is_db_nonlife | is_heungkuk,
            is_kb_nonlife | is_hanhwa_nonlife,
            is_nonlife_generic,
            is_life_other & (df["납입기간_num"] >= 10),
            is_life_other & (df["납입기간_num"] < 10),
        ],
        [
            150,  # 한화생명
            300,  # DB손해/흥국화재
            250,  # KB손해/한화손해
            200,  # 손해보험 일반
            100,  # 생명보험 10년납 이상
            50,   # 생명보험 10년납 미만
        ],
        default=0
    ).astype(int)

    df["컨벤션율"] = conv_rate
    df["썸머율"] = conv_rate  # 썸머 토글 기능 유지(기준표 별도 없으므로 동일)

    # ✅ 쉐어율은 참고 컬럼 (보험료가 이미 반영된 값 유지)
    df["쉐어율"] = df["쉐어율"].apply(lambda x: float(str(x).replace("%", "")) if pd.notnull(x) else x)

    df["실적보험료"] = df["보험료"]
    df["컨벤션환산금액"] = df["실적보험료"] * df["컨벤션율"] / 100
    df["썸머환산금액"] = df["실적보험료"] * df["썸머율"] / 100

    df["계약일자_raw"] = pd.to_datetime(df["계약일자"], errors="coerce")
    return df


def make_group(df: pd.DataFrame, show_summer: bool) -> pd.DataFrame:
    """
    ✅ Q1+Q1(추가) 반영:
    - 컨벤션 달성: ✅/❌
    - 필수조건(5건/한화생명)도 ✅/❌ 컬럼으로 추가
    """
    group_sum = df.groupby("수금자명", dropna=False).agg(
        실적보험료합계=("실적보험료", "sum"),
        컨벤션합계=("컨벤션환산금액", "sum"),
        썸머합계=("썸머환산금액", "sum") if show_summer else ("실적보험료", "sum"),
        건수=("수금자명", "size"),
        한화생명3만=("보험료", lambda s: 0),  # placeholder
    ).reset_index()

    # 한화생명3만 계산(수금자별)
    tmp = df.copy()
    tmp["보험료_num"] = pd.to_numeric(tmp["보험료"], errors="coerce").fillna(0)
    tmp["is_hanwha_ok"] = (tmp["보험사"].astype(str).str.strip() == "한화생명") & (tmp["보험료_num"] >= HANWHA_MIN_PREMIUM)

    hanwha_cnt = tmp.groupby("수금자명", dropna=False)["is_hanwha_ok"].any().reset_index(name="hanwha_ok")
    group_sum = group_sum.drop(columns=["한화생명3만"])
    group_sum = group_sum.merge(hanwha_cnt, on="수금자명", how="left")
    group_sum["hanwha_ok"] = group_sum["hanwha_ok"].fillna(False)

    if not show_summer:
        group_sum.drop(columns=["썸머합계"], inplace=True)

    # ✅ 컨벤션 달성 여부
    for label, target in CONV_TARGETS:
        group_sum[f"컨벤션_{label}달성"] = (group_sum["컨벤션합계"] >= target).map(mark)

    if show_summer:
        group_sum["썸머달성"] = (group_sum["썸머합계"] >= SUMM_TARGET).map(mark)

    # ✅ 필수조건 달성 여부
    group_sum["5건"] = (group_sum["건수"] >= MIN_COUNT).map(mark)
    group_sum["한화생명3만"] = group_sum["hanwha_ok"].map(mark)
    group_sum["전체"] = ((group_sum["건수"] >= MIN_COUNT) & (group_sum["hanwha_ok"])).map(mark)

    # 보기용: 중간 컬럼 정리
    group_sum.drop(columns=["hanwha_ok"], inplace=True)

    # 컬럼 순서 정리(가독성)
    base_cols = ["수금자명", "건수", "5건", "한화생명3만", "실적보험료합계", "컨벤션합계"]
    conv_cols = [f"컨벤션_{label}달성" for label, _ in CONV_TARGETS]
    summer_cols = ["썸머합계", "썸머달성"] if show_summer else []
    group_sum = group_sum[base_cols + conv_cols + summer_cols]

    return group_sum


# ── 화면 표시 ────────────────────────────────────────────────
def to_styled(dfin: pd.DataFrame, show_summer: bool) -> pd.DataFrame:
    _ = dfin.copy()
    _["계약일자"] = pd.to_datetime(_["계약일자"], errors="coerce").dt.strftime("%Y-%m-%d")
    _["납입기간"] = _["납입기간"].astype(str) + "년"
    _["보험료"] = _["보험료"].map("{:,.0f} 원".format)
    _["쉐어율"] = _["쉐어율"].astype(str) + " %"
    _["실적보험료"] = _["실적보험료"].map("{:,.0f} 원".format)
    _["컨벤션율"] = _["컨벤션율"].astype(str) + " %"
    if show_summer:
        _["썸머율"] = _["썸머율"].astype(str) + " %"
    _["컨벤션환산금액"] = _["컨벤션환산금액"].map("{:,.0f} 원".format)
    if show_summer:
        _["썸머환산금액"] = _["썸머환산금액"].map("{:,.0f} 원".format)

    cols = ["수금자명","계약일자","보험사","상품명","납입기간","보험료","컨벤션율"]
    if show_summer: cols += ["썸머율"]
    cols += ["실적보험료","컨벤션환산금액"]
    if show_summer: cols += ["썸머환산금액"]
    return _[cols]


def sums(dfin: pd.DataFrame, show_summer: bool):
    perf = dfin["실적보험료"].sum()
    conv = dfin["컨벤션환산금액"].sum()
    summ = dfin["썸머환산금액"].sum() if show_summer else 0
    return perf, conv, summ


def gap_box(title, amount):
    if amount > 0:
        color = "#e6f4ea"; txt = "#0c6b2c"; sym = f"+{amount:,.0f} 원 초과"
    elif amount < 0:
        color = "#fdecea"; txt = "#b80000"; sym = f"{amount:,.0f} 원 부족"
    else:
        color = "#f3f3f3"; txt = "#000000"; sym = "기준 달성"
    return f"""
    <div style='border: 1px solid {txt}; border-radius: 8px; background-color: {color}; padding: 12px; margin: 10px 0;'>
        <strong style='color:{txt};'>{title}: {sym}</strong>
    </div>
    """


def req_box(title, ok):
    color = "#e6f4ea" if ok else "#fdecea"
    txt = "#0c6b2c" if ok else "#b80000"
    mark_txt = "✅ 충족" if ok else "❌ 미충족"
    return f"""
    <div style='border: 1px solid {txt}; border-radius: 8px; background-color: {color}; padding: 12px; margin: 10px 0;'>
        <strong style='color:{txt};'>{title}: {mark_txt}</strong>
    </div>
    """


# ── 엑셀 출력 ────────────────────────────────────────────────
def write_table(ws, df_for_sheet: pd.DataFrame, start_row: int = 1, name_suffix: str = "A"):
    global TABLE_SEQ
    r_idx = start_row - 1

    for r_idx, row in enumerate(dataframe_to_rows(df_for_sheet, index=False, header=True), start_row):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.alignment = Alignment(horizontal="center", vertical="center")

    end_col_letter = ws.cell(row=start_row, column=df_for_sheet.shape[1]).column_letter
    last_row = r_idx if df_for_sheet.shape[0] > 0 else start_row

    TABLE_SEQ += 1
    display_name = safe_table_name(f"tbl_{ws.title}_{name_suffix}_{TABLE_SEQ}")

    table = Table(displayName=display_name, ref=f"A{start_row}:{end_col_letter}{last_row}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(table)

    autosize_columns_full(ws, padding=5)
    return last_row


def write_requirements_line(ws, base_row: int, dfin: pd.DataFrame):
    c_ok, h_ok = check_requirements(dfin)
    line = (
        f"필수 조건: 최소 {MIN_COUNT}건 이상 {mark(c_ok)}  |  "
        f"한화생명 가동 {HANWHA_MIN_PREMIUM:,.0f}원 이상 1건 {mark(h_ok)}"
    )
    cell = ws.cell(row=base_row, column=1, value=line)
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="left", vertical="center")


def sums_and_gaps_block(ws, perf, conv, summ, show_summer: bool, start_row: int):
    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                         top=Side(style="thin"), bottom=Side(style="thin"))
    sum_fill = PatternFill("solid", fgColor="F2F2F2")

    col_conv_rate = header_idx(ws, "컨벤션율", 1)
    col_perf = header_idx(ws, "실적보험료", 2)
    col_conv_amt = header_idx(ws, "컨벤션환산금액", 3)
    col_summ_amt = header_idx(ws, "썸머환산금액", None)

    # 총합
    sum_row = start_row + 2
    ws.cell(row=sum_row, column=col_conv_rate, value="총 합계").alignment = Alignment(horizontal="center")

    c1 = ws.cell(row=sum_row, column=col_perf, value=f"{perf:,.0f} 원")
    c2 = ws.cell(row=sum_row, column=col_conv_amt, value=f"{conv:,.0f} 원")
    c1.font = Font(bold=True); c2.font = Font(bold=True)
    c1.alignment = Alignment(horizontal="center"); c2.alignment = Alignment(horizontal="center")

    cols_to_style = [col_conv_rate, col_perf, col_conv_amt]
    if show_summer and col_summ_amt:
        c3 = ws.cell(row=sum_row, column=col_summ_amt, value=f"{summ:,.0f} 원")
        c3.font = Font(bold=True)
        c3.alignment = Alignment(horizontal="center")
        cols_to_style.append(col_summ_amt)

    for c in cols_to_style:
        cell = ws.cell(row=sum_row, column=c)
        cell.fill = sum_fill
        cell.border = thin_border

    def style_gap(amount):
        if amount > 0: return f"+{amount:,.0f} 원 초과", "008000"
        if amount < 0: return f"{amount:,.0f} 원 부족", "FF0000"
        return "기준 달성", "000000"

    # 목표 대비(일반/더블/트리플)
    gap_row = sum_row + 2
    r = gap_row
    for label, target in CONV_TARGETS:
        txt, col = style_gap(conv - target)
        ws.cell(row=r, column=col_conv_amt, value=f"컨벤션 {label}({target:,.0f}) 대비").alignment = Alignment(horizontal="center")
        g = ws.cell(row=r, column=col_perf, value=txt)
        g.alignment = Alignment(horizontal="center")
        g.font = Font(bold=True, color=col)
        r += 1

    if show_summer and col_summ_amt:
        txt2, col2 = style_gap(summ - SUMM_TARGET)
        ws.cell(row=r, column=col_conv_amt, value=f"썸머({SUMM_TARGET:,.0f}) 대비").alignment = Alignment(horizontal="center")
        g2 = ws.cell(row=r, column=col_perf, value=txt2)
        g2.alignment = Alignment(horizontal="center")
        g2.font = Font(bold=True, color=col2)
        r += 1

    return r


def build_workbook(df: pd.DataFrame, group: pd.DataFrame, excluded_disp_all: pd.DataFrame, show_summer: bool):
    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = "요약"

    # ✅ 요약표 포맷
    summary_fmt = group.copy()
    if "실적보험료합계" in summary_fmt.columns:
        summary_fmt["실적보험료합계"] = summary_fmt["실적보험료합계"].map(lambda x: f"{x:,.0f} 원")
    if "컨벤션합계" in summary_fmt.columns:
        summary_fmt["컨벤션합계"] = summary_fmt["컨벤션합계"].map(lambda x: f"{x:,.0f} 원")
    if show_summer and "썸머합계" in summary_fmt.columns:
        summary_fmt["썸머합계"] = summary_fmt["썸머합계"].map(lambda x: f"{x:,.0f} 원")

    next_row = write_table(ws_summary, summary_fmt, start_row=1, name_suffix="SUM")
    write_requirements_line(ws_summary, base_row=next_row + 2, dfin=df)

    if not excluded_disp_all.empty:
        ws_summary.cell(row=next_row + 4, column=1, value="제외 계약 목록").font = Font(bold=True)
        next_row = write_table(ws_summary, excluded_disp_all, start_row=next_row + 5, name_suffix="EXC")

    # 수금자별 시트
    collectors = sorted(df["수금자명"].astype(str).unique().tolist())
    for collector in collectors:
        sub = df[df["수금자명"].astype(str) == collector].copy()
        sheet_title = unique_sheet_name(wb, collector)
        ws = wb.create_sheet(title=sheet_title)

        styled_sub = to_styled(sub, show_summer)
        table_last_row = write_table(ws, styled_sub, start_row=1, name_suffix="NORM")

        # 주요 금액 컬럼 최소 열 너비 20
        for header in ["실적보험료", "컨벤션환산금액", "썸머환산금액"]:
            idx = header_idx(ws, header)
            if idx:
                col_letter = ws.cell(row=1, column=idx).column_letter
                cur = ws.column_dimensions[col_letter].width
                ws.column_dimensions[col_letter].width = 20 if (cur is None or cur < 20) else cur

        perf = sub["실적보험료"].sum()
        conv = sub["컨벤션환산금액"].sum()
        summ = sub["썸머환산금액"].sum() if show_summer else 0

        next_row = sums_and_gaps_block(ws, perf, conv, summ, show_summer, start_row=table_last_row)

        # ✅ 수금자별 필수조건 체크(시트에도 유지)
        write_requirements_line(ws, base_row=next_row + 1, dfin=sub)
        next_row = next_row + 2

        # 제외 계약(해당 수금자)
        ex_sub = excluded_disp_all[excluded_disp_all["수금자명"].astype(str) == collector]
        if not ex_sub.empty:
            ws.cell(row=next_row + 1, column=1, value="제외 계약").font = Font(bold=True)
            write_table(ws, ex_sub, start_row=next_row + 2, name_suffix="EXC")

    return wb


# ── 메인 ────────────────────────────────────────────────────
def run():
    st.set_page_config(page_title="보험 계약 환산기", layout="wide")

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
        SHOW_SUMMER = st.toggle("🌞 썸머 기준 포함", value=False)

    st.title("📊 TEST!!! 보험 계약 실적 환산기 (컨벤션{} 기준)".format(" & 썸머" if SHOW_SUMMER else ""))

    uploaded_file = st.file_uploader("📂 계약 목록 Excel 파일 업로드 (.xlsx)", type=["xlsx"])
    if not uploaded_file:
        st.info("📤 계약 목록 Excel 파일(.xlsx)을 업로드해주세요.")
        return

    base_filename = os.path.splitext(uploaded_file.name)[0]
    download_filename = f"{base_filename}_환산결과.xlsx"

    raw = load_df(uploaded_file)
    df_valid, excluded_df = exclude_contracts(raw)
    excluded_disp_all = build_excluded_with_reason(excluded_df)

    # 필수 컬럼 체크(유효 df 기준)
    df_valid.rename(columns={"계약일": "계약일자", "초회보험료": "보험료"}, inplace=True)
    required_columns = {"수금자명", "계약일자", "보험사", "상품명", "납입기간", "보험료", "쉐어율"}
    if not required_columns.issubset(df_valid.columns):
        st.error("❌ 업로드된 파일에 다음 항목이 모두 포함되어 있어야 합니다:\n" + ", ".join(sorted(required_columns)))
        st.stop()
    if df_valid["쉐어율"].isnull().any():
        st.error("❌ '쉐어율'에 빈 값이 포함되어 있습니다. 모든 행에 값을 입력해주세요.")
        st.stop()

    # 계산
    df = compute_rates_and_amounts(df_valid)

    # 날짜 경고
    invalid_dates = df[df["계약일자_raw"].isna()]
    if not invalid_dates.empty:
        st.warning(f"⚠️ {len(invalid_dates)}건의 계약일자가 날짜로 인식되지 않았습니다. 엑셀에서 '2025-07-23'처럼 입력해주세요.")

    # 제외 건 화면 표시(있을 때만)
    if not excluded_df.empty:
        st.warning(f"⚠️ 제외된 계약 {len(excluded_df)}건 (일시납 / 연금성·저축성 / 철회|해약|실효)")
        st.subheader("🚫 제외된 계약 목록")
        if excluded_disp_all.empty:
            st.caption("제외된 계약이 없습니다.")
        else:
            st.dataframe(excluded_disp_all, use_container_width=True)


    # 수금자 선택
    collectors = ["전체"] + sorted(df["수금자명"].astype(str).unique().tolist())
    selected_collector = st.selectbox("👤 수금자명 선택", collectors, index=0)
    show_df = df if selected_collector == "전체" else df[df["수금자명"].astype(str) == selected_collector].copy()

    # 메인 표
    st.subheader(f"📄 {'전체' if selected_collector=='전체' else selected_collector} 환산 결과")
    st.dataframe(to_styled(show_df, SHOW_SUMMER), use_container_width=True)

    # 총합/목표 대비
    perf_sum, conv_sum, summ_sum = sums(show_df, SHOW_SUMMER)
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
        unsafe_allow_html=True,
    )

    for label, target in CONV_TARGETS:
        st.markdown(gap_box(f"컨벤션 {label}({target:,.0f}) 목표 대비", conv_sum - target), unsafe_allow_html=True)
    if SHOW_SUMMER:
        st.markdown(gap_box(f"썸머({SUMM_TARGET:,.0f}) 목표 대비", summ_sum - SUMM_TARGET), unsafe_allow_html=True)

    # 필수 조건 체크(선택된 수금자 기준)
    st.subheader("✅ 필수 조건 체크")
    c_ok, h_ok = check_requirements(show_df)
    st.markdown(req_box(f"필수 건수 {MIN_COUNT}건 이상", c_ok), unsafe_allow_html=True)
    st.markdown(req_box(f"한화생명 가동 {HANWHA_MIN_PREMIUM:,.0f}원 이상 1건", h_ok), unsafe_allow_html=True)

    # ✅ 수금자별 요약 (필수조건 컬럼 포함)
    st.subheader("🧮 수금자명별 요약")
    group = make_group(df, SHOW_SUMMER)

    disp_group = group.copy()
    # 금액 컬럼 포맷
    disp_group["실적보험료합계"] = disp_group["실적보험료합계"].map("{:,.0f} 원".format)
    disp_group["컨벤션합계"] = disp_group["컨벤션합계"].map("{:,.0f} 원".format)
    if SHOW_SUMMER and "썸머합계" in disp_group.columns:
        disp_group["썸머합계"] = disp_group["썸머합계"].map("{:,.0f} 원".format)

    st.dataframe(disp_group, use_container_width=True)

    # 엑셀 생성/다운로드 (요약 시트에도 동일 요약표가 들어감)
    wb = build_workbook(df, group, excluded_disp_all, SHOW_SUMMER)
    excel_output = BytesIO()
    wb.save(excel_output)
    excel_output.seek(0)

    st.download_button(
        label="📥 환산 결과 엑셀 다운로드 (요약 + 수금자별 시트 + 제외사유 + 필수조건 표시)",
        data=excel_output,
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    run()
