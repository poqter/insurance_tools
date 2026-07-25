from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from io import BytesIO
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

APP_TITLE = "보험 리모델링 비교 제안서"

# ---------- colors ----------
NAVY = "17365D"
NAVY_DARK = "102A43"
BLUE = "2F75B5"
GREEN = "2E7D5B"
ORANGE = "B96B26"
TEAL = "168A8A"
WHITE = "FFFFFF"
BLACK = "1F2937"
GRAY = "E9EDF2"
GRAY_DARK = "667085"
LIGHT_BLUE = "EAF2F8"
LIGHT_GREEN = "E8F5EE"
LIGHT_ORANGE = "FFF2E6"
LIGHT_GRAY = "F7F9FC"
LIGHT_GOLD = "FBF4E6"
THIN_GRAY = Side(style="thin", color="D7DDE5")
MEDIUM_NAVY = Side(style="medium", color=NAVY)

CHANGE_OPTIONS = [
    "선택하세요 ▼", "새로 추가", "보장금액 증가", "보장금액 감소", "보장 범위 확대",
    "보장 범위 축소", "보장기간 연장", "보장기간 단축", "지급 횟수 증가",
    "그대로 유지", "정리 또는 삭제", "직접 입력",
]
DISPLAY_OPTIONS = ["선택하세요 ▼", "핵심으로 표시", "상세에만 표시", "출력하지 않음"]
CONTRACT_OPTIONS = [
    "선택하세요 ▼", "유지", "감액 검토", "일부 특약 조정", "해지 검토",
    "신규 승인 후 결정", "추가 확인 필요",
]
PROPOSAL_OPTIONS = [
    "선택하세요 ▼",
    "보험료 부담 완화형 추천안",
    "핵심 보장 보완형 추천안",
    "동일 예산 재구성형 추천안",
    "균형 보장형 추천안",
    "보장 강화형 추천안",
    "특정 위험 집중형 추천안",
    "기존 계약 유지 중심 제안안",
    "맞춤 재구성형 추천안",
    "직접 입력",
]
PRIORITY_OPTIONS = [
    "월 보험료 부담", "총 납입 부담", "암 진단비", "암 치료비", "뇌·심장 보장",
    "수술비", "간병 보장", "소득 공백 대비", "기존 실손보험 유지",
    "기존 계약 최대한 유지", "비갱신형 중심", "보장기간", "가족력 관련 보장",
]

POSITIVE_CHANGES = {"새로 추가", "보장금액 증가", "보장 범위 확대", "보장기간 연장", "지급 횟수 증가"}
NEGATIVE_CHANGES = {"보장금액 감소", "보장 범위 축소", "보장기간 단축", "정리 또는 삭제"}


@dataclass(frozen=True)
class AnalysisResult:
    monthly_delta: int
    annual_delta: int
    total_delta: int | None
    price_direction: str
    coverage_direction: str
    result_type: str
    warnings: tuple[str, ...]


@dataclass
class CustomerData:
    index: int
    name: str
    proposal_label: str
    priorities: list[str]
    old_monthly: int
    new_monthly: int
    old_total: int | None
    new_total: int | None
    changes: pd.DataFrame
    contracts: pd.DataFrame
    analysis: AnalysisResult
    headline: str


# ---------- general helpers ----------
def normalize_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def parse_money(value: Any) -> int:
    text = re.sub(r"[^0-9-]", "", normalize_text(value))
    if text in {"", "-"}:
        return 0
    try:
        return int(text)
    except ValueError:
        return 0


def format_change(value: int | None) -> str:
    if value is None:
        return "비교하지 않음"
    if value < 0:
        return f"{abs(value):,}원 절감"
    if value > 0:
        return f"{value:,}원 증가"
    return "변동 없음"


def format_compact_won(value: int | None) -> str:
    if value is None:
        return "-"
    sign = "-" if value < 0 else ""
    amount = abs(int(value))
    if amount >= 100_000_000:
        eok = amount // 100_000_000
        man = (amount % 100_000_000) // 10_000
        return f"{sign}{eok}억 {man:,}만원" if man else f"{sign}{eok}억원"
    if amount >= 10_000:
        return f"{sign}{amount // 10_000:,}만원"
    return f"{sign}{amount:,}원"


def safe_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip() or "보험리모델링_비교안"


def default_changes_df() -> pd.DataFrame:
    return pd.DataFrame([
        {
            "변경할 보장": "", "기존에는": "", "변경 후에는": "",
            "어떻게 달라지나요? [목록 선택]": "선택하세요 ▼", "왜 바꾸나요?": "",
            "변경 후 월 보험료": "", "첫 장 표시 [목록 선택]": "핵심으로 표시",
        },
        {
            "변경할 보장": "", "기존에는": "", "변경 후에는": "",
            "어떻게 달라지나요? [목록 선택]": "선택하세요 ▼", "왜 바꾸나요?": "",
            "변경 후 월 보험료": "", "첫 장 표시 [목록 선택]": "핵심으로 표시",
        },
    ])


def default_contract_df() -> pd.DataFrame:
    return pd.DataFrame([
        {"기존 계약/보장": "", "처리 방향 [목록 선택]": "선택하세요 ▼", "판단 근거": "", "진행 조건": ""}
    ])


def clean_changes(df: pd.DataFrame | None) -> pd.DataFrame:
    if df is None or df.empty:
        return default_changes_df().iloc[0:0]
    result = df.copy()
    result["변경할 보장"] = result["변경할 보장"].map(normalize_text)
    result = result[result["변경할 보장"] != ""].reset_index(drop=True)
    for col in ["어떻게 달라지나요? [목록 선택]", "첫 장 표시 [목록 선택]"]:
        if col in result:
            result[col] = result[col].replace("선택하세요 ▼", "")
    return result


def clean_contracts(df: pd.DataFrame | None) -> pd.DataFrame:
    if df is None or df.empty:
        return default_contract_df().iloc[0:0]
    result = df.copy()
    result["기존 계약/보장"] = result["기존 계약/보장"].map(normalize_text)
    result = result[result["기존 계약/보장"] != ""].reset_index(drop=True)
    result["처리 방향 [목록 선택]"] = result["처리 방향 [목록 선택]"].replace("선택하세요 ▼", "")
    return result


def detect_coverage_direction(changes: pd.DataFrame) -> str:
    if changes.empty:
        return "정보 부족"
    types = set(changes["어떻게 달라지나요? [목록 선택]"].fillna("").astype(str))
    if types & POSITIVE_CHANGES and types & NEGATIVE_CHANGES:
        return "강화와 조정 혼합"
    if types & POSITIVE_CHANGES:
        return "강화"
    if types & NEGATIVE_CHANGES:
        return "축소/조정"
    return "유지/재배분"


def analyze(old_monthly: int, new_monthly: int, old_total: int | None, new_total: int | None, changes: pd.DataFrame) -> AnalysisResult:
    monthly_delta = new_monthly - old_monthly
    annual_delta = monthly_delta * 12
    total_delta = None if old_total is None or new_total is None else new_total - old_total
    price_direction = "감소" if monthly_delta < 0 else "증가" if monthly_delta > 0 else "동일"
    coverage_direction = detect_coverage_direction(changes)
    if price_direction == "감소" and coverage_direction in {"강화", "강화와 조정 혼합"}:
        result_type = "효율 개선형"
    elif price_direction == "감소":
        result_type = "보험료 절감형"
    elif price_direction == "동일" and coverage_direction == "강화":
        result_type = "동일 예산 강화형"
    elif price_direction == "증가" and coverage_direction in {"강화", "강화와 조정 혼합"}:
        result_type = "보장 강화형"
    else:
        result_type = "맞춤 재구성형"
    warnings: list[str] = []
    if not old_monthly:
        warnings.append("기존 월 보험료가 입력되지 않았습니다.")
    if not new_monthly:
        warnings.append("변경 후 월 보험료가 입력되지 않았습니다.")
    if changes.empty:
        warnings.append("핵심 변경 내용이 없습니다.")
    return AnalysisResult(monthly_delta, annual_delta, total_delta, price_direction, coverage_direction, result_type, tuple(warnings))


def priority_keywords(priority: str) -> list[str]:
    mapping = {
        "월 보험료 부담": ["보험료", "월납", "부담"],
        "총 납입 부담": ["총납", "납입", "보험료"],
        "암 진단비": ["암", "진단"],
        "암 치료비": ["암", "치료"],
        "뇌·심장 보장": ["뇌", "심장", "심혈관", "뇌혈관", "순환계"],
        "수술비": ["수술"],
        "간병 보장": ["간병", "간호"],
        "소득 공백 대비": ["생활", "소득", "진단", "후유"],
        "기존 실손보험 유지": ["실손", "실비"],
        "기존 계약 최대한 유지": ["유지", "기존"],
        "비갱신형 중심": ["비갱신", "갱신"],
        "보장기간": ["만기", "보장기간", "기간"],
        "가족력 관련 보장": ["암", "뇌", "심장", "가족력"],
    }
    return mapping.get(priority, [priority])


def prioritized_changes(changes: pd.DataFrame, priorities: list[str], max_items: int) -> pd.DataFrame:
    if changes.empty:
        return changes
    rows = changes[changes["첫 장 표시 [목록 선택]"] == "핵심으로 표시"].copy()
    if rows.empty:
        rows = changes[changes["첫 장 표시 [목록 선택]"] != "출력하지 않음"].copy()
    if rows.empty:
        rows = changes.copy()
    def score(row: pd.Series) -> int:
        text = " ".join(normalize_text(row.get(c)) for c in ["변경할 보장", "기존에는", "변경 후에는", "왜 바꾸나요?"])
        total = 0
        for rank, priority in enumerate(priorities[:2]):
            weight = 20 - rank * 5
            if any(k in text for k in priority_keywords(priority)):
                total += weight
        if normalize_text(row.get("어떻게 달라지나요? [목록 선택]")) in POSITIVE_CHANGES:
            total += 2
        return total
    rows["__score"] = rows.apply(score, axis=1)
    return rows.sort_values(["__score"], ascending=False, kind="stable").drop(columns="__score").head(max_items)


def suggest_proposal(priorities: list[str], analysis: AnalysisResult) -> str:
    p = set(priorities)
    if "월 보험료 부담" in p or "총 납입 부담" in p:
        if analysis.price_direction == "감소":
            return "보험료 부담 완화형 추천안"
    if "기존 계약 최대한 유지" in p or "기존 실손보험 유지" in p:
        return "기존 계약 유지 중심 제안안"
    if p & {"암 진단비", "암 치료비", "뇌·심장 보장", "수술비", "간병 보장"}:
        return "핵심 보장 보완형 추천안" if analysis.price_direction != "증가" else "보장 강화형 추천안"
    if analysis.result_type == "동일 예산 강화형":
        return "동일 예산 재구성형 추천안"
    if analysis.result_type == "보장 강화형":
        return "보장 강화형 추천안"
    if analysis.result_type == "효율 개선형":
        return "균형 보장형 추천안"
    return "맞춤 재구성형 추천안"


def headline_candidates(customer: CustomerData) -> list[str]:
    a = customer.analysis
    names = [normalize_text(v) for v in prioritized_changes(customer.changes, customer.priorities, 3)["변경할 보장"].tolist()]
    focus = "·".join([n for n in names if n]) or "핵심 보장"
    priority = "·".join(customer.priorities[:2])
    delta = abs(a.monthly_delta)
    result: list[str] = []
    if a.price_direction == "감소" and a.coverage_direction in {"강화", "강화와 조정 혼합"}:
        result += [
            f"월 보험료를 {delta:,}원 줄이면서 {focus} 중심으로 보장을 다시 구성한 제안입니다.",
            f"월 부담은 낮추고 {focus}의 필요한 부분은 보완한 제안입니다.",
            f"보험료 부담을 줄이면서 필요한 보장에 예산을 집중했습니다.",
        ]
    elif a.price_direction == "감소":
        result += [
            f"월 보험료를 {customer.old_monthly:,}원에서 {customer.new_monthly:,}원으로 조정한 제안입니다.",
            f"월 {delta:,}원의 부담을 줄여 장기적으로 유지하기 쉬운 구조로 조정했습니다.",
        ]
    elif a.price_direction == "증가" and a.coverage_direction in {"강화", "강화와 조정 혼합"}:
        result += [
            f"월 {delta:,}원의 추가 부담으로 {focus}를 강화하는 제안입니다.",
            f"보험료 절감보다 {focus}의 보장 공백을 줄이는 데 초점을 두었습니다.",
        ]
    elif a.price_direction == "동일" and a.coverage_direction in {"강화", "강화와 조정 혼합"}:
        result += [
            f"현재와 비슷한 월 보험료 안에서 {focus}를 강화한 제안입니다.",
            "추가 부담 없이 보험료가 필요한 보장에 쓰이도록 재배분했습니다.",
        ]
    else:
        result.append("보험료와 보장 변화를 함께 고려해 현재 상황에 맞게 재구성한 제안입니다.")
    if priority:
        result.append(f"고객님이 중요하게 생각하신 {priority}를 우선 반영한 제안입니다.")
    return list(dict.fromkeys(result))


def validate_customer(customer: CustomerData, max_core: int) -> list[str]:
    warnings = list(customer.analysis.warnings)
    if not customer.name:
        warnings.append(f"고객 {customer.index}의 이름을 입력해 주세요.")
    if customer.proposal_label in {"", "선택하세요 ▼", "직접 입력"}:
        warnings.append(f"{customer.name or f'고객 {customer.index}'}의 제안 유형을 확인해 주세요.")
    core_count = len(prioritized_changes(customer.changes, customer.priorities, 999))
    if core_count > max_core:
        warnings.append(f"{customer.name or f'고객 {customer.index}'}의 첫 장 핵심 항목은 {max_core}개까지만 표시됩니다.")
    for idx, row in customer.changes.iterrows():
        name = normalize_text(row.get("변경할 보장")) or f"{idx + 1}번째 항목"
        change = normalize_text(row.get("어떻게 달라지나요? [목록 선택]"))
        if not change:
            warnings.append(f"{name}: 변화 유형을 선택해 주세요.")
    return warnings


def _money_input(label: str, key: str, disabled: bool = False) -> int:
    raw = st.text_input(label, key=key, placeholder="예: 694,580", disabled=disabled)
    amount = parse_money(raw)
    if raw and not disabled:
        st.caption(f"**{amount:,}원** · 약 {format_compact_won(amount)}")
    return amount


def proposal_label_from_state(i: int, analysis: AnalysisResult, priorities: list[str]) -> str:
    selected = st.session_state.get(f"rv2_proposal_{i}", "선택하세요 ▼")
    if selected == "직접 입력":
        return normalize_text(st.session_state.get(f"rv2_proposal_custom_{i}", ""))
    if selected == "선택하세요 ▼":
        return suggest_proposal(priorities, analysis)
    return selected


def auto_title(names: list[str]) -> str:
    valid = [n for n in names if n]
    if not valid:
        return "보험 리모델링 비교안"
    if len(valid) == 1:
        return f"{valid[0]}님 보험 리모델링 비교안"
    title = f"{valid[0]}님·{valid[1]}님 보험 리모델링 비교안"
    return title if len(title) <= 34 else f"{valid[0]}님 가족 보험 리모델링 비교안"


# ---------- Streamlit preview ----------
def direction_style(direction: str) -> tuple[str, str]:
    if direction == "감소":
        return "#2E7D5B", "#E8F5EE"
    if direction == "증가":
        return "#B96B26", "#FFF2E6"
    return "#17365D", "#EAF2F8"


def metric_card_html(title: str, main: str, sub: str, direction: str, compact: bool = False) -> str:
    color, bg = direction_style(direction)
    main_size = 20 if compact else 27
    min_height = 105 if compact else 132
    return f"""
    <div style="border:1px solid #D7DDE5;border-radius:12px;padding:14px;background:{bg};min-height:{min_height}px;box-shadow:0 2px 8px rgba(16,42,67,.06);">
      <div style="font-size:13px;color:#667085;font-weight:700;margin-bottom:7px;">{title}</div>
      <div style="font-size:{main_size}px;color:{color};font-weight:800;line-height:1.18;margin-bottom:8px;">{main}</div>
      <div style="font-size:12px;color:#344054;">{sub}</div>
    </div>"""


def change_card_html(row: pd.Series, compact: bool = False) -> str:
    name = normalize_text(row.get("변경할 보장"))
    old = normalize_text(row.get("기존에는")) or "-"
    new = normalize_text(row.get("변경 후에는")) or "-"
    change = normalize_text(row.get("어떻게 달라지나요? [목록 선택]")) or "변경"
    color, bg = "#2E7D5B", "#E8F5EE"
    if change in NEGATIVE_CHANGES:
        color, bg = "#B96B26", "#FFF2E6"
    elif change == "그대로 유지":
        color, bg = "#17365D", "#EAF2F8"
    title_size = 14 if compact else 16
    value_size = 14 if compact else 16
    return f"""
    <div style="border:1px solid #D7DDE5;border-radius:11px;padding:12px;background:#fff;min-height:{132 if compact else 152}px;">
      <div style="font-size:{title_size}px;color:#17365D;font-weight:800;margin-bottom:10px;">{name}</div>
      <div style="display:flex;align-items:center;gap:6px;justify-content:space-between;">
        <div style="width:42%;text-align:center;color:#667085;font-size:11px;">기존<br><strong style="font-size:{value_size}px;color:#344054;">{old}</strong></div>
        <div style="font-size:19px;color:#98A2B3;">→</div>
        <div style="width:42%;text-align:center;color:#667085;font-size:11px;">변경 후<br><strong style="font-size:{value_size}px;color:#2F75B5;">{new}</strong></div>
      </div>
      <div style="margin-top:11px;padding:6px;border-radius:7px;background:{bg};color:{color};font-size:13px;font-weight:800;text-align:center;">{change}</div>
    </div>"""


def render_customer_preview(customer: CustomerData, compact: bool) -> None:
    title_cols = st.columns([3, 2])
    with title_cols[0]:
        st.markdown(f"### {customer.name or f'고객 {customer.index}'}")
    with title_cols[1]:
        st.markdown(
            f"<div style='padding:7px 9px;border-radius:8px;background:#168A8A;color:white;font-size:12px;font-weight:800;text-align:center;'>{customer.proposal_label}</div>",
            unsafe_allow_html=True,
        )
    if customer.priorities:
        st.markdown(
            f"<div style='font-size:12px;color:#17365D;margin:2px 0 8px 0;'><b>고객 우선사항:</b> {' · '.join(customer.priorities[:2])}</div>",
            unsafe_allow_html=True,
        )
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(metric_card_html("월 보험료", format_change(customer.analysis.monthly_delta), f"{customer.old_monthly:,}원 → {customer.new_monthly:,}원", customer.analysis.price_direction, compact), unsafe_allow_html=True)
    with c2:
        st.markdown(metric_card_html("연간 변화", format_change(customer.analysis.annual_delta), "월 차액 × 12개월", customer.analysis.price_direction, compact), unsafe_allow_html=True)
    with c3:
        td = customer.analysis.total_delta
        d = "동일" if td == 0 or td is None else "감소" if td < 0 else "증가"
        st.markdown(metric_card_html("총 납입액", format_change(td), f"{format_compact_won(customer.old_total)} → {format_compact_won(customer.new_total)}", d, compact), unsafe_allow_html=True)
    rows = prioritized_changes(customer.changes, customer.priorities, 3 if compact else 4)
    if not rows.empty:
        st.markdown("**핵심 보장 변화**")
        columns_per_row = 1 if compact else 2
        items = list(rows.iterrows())
        for start in range(0, len(items), columns_per_row):
            cols = st.columns(columns_per_row)
            for col, (_, row) in zip(cols, items[start:start + columns_per_row]):
                with col:
                    st.markdown(change_card_html(row, compact), unsafe_allow_html=True)
    if not customer.contracts.empty:
        dirs = customer.contracts["처리 방향 [목록 선택]"].fillna("").astype(str)
        keep = int(dirs.str.contains("유지").sum())
        adjust = int(dirs.str.contains("감액|조정|해지|결정").sum())
        check = int(dirs.str.contains("확인").sum())
        st.markdown(f"<div style='margin-top:8px;padding:9px;background:#F7F9FC;border-radius:8px;font-size:13px;'><b>계약 정리:</b> 유지 {keep}건 · 조정·검토 {adjust}건 · 추가 확인 {check}건</div>", unsafe_allow_html=True)
    if customer.headline:
        st.markdown(f"<div style='margin-top:10px;padding:10px 12px;border-left:4px solid #17365D;background:#F7F9FC;font-size:14px;font-weight:800;color:#17365D;'>{customer.headline}</div>", unsafe_allow_html=True)


# ---------- Excel helpers ----------
def border_all() -> Border:
    return Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)


def merge_write(ws, rng: str, value: Any, *, font: Font | None = None, fill: PatternFill | None = None,
                alignment: Alignment | None = None, border: Border | None = None) -> None:
    ws.merge_cells(rng)
    cell = ws[rng.split(":")[0]]
    cell.value = value
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment
    if border:
        for row in ws[rng]:
            for c in row:
                c.border = border


def page_setup(ws, fit_height: int = 1) -> None:
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_height
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.35, bottom=0.35, header=0.1, footer=0.1)
    ws.sheet_view.showGridLines = False
    ws.print_options.horizontalCentered = True


def excel_direction(direction: str) -> tuple[str, str]:
    if direction == "감소":
        return GREEN, LIGHT_GREEN
    if direction == "증가":
        return ORANGE, LIGHT_ORANGE
    return NAVY, LIGHT_BLUE


def write_metric_box(ws, start_col: int, end_col: int, row: int, title: str, main: str, sub: str, direction: str, compact: bool) -> None:
    letter1, letter2 = get_column_letter(start_col), get_column_letter(end_col)
    color, bg = excel_direction(direction)
    merge_write(ws, f"{letter1}{row}:{letter2}{row}", title,
                font=Font(name="맑은 고딕", size=10 if compact else 11, bold=True, color=GRAY_DARK),
                fill=PatternFill("solid", fgColor=bg), alignment=Alignment(horizontal="center", vertical="center"), border=border_all())
    merge_write(ws, f"{letter1}{row+1}:{letter2}{row+2}", main,
                font=Font(name="맑은 고딕", size=14 if compact else 18, bold=True, color=color),
                fill=PatternFill("solid", fgColor=bg), alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=border_all())
    merge_write(ws, f"{letter1}{row+3}:{letter2}{row+3}", sub,
                font=Font(name="맑은 고딕", size=9 if compact else 10, color=BLACK),
                fill=PatternFill("solid", fgColor=bg), alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=border_all())


def write_change_box(ws, start_col: int, end_col: int, row: int, item: pd.Series, compact: bool) -> None:
    l1, l2 = get_column_letter(start_col), get_column_letter(end_col)
    name = normalize_text(item.get("변경할 보장"))
    old = normalize_text(item.get("기존에는")) or "-"
    new = normalize_text(item.get("변경 후에는")) or "-"
    change = normalize_text(item.get("어떻게 달라지나요? [목록 선택]")) or "변경"
    color, bg = GREEN, LIGHT_GREEN
    if change in NEGATIVE_CHANGES:
        color, bg = ORANGE, LIGHT_ORANGE
    elif change == "그대로 유지":
        color, bg = NAVY, LIGHT_BLUE
    merge_write(ws, f"{l1}{row}:{l2}{row}", name,
                font=Font(name="맑은 고딕", size=11 if compact else 12, bold=True, color=NAVY),
                fill=PatternFill("solid", fgColor=WHITE), alignment=Alignment(horizontal="left", vertical="center"), border=border_all())
    mid = (start_col + end_col) // 2
    lm, rm = get_column_letter(mid), get_column_letter(mid + 1)
    merge_write(ws, f"{l1}{row+1}:{lm}{row+2}", f"기존\n{old}",
                font=Font(name="맑은 고딕", size=10 if compact else 11, bold=True, color=BLACK),
                fill=PatternFill("solid", fgColor=LIGHT_GRAY), alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=border_all())
    merge_write(ws, f"{rm}{row+1}:{l2}{row+2}", f"변경 후\n{new}",
                font=Font(name="맑은 고딕", size=10 if compact else 11, bold=True, color=BLUE),
                fill=PatternFill("solid", fgColor=LIGHT_BLUE), alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=border_all())
    merge_write(ws, f"{l1}{row+3}:{l2}{row+3}", change,
                font=Font(name="맑은 고딕", size=10 if compact else 11, bold=True, color=color),
                fill=PatternFill("solid", fgColor=bg), alignment=Alignment(horizontal="center", vertical="center"), border=border_all())


def write_single_customer_sheet(ws, customer: CustomerData, title: str, consultation_date: date, consultant: str) -> None:
    for col in range(1, 13):
        ws.column_dimensions[get_column_letter(col)].width = 11.5
    merge_write(ws, "A1:I3", title, font=Font(name="맑은 고딕", size=20, bold=True, color=NAVY_DARK), alignment=Alignment(horizontal="left", vertical="center"))
    merge_write(ws, "J1:L3", customer.proposal_label, font=Font(name="맑은 고딕", size=12, bold=True, color=WHITE), fill=PatternFill("solid", fgColor=TEAL), alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=border_all())
    ws.row_dimensions[1].height = 24; ws.row_dimensions[2].height = 24; ws.row_dimensions[3].height = 24
    if customer.priorities:
        merge_write(ws, "A4:L4", f"고객 우선사항  |  {' · '.join(customer.priorities[:2])}", font=Font(name="맑은 고딕", size=10, bold=True, color=NAVY), fill=PatternFill("solid", fgColor=LIGHT_GOLD), alignment=Alignment(horizontal="left", vertical="center"), border=border_all())
    write_metric_box(ws, 1, 4, 6, "월 보험료 변화", format_change(customer.analysis.monthly_delta), f"{customer.old_monthly:,}원 → {customer.new_monthly:,}원", customer.analysis.price_direction, False)
    write_metric_box(ws, 5, 8, 6, "연간 보험료 변화", format_change(customer.analysis.annual_delta), "월 차액 × 12개월", customer.analysis.price_direction, False)
    td = customer.analysis.total_delta
    td_dir = "동일" if td is None or td == 0 else "감소" if td < 0 else "증가"
    write_metric_box(ws, 9, 12, 6, "총 납입액 변화", format_change(td), f"{format_compact_won(customer.old_total)} → {format_compact_won(customer.new_total)}", td_dir, False)
    for r, h in {6:20, 7:25, 8:25, 9:22}.items(): ws.row_dimensions[r].height = h
    rows = list(prioritized_changes(customer.changes, customer.priorities, 4).iterrows())
    start_rows = [11, 16]
    for n, (_, item) in enumerate(rows):
        row = start_rows[n // 2]
        start_col = 1 if n % 2 == 0 else 7
        end_col = 6 if n % 2 == 0 else 12
        write_change_box(ws, start_col, end_col, row, item, False)
    for r in range(11, 20): ws.row_dimensions[r].height = 23
    dirs = customer.contracts["처리 방향 [목록 선택]"].fillna("").astype(str) if not customer.contracts.empty else pd.Series(dtype=str)
    keep = int(dirs.str.contains("유지").sum()); adjust = int(dirs.str.contains("감액|조정|해지|결정").sum()); check = int(dirs.str.contains("확인").sum())
    merge_write(ws, "A21:L21", f"계약 정리 요약   유지 {keep}건   |   조정·검토 {adjust}건   |   추가 확인 {check}건", font=Font(name="맑은 고딕", size=11, bold=True, color=NAVY), fill=PatternFill("solid", fgColor=LIGHT_BLUE), alignment=Alignment(horizontal="center", vertical="center"), border=border_all())
    merge_write(ws, "A23:L24", customer.headline, font=Font(name="맑은 고딕", size=12, bold=True, color=NAVY_DARK), fill=PatternFill("solid", fgColor=LIGHT_GRAY), alignment=Alignment(horizontal="left", vertical="center", wrap_text=True), border=Border(left=MEDIUM_NAVY))
    merge_write(ws, "A26:L26", f"상담일 {consultation_date:%Y.%m.%d}    담당자 {consultant or '-'}", font=Font(name="맑은 고딕", size=9.5, color=GRAY_DARK), alignment=Alignment(horizontal="right", vertical="center"))
    ws.print_area = "A1:L26"
    page_setup(ws, 1)


def write_two_customer_sheet(ws, customers: list[CustomerData], title: str, consultation_date: date, consultant: str) -> None:
    for col in range(1, 17):
        ws.column_dimensions[get_column_letter(col)].width = 9.4
    merge_write(ws, "A1:M3", title, font=Font(name="맑은 고딕", size=18, bold=True, color=NAVY_DARK), alignment=Alignment(horizontal="left", vertical="center"))
    family_delta = sum(c.analysis.monthly_delta for c in customers)
    family_direction = "감소" if family_delta < 0 else "증가" if family_delta > 0 else "동일"
    color, bg = excel_direction(family_direction)
    merge_write(ws, "N1:P3", f"가족 월 보험료\n{format_change(family_delta)}", font=Font(name="맑은 고딕", size=11, bold=True, color=color), fill=PatternFill("solid", fgColor=bg), alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=border_all())
    for i, customer in enumerate(customers):
        start = 1 if i == 0 else 9
        end = 8 if i == 0 else 16
        l1, l2 = get_column_letter(start), get_column_letter(end)
        merge_write(ws, f"{l1}5:{get_column_letter(end-3)}6", customer.name, font=Font(name="맑은 고딕", size=15, bold=True, color=NAVY), fill=PatternFill("solid", fgColor=WHITE), alignment=Alignment(horizontal="left", vertical="center"), border=border_all())
        merge_write(ws, f"{get_column_letter(end-2)}5:{l2}6", customer.proposal_label, font=Font(name="맑은 고딕", size=9.5, bold=True, color=WHITE), fill=PatternFill("solid", fgColor=TEAL), alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=border_all())
        if customer.priorities:
            merge_write(ws, f"{l1}7:{l2}7", f"우선사항  {' · '.join(customer.priorities[:2])}", font=Font(name="맑은 고딕", size=9.5, bold=True, color=NAVY), fill=PatternFill("solid", fgColor=LIGHT_GOLD), alignment=Alignment(horizontal="left", vertical="center"), border=border_all())
        write_metric_box(ws, start, start+1, 9, "월 보험료", format_change(customer.analysis.monthly_delta), f"{customer.old_monthly:,} → {customer.new_monthly:,}", customer.analysis.price_direction, True)
        write_metric_box(ws, start+2, start+4, 9, "연간 변화", format_change(customer.analysis.annual_delta), "월 차액 × 12", customer.analysis.price_direction, True)
        td = customer.analysis.total_delta; td_dir = "동일" if td is None or td == 0 else "감소" if td < 0 else "증가"
        write_metric_box(ws, start+5, end, 9, "총 납입액", format_change(td), f"{format_compact_won(customer.old_total)} → {format_compact_won(customer.new_total)}", td_dir, True)
        rows = list(prioritized_changes(customer.changes, customer.priorities, 3).iterrows())
        for n, (_, item) in enumerate(rows):
            write_change_box(ws, start, end, 14 + n*4, item, True)
        dirs = customer.contracts["처리 방향 [목록 선택]"].fillna("").astype(str) if not customer.contracts.empty else pd.Series(dtype=str)
        keep = int(dirs.str.contains("유지").sum()); adjust = int(dirs.str.contains("감액|조정|해지|결정").sum()); check = int(dirs.str.contains("확인").sum())
        merge_write(ws, f"{l1}27:{l2}27", f"계약 정리  유지 {keep}건 · 조정·검토 {adjust}건 · 확인 {check}건", font=Font(name="맑은 고딕", size=9.5, bold=True, color=NAVY), fill=PatternFill("solid", fgColor=LIGHT_BLUE), alignment=Alignment(horizontal="center", vertical="center"), border=border_all())
        merge_write(ws, f"{l1}29:{l2}31", customer.headline, font=Font(name="맑은 고딕", size=10.5, bold=True, color=NAVY_DARK), fill=PatternFill("solid", fgColor=LIGHT_GRAY), alignment=Alignment(horizontal="left", vertical="center", wrap_text=True), border=Border(left=MEDIUM_NAVY))
    merge_write(ws, "A33:P33", f"상담일 {consultation_date:%Y.%m.%d}    담당자 {consultant or '-'}", font=Font(name="맑은 고딕", size=9, color=GRAY_DARK), alignment=Alignment(horizontal="right", vertical="center"))
    for r in range(1, 34):
        if ws.row_dimensions[r].height is None:
            ws.row_dimensions[r].height = 21
    ws.print_area = "A1:P33"
    page_setup(ws, 1)


def write_contract_sheet(ws, customer: CustomerData) -> None:
    for col, width in {"A":22, "B":20, "C":42, "D":42}.items():
        ws.column_dimensions[col].width = width
    merge_write(ws, "A1:D3", f"{customer.name}님 기존 계약 정리 및 확인사항", font=Font(name="맑은 고딕", size=19, bold=True, color=NAVY_DARK), alignment=Alignment(horizontal="left", vertical="center"))
    headers = ["기존 계약/보장", "처리 방향", "판단 근거", "진행 조건"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(5, c, h)
        cell.font = Font(name="맑은 고딕", size=11.5, bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_all()
    row = 6
    contracts = customer.contracts if not customer.contracts.empty else default_contract_df().iloc[0:0]
    for _, item in contracts.iterrows():
        values = [normalize_text(item.get("기존 계약/보장")), normalize_text(item.get("처리 방향 [목록 선택]")), normalize_text(item.get("판단 근거")), normalize_text(item.get("진행 조건"))]
        for c, value in enumerate(values, 1):
            cell = ws.cell(row, c, value)
            cell.font = Font(name="맑은 고딕", size=11, bold=c == 2, color=NAVY if c == 2 else BLACK)
            cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE if c == 2 else WHITE)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border = border_all()
        ws.row_dimensions[row].height = 42
        row += 1
    ws.print_area = f"A1:D{max(row, 7)}"
    page_setup(ws, 1)


def create_excel(customers: list[CustomerData], title: str, consultation_date: date, consultant: str, include_contract_sheets: bool) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "가족_비교안" if len(customers) == 2 else f"{customers[0].name[:20]}_비교안"
    if len(customers) == 1:
        write_single_customer_sheet(ws, customers[0], title, consultation_date, consultant)
    else:
        write_two_customer_sheet(ws, customers, title, consultation_date, consultant)
    if include_contract_sheets:
        for customer in customers:
            if not customer.contracts.empty:
                sheet_name = f"{customer.name}_계약정리"[:31]
                ws2 = wb.create_sheet(sheet_name)
                write_contract_sheet(ws2, customer)
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ---------- example and UI ----------
def load_example(person_count: int) -> None:
    names = ["홍길동", "김영희"]
    for i in range(1, person_count + 1):
        st.session_state[f"rv2_name_{i}"] = names[i-1]
        st.session_state[f"rv2_old_month_{i}"] = "694,580" if i == 1 else "495,470"
        st.session_state[f"rv2_new_month_{i}"] = "367,707" if i == 1 else "329,210"
        st.session_state[f"rv2_old_total_{i}"] = "144,940,000" if i == 1 else "102,950,000"
        st.session_state[f"rv2_new_total_{i}"] = "88,240,000" if i == 1 else "79,010,000"
        st.session_state[f"rv2_priorities_{i}"] = ["월 보험료 부담", "암 치료비"] if i == 1 else ["암 진단비", "뇌·심장 보장"]
        st.session_state[f"rv2_proposal_{i}"] = "균형 보장형 추천안"
        st.session_state[f"rv2_changes_{i}"] = pd.DataFrame([
            {"변경할 보장":"암 진단비", "기존에는":"2,000만원", "변경 후에는":"5,000만원", "어떻게 달라지나요? [목록 선택]":"보장금액 증가", "왜 바꾸나요?":"진단 후 필요한 자금 준비", "변경 후 월 보험료":"80,000", "첫 장 표시 [목록 선택]":"핵심으로 표시"},
            {"변경할 보장":"암 주요치료비", "기존에는":"없음", "변경 후에는":"신규 구성", "어떻게 달라지나요? [목록 선택]":"새로 추가", "왜 바꾸나요?":"치료 과정의 비용 보완", "변경 후 월 보험료":"50,000", "첫 장 표시 [목록 선택]":"핵심으로 표시"},
            {"변경할 보장":"기존 실손보험", "기존에는":"가입 중", "변경 후에는":"유지", "어떻게 달라지나요? [목록 선택]":"그대로 유지", "왜 바꾸나요?":"기존 가입 조건 유지", "변경 후 월 보험료":"", "첫 장 표시 [목록 선택]":"상세에만 표시"},
        ])
        st.session_state[f"rv2_contracts_{i}"] = pd.DataFrame([
            {"기존 계약/보장":"기존 실손보험", "처리 방향 [목록 선택]":"유지", "판단 근거":"기존 가입 조건 유지", "진행 조건":"계속 유지"},
            {"기존 계약/보장":"기존 종합보험", "처리 방향 [목록 선택]":"신규 승인 후 결정", "판단 근거":"중복 및 부족 보장 재검토", "진행 조건":"신규 계약 승인 후"},
        ])
    st.session_state["rv2_example"] = True


def init_state() -> None:
    for i in [1, 2]:
        st.session_state.setdefault(f"rv2_changes_{i}", default_changes_df())
        st.session_state.setdefault(f"rv2_contracts_{i}", default_contract_df())


def collect_customer(i: int) -> CustomerData:
    name = normalize_text(st.session_state.get(f"rv2_name_{i}", ""))
    old_monthly = parse_money(st.session_state.get(f"rv2_old_month_{i}", ""))
    new_monthly = parse_money(st.session_state.get(f"rv2_new_month_{i}", ""))
    use_total = bool(st.session_state.get(f"rv2_use_total_{i}", True))
    old_total = parse_money(st.session_state.get(f"rv2_old_total_{i}", "")) if use_total else None
    new_total = parse_money(st.session_state.get(f"rv2_new_total_{i}", "")) if use_total else None
    if use_total and old_total == 0 and new_total == 0:
        old_total = new_total = None
    changes = clean_changes(st.session_state.get(f"rv2_changes_{i}"))
    contracts = clean_contracts(st.session_state.get(f"rv2_contracts_{i}"))
    priorities = list(st.session_state.get(f"rv2_priorities_{i}", []))[:2]
    analysis = analyze(old_monthly, new_monthly, old_total, new_total, changes)
    proposal = proposal_label_from_state(i, analysis, priorities)
    temp = CustomerData(i, name, proposal, priorities, old_monthly, new_monthly, old_total, new_total, changes, contracts, analysis, "")
    candidates = headline_candidates(temp)
    selected_idx = int(st.session_state.get(f"rv2_headline_idx_{i}", 0))
    headline = normalize_text(st.session_state.get(f"rv2_headline_custom_{i}", "")) or candidates[min(selected_idx, len(candidates)-1)]
    temp.headline = headline
    return temp


def render_customer_inputs(i: int, person_count: int) -> None:
    label = f"고객 {i}"
    st.markdown(f"### {label}")
    c1, c2 = st.columns(2)
    with c1:
        st.text_input("고객명", key=f"rv2_name_{i}", placeholder="예: 홍길동")
        priorities = st.multiselect(
            "고객이 가장 중요하게 생각하는 부분 · 최대 2개",
            PRIORITY_OPTIONS,
            max_selections=2,
            key=f"rv2_priorities_{i}",
            help="핵심 변경 순서, 제안 유형 추천, 우선사항 배지와 추천 문구에 반영됩니다.",
        )
    with c2:
        # recommendation is shown using currently available inputs
        current = collect_customer(i)
        suggested = suggest_proposal(priorities, current.analysis)
        st.caption(f"추천 제안 유형: **{suggested}**")
        selected = st.selectbox("제안 유형 [목록 선택] ▼", PROPOSAL_OPTIONS, key=f"rv2_proposal_{i}")
        if selected == "직접 입력":
            st.text_input("제안 유형 직접 입력", key=f"rv2_proposal_custom_{i}", placeholder="예: 실손 유지 + 진단비 보완안")
    if person_count == 2 and i == 2:
        st.caption("고객 2는 고객 1과 별도의 제안 유형과 우선사항을 적용할 수 있습니다.")


def run() -> None:
    st.title(f"🔁 {APP_TITLE}")
    st.caption("고객정보 → 보험료 비교 → 핵심 변경 → 고객 시점 미리보기·엑셀")
    init_state()

    person_count = int(st.selectbox("비교 대상 인원 [목록 선택] ▼", [1, 2], format_func=lambda x: f"{x}명", key="rv2_person_count"))
    top1, top2 = st.columns([1, 4])
    with top1:
        if st.button("예시로 먼저 보기", use_container_width=True):
            load_example(person_count)
            st.rerun()
    with top2:
        if st.session_state.get("rv2_example"):
            st.info("예시 데이터가 입력되어 있습니다. 실제 고객 정보로 바꾸어 사용하세요.")

    tabs = st.tabs(["1. 공통·고객정보", "2. 보험료 비교", "3. 핵심 변경 내용", "4. 미리보기·엑셀"])

    with tabs[0]:
        st.subheader("공통 정보")
        common1, common2 = st.columns(2)
        with common1:
            consultation_date = st.date_input("상담일", value=date.today(), key="rv2_date")
            consultant = st.text_input("담당자", key="rv2_consultant", placeholder="예: 박병선 팀장")
        with common2:
            include_contract = st.checkbox("계약 상세 시트 포함", value=False, key="rv2_include_contract")
            st.caption("계약 처리 내용이 있을 때만 고객별 상세 시트가 추가됩니다.")
        st.divider()
        for i in range(1, person_count + 1):
            render_customer_inputs(i, person_count)
            if i < person_count:
                st.divider()
        names = [normalize_text(st.session_state.get(f"rv2_name_{i}", "")) for i in range(1, person_count + 1)]
        auto = auto_title(names)
        title = st.text_input("자료 제목", key="rv2_title", placeholder=auto)
        if names and all(names):
            st.caption(f"비워두면 **‘{auto}’** 형식으로 자동 생성됩니다.")
        else:
            st.caption("고객명을 입력하면 자료 제목이 자동 생성됩니다.")

    with tabs[1]:
        for i in range(1, person_count + 1):
            name = normalize_text(st.session_state.get(f"rv2_name_{i}", "")) or f"고객 {i}"
            st.subheader(f"{name} 보험료 비교")
            c1, c2 = st.columns(2)
            with c1:
                _money_input("기존 월 보험료", f"rv2_old_month_{i}")
                use_total = st.checkbox("총 납입액도 비교하기", value=True, key=f"rv2_use_total_{i}")
                _money_input("기존 납입예정 총액", f"rv2_old_total_{i}", disabled=not use_total)
            with c2:
                _money_input("변경 후 월 보험료", f"rv2_new_month_{i}")
                _money_input("변경 후 납입예정 총액", f"rv2_new_total_{i}", disabled=not use_total)
            old_m = parse_money(st.session_state.get(f"rv2_old_month_{i}", "")); new_m = parse_money(st.session_state.get(f"rv2_new_month_{i}", ""))
            if old_m or new_m:
                st.success(f"월 보험료 {format_change(new_m-old_m)} · 연간 {format_change((new_m-old_m)*12)}")
            if i < person_count: st.divider()

    with tabs[2]:
        st.info("고객에게 설명할 중요한 변경만 입력하세요. 모든 담보를 입력할 필요는 없습니다.")
        st.markdown("**흰색 칸은 직접 입력**, 제목에 **[목록 선택] ▼**가 있는 칸은 목록에서 선택합니다.")
        for i in range(1, person_count + 1):
            name = normalize_text(st.session_state.get(f"rv2_name_{i}", "")) or f"고객 {i}"
            st.subheader(f"{name} 핵심 변경")
            with st.form(f"rv2_changes_form_{i}", clear_on_submit=False):
                edited = st.data_editor(
                    st.session_state[f"rv2_changes_{i}"], num_rows="dynamic", use_container_width=True, hide_index=True,
                    column_config={
                        "변경할 보장": st.column_config.TextColumn("변경할 보장", required=True, width="medium", help="고객 자료에는 입력한 담보명을 그대로 사용합니다."),
                        "기존에는": st.column_config.TextColumn("기존에는", width="medium", help="예: 2,000만원 / 없음 / 가입 중"),
                        "변경 후에는": st.column_config.TextColumn("변경 후에는", width="medium", help="예: 5,000만원 / 신규 구성 / 유지"),
                        "어떻게 달라지나요? [목록 선택]": st.column_config.SelectboxColumn("어떻게 달라지나요? [목록 선택] ▼", options=CHANGE_OPTIONS, required=True, width="medium"),
                        "왜 바꾸나요?": st.column_config.TextColumn("왜 바꾸나요?", width="large"),
                        "변경 후 월 보험료": st.column_config.TextColumn("변경 후 월 보험료", width="small"),
                        "첫 장 표시 [목록 선택]": st.column_config.SelectboxColumn("첫 장 표시 [목록 선택] ▼", options=DISPLAY_OPTIONS, required=True, width="medium"),
                    }, key=f"rv2_changes_editor_{i}",
                )
                submit = st.form_submit_button("핵심 변경 내용 적용", type="primary", use_container_width=True)
            if submit:
                st.session_state[f"rv2_changes_{i}"] = edited
                st.success("핵심 변경 내용을 적용했습니다.")
            with st.expander("기존 계약 처리 방향 입력 · 선택사항"):
                with st.form(f"rv2_contract_form_{i}", clear_on_submit=False):
                    cedit = st.data_editor(
                        st.session_state[f"rv2_contracts_{i}"], num_rows="dynamic", use_container_width=True, hide_index=True,
                        column_config={
                            "기존 계약/보장": st.column_config.TextColumn("기존 계약/보장", required=True, width="large"),
                            "처리 방향 [목록 선택]": st.column_config.SelectboxColumn("처리 방향 [목록 선택] ▼", options=CONTRACT_OPTIONS, required=True),
                            "판단 근거": st.column_config.TextColumn("판단 근거", width="large"),
                            "진행 조건": st.column_config.TextColumn("진행 조건", width="large"),
                        }, key=f"rv2_contract_editor_{i}",
                    )
                    csubmit = st.form_submit_button("계약 처리 내용 적용", use_container_width=True)
                if csubmit:
                    st.session_state[f"rv2_contracts_{i}"] = cedit
                    st.success("계약 처리 내용을 적용했습니다.")
            if i < person_count: st.divider()

    customers = [collect_customer(i) for i in range(1, person_count + 1)]
    names = [c.name for c in customers]
    title = normalize_text(st.session_state.get("rv2_title", "")) or auto_title(names)
    consultation_date = st.session_state.get("rv2_date", date.today())
    consultant = normalize_text(st.session_state.get("rv2_consultant", ""))
    include_contract = bool(st.session_state.get("rv2_include_contract", False))

    with tabs[3]:
        st.subheader("고객 시점 미리보기")
        for customer in customers:
            candidates = headline_candidates(customer)
            st.markdown(f"**{customer.name or f'고객 {customer.index}'} 추천 멘트 선택**")
            choice = st.selectbox("추천 멘트", candidates, key=f"rv2_headline_select_{customer.index}", label_visibility="collapsed")
            st.session_state[f"rv2_headline_custom_{customer.index}"] = st.text_input("추천 멘트 직접 수정", value=choice, key=f"rv2_headline_edit_{customer.index}")
        customers = [collect_customer(i) for i in range(1, person_count + 1)]
        st.markdown(f"## {title}")
        if person_count == 2:
            family_delta = sum(c.analysis.monthly_delta for c in customers)
            st.markdown(f"<div style='padding:11px 14px;background:#EAF2F8;border-radius:9px;color:#17365D;font-weight:800;margin-bottom:12px;'>가족 월 보험료 변화: {format_change(family_delta)}</div>", unsafe_allow_html=True)
            cols = st.columns(2)
            for col, customer in zip(cols, customers):
                with col:
                    render_customer_preview(customer, True)
        else:
            render_customer_preview(customers[0], False)
        all_warnings: list[str] = []
        for customer in customers:
            all_warnings.extend(validate_customer(customer, 3 if person_count == 2 else 4))
        if all_warnings:
            with st.expander("출력 전 확인사항", expanded=True):
                for warning in all_warnings:
                    st.warning(warning)
        else:
            st.success("출력 준비가 완료되었습니다.")
        if person_count == 2 and any(not c.name for c in customers):
            st.error("2명 비교를 선택했으므로 고객 1과 고객 2의 이름을 모두 입력해 주세요.")
        else:
            excel = create_excel(customers, title, consultation_date, consultant, include_contract)
            filename = safe_filename(f"{title}_{consultation_date:%Y%m%d}.xlsx")
            st.download_button("엑셀 다운로드", data=excel, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)


if __name__ == "__main__":
    run()
