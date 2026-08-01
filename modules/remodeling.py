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
from .ui_components import page_header

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
    additions: pd.DataFrame
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
        },
        {
            "변경할 보장": "", "기존에는": "", "변경 후에는": "",
            "어떻게 달라지나요? [목록 선택]": "선택하세요 ▼", "왜 바꾸나요?": "",
        },
    ])


def default_contract_df() -> pd.DataFrame:
    return pd.DataFrame([
        {
            "보험회사": "", "상품명": "",
            "처리 방향 [목록 선택]": "선택하세요 ▼",
            "구체적인 변경 내용": ""
        }
    ])


def default_additions_df() -> pd.DataFrame:
    """1페이지에 표시할 신규 가입 보험 요약: 대표 보장 내용과 월납 금액만 입력."""
    return pd.DataFrame([
        {"대표 보장 내용": "", "월납 보험료": ""},
        {"대표 보장 내용": "", "월납 보험료": ""},
        {"대표 보장 내용": "", "월납 보험료": ""},
        {"대표 보장 내용": "", "월납 보험료": ""},
    ])


def clean_changes(df: pd.DataFrame | None) -> pd.DataFrame:
    if df is None or df.empty:
        return default_changes_df().iloc[0:0]
    result = df.copy()
    result["변경할 보장"] = result["변경할 보장"].map(normalize_text)
    result = result[result["변경할 보장"] != ""].reset_index(drop=True)
    col = "어떻게 달라지나요? [목록 선택]"
    if col in result:
        result[col] = result[col].replace("선택하세요 ▼", "")
    return result


def clean_contracts(df: pd.DataFrame | None) -> pd.DataFrame:
    if df is None or df.empty:
        return default_contract_df().iloc[0:0]
    result = df.copy()
    for col in ["보험회사", "상품명", "처리 방향 [목록 선택]", "구체적인 변경 내용"]:
        if col not in result.columns:
            result[col] = ""
        result[col] = result[col].map(normalize_text)
    result = result[(result["보험회사"] != "") | (result["상품명"] != "")].reset_index(drop=True)
    result["처리 방향 [목록 선택]"] = result["처리 방향 [목록 선택]"].replace("선택하세요 ▼", "")
    return result[["보험회사", "상품명", "처리 방향 [목록 선택]", "구체적인 변경 내용"]]


def clean_additions(df: pd.DataFrame | None) -> pd.DataFrame:
    if df is None or df.empty:
        return default_additions_df().iloc[0:0]
    result = df.copy()
    # 구버전 열 이름도 읽을 수 있게 호환 처리
    if "월납 보험료" not in result.columns and "월 보험료" in result.columns:
        result["월납 보험료"] = result["월 보험료"]
    for col in ["대표 보장 내용", "월납 보험료"]:
        if col not in result.columns:
            result[col] = ""
        result[col] = result[col].map(normalize_text)
    result = result[result["대표 보장 내용"] != ""].reset_index(drop=True)
    return result[["대표 보장 내용", "월납 보험료"]]


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
    if not customer.additions.empty:
        st.markdown("**새롭게 가입하는 보험**")
        preview_df = customer.additions.copy()
        preview_df["월납 보험료"] = preview_df["월납 보험료"].map(lambda v: f"{parse_money(v):,}원" if parse_money(v) else "-")
        st.dataframe(preview_df[["대표 보장 내용", "월납 보험료"]], hide_index=True, use_container_width=True)
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


def _write_title(ws, title: str, max_col: int, subtitle: str) -> None:
    last = get_column_letter(max_col)
    merge_write(ws, f"A1:{last}2", title,
                font=Font(name="맑은 고딕", size=24, bold=True, color=NAVY_DARK),
                alignment=Alignment(horizontal="center", vertical="center"))
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 30
    merge_write(ws, f"A3:{last}3", subtitle,
                font=Font(name="맑은 고딕", size=11, bold=True, color=NAVY),
                fill=PatternFill("solid", fgColor=LIGHT_BLUE),
                alignment=Alignment(horizontal="center", vertical="center", wrap_text=True))
    ws.row_dimensions[3].height = 24


def _write_premium_table(ws, start_col: int, end_col: int, top_row: int, customer: CustomerData) -> None:
    mid = (start_col + end_col) // 2
    spans = [(start_col, start_col+1), (start_col+2, mid), (mid+1, mid+2), (mid+3, end_col)]
    labels = ["기존 월 보험료", f"{customer.old_monthly:,}원", "리모델링 후", f"{customer.new_monthly:,}원"]
    for idx, ((a,b), value) in enumerate(zip(spans, labels)):
        merge_write(ws, f"{get_column_letter(a)}{top_row}:{get_column_letter(b)}{top_row}", value,
                    font=Font(name="맑은 고딕", size=9.5 if idx%2==0 else 11, bold=True, color=BLUE if idx==3 else BLACK),
                    fill=PatternFill("solid", fgColor=LIGHT_BLUE if idx%2==0 else WHITE),
                    alignment=Alignment(horizontal="center", vertical="center"), border=border_all())
    old_total = format_compact_won(customer.old_total)
    new_total = format_compact_won(customer.new_total)
    labels2 = ["기존 납입 총액", old_total, "변경 납입 총액", new_total]
    for idx, ((a,b), value) in enumerate(zip(spans, labels2)):
        merge_write(ws, f"{get_column_letter(a)}{top_row+1}:{get_column_letter(b)}{top_row+1}", value,
                    font=Font(name="맑은 고딕", size=9.5 if idx%2==0 else 10.5, bold=True, color=BLUE if idx==3 else BLACK),
                    fill=PatternFill("solid", fgColor=LIGHT_BLUE if idx%2==0 else WHITE),
                    alignment=Alignment(horizontal="center", vertical="center"), border=border_all())


def _write_new_insurance_list(ws, start_col: int, end_col: int, top_row: int, customer: CustomerData, max_items: int) -> int:
    """1페이지에 대표 보장 내용과 월납 금액만 간결한 2열 표로 출력."""
    l1, l2 = get_column_letter(start_col), get_column_letter(end_col)
    merge_write(ws, f"{l1}{top_row}:{l2}{top_row}", "새롭게 가입하는 보험",
                font=Font(name="맑은 고딕", size=11, bold=True, color=NAVY),
                alignment=Alignment(horizontal="left", vertical="center"))
    cur = top_row + 1
    width = end_col - start_col + 1
    content_end = start_col + max(2, int(width * 0.68)) - 1
    content_end = min(content_end, end_col - 2)
    # 표 머리글
    merge_write(ws, f"{l1}{cur}:{get_column_letter(content_end)}{cur}", "대표 보장 내용",
                font=Font(name="맑은 고딕", size=9.5, bold=True, color=WHITE),
                fill=PatternFill("solid", fgColor=NAVY),
                alignment=Alignment(horizontal="center", vertical="center"), border=border_all())
    merge_write(ws, f"{get_column_letter(content_end+1)}{cur}:{l2}{cur}", "월납 보험료",
                font=Font(name="맑은 고딕", size=9.5, bold=True, color=WHITE),
                fill=PatternFill("solid", fgColor=NAVY),
                alignment=Alignment(horizontal="center", vertical="center"), border=border_all())
    cur += 1
    rows = list(customer.additions.head(max_items).iterrows())
    if not rows:
        rows = [(0, pd.Series({"대표 보장 내용":"입력된 신규 가입 내용이 없습니다.", "월납 보험료":""}))]
    total = 0
    for _, item in rows:
        content = normalize_text(item.get("대표 보장 내용")) or "-"
        premium = parse_money(item.get("월납 보험료"))
        total += premium
        premium_text = f"{premium:,}원" if premium else "-"
        merge_write(ws, f"{l1}{cur}:{get_column_letter(content_end)}{cur}", content,
                    font=Font(name="맑은 고딕", size=10, bold=True, color=NAVY),
                    fill=PatternFill("solid", fgColor=WHITE),
                    alignment=Alignment(horizontal="left", vertical="center", wrap_text=True), border=border_all())
        merge_write(ws, f"{get_column_letter(content_end+1)}{cur}:{l2}{cur}", premium_text,
                    font=Font(name="맑은 고딕", size=10.5, bold=True, color=BLUE),
                    fill=PatternFill("solid", fgColor=WHITE),
                    alignment=Alignment(horizontal="right", vertical="center"), border=border_all())
        ws.row_dimensions[cur].height = 28
        cur += 1
    if rows and customer.additions.shape[0] > 0:
        merge_write(ws, f"{l1}{cur}:{get_column_letter(content_end)}{cur}", "합계",
                    font=Font(name="맑은 고딕", size=10, bold=True, color=BLACK),
                    fill=PatternFill("solid", fgColor=LIGHT_GREEN),
                    alignment=Alignment(horizontal="center", vertical="center"), border=border_all())
        merge_write(ws, f"{get_column_letter(content_end+1)}{cur}:{l2}{cur}", f"{total:,}원",
                    font=Font(name="맑은 고딕", size=11, bold=True, color=GREEN),
                    fill=PatternFill("solid", fgColor=LIGHT_GREEN),
                    alignment=Alignment(horizontal="right", vertical="center"), border=border_all())
        cur += 1
    return cur


def _write_coverage_list(ws, start_col: int, end_col: int, top_row: int, customer: CustomerData, max_items: int) -> int:
    return _write_new_insurance_list(ws, start_col, end_col, top_row, customer, max_items)


def _write_saving_boxes(ws, start_col: int, end_col: int, top_row: int, customer: CustomerData) -> None:
    mid = (start_col + end_col)//2
    monthly = abs(customer.analysis.monthly_delta)
    total = abs(customer.analysis.total_delta or 0)
    direction = customer.analysis.price_direction
    color, bg = excel_direction(direction)
    merge_write(ws, f"{get_column_letter(start_col)}{top_row}:{get_column_letter(mid)}{top_row+1}",
                f"월 {'절감' if direction=='감소' else '증가' if direction=='증가' else '변동'}금액\n{monthly:,}원",
                font=Font(name="맑은 고딕", size=11, bold=True, color=color), fill=PatternFill("solid", fgColor=bg),
                alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=border_all())
    total_label = "총 납입 절감액" if (customer.analysis.total_delta or 0) < 0 else "총 납입 변동액"
    merge_write(ws, f"{get_column_letter(mid+1)}{top_row}:{get_column_letter(end_col)}{top_row+1}",
                f"{total_label}\n{format_compact_won(total)}",
                font=Font(name="맑은 고딕", size=11, bold=True, color=color), fill=PatternFill("solid", fgColor=bg),
                alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=border_all())


def write_single_customer_sheet(ws, customer: CustomerData, title: str, consultation_date: date, consultant: str) -> None:
    for col in range(1, 13): ws.column_dimensions[get_column_letter(col)].width = 11.5
    _write_title(ws, title, 12, "필요한 진단비와 주요 치료비 중심으로 보장을 재구성하고, 월 보험료와 총 납입 부담을 함께 비교한 제안입니다.")
    color, bg = excel_direction(customer.analysis.price_direction)
    pct = abs(customer.analysis.monthly_delta) / customer.old_monthly * 100 if customer.old_monthly else 0
    merge_write(ws, "A4:L5", f"월 {format_change(customer.analysis.monthly_delta)}   |   연간 {format_change(customer.analysis.annual_delta)}   |   월 보험료 약 {pct:.1f}% {'감소' if customer.analysis.monthly_delta<0 else '증가' if customer.analysis.monthly_delta>0 else '동일'}",
                font=Font(name="맑은 고딕", size=12, bold=True, color=color), fill=PatternFill("solid", fgColor=bg),
                alignment=Alignment(horizontal="center", vertical="center"), border=border_all())
    merge_write(ws, "A7:L8", customer.name,
                font=Font(name="맑은 고딕", size=17, bold=True, color=WHITE), fill=PatternFill("solid", fgColor=NAVY),
                alignment=Alignment(horizontal="center", vertical="center"), border=border_all())
    _write_premium_table(ws, 1, 12, 9, customer)
    next_row = _write_coverage_list(ws, 1, 12, 12, customer, 4)
    _write_saving_boxes(ws, 1, 12, max(next_row+1, 18), customer)
    r = max(next_row+4, 22)
    merge_write(ws, f"A{r}:L{r+1}", customer.headline,
                font=Font(name="맑은 고딕", size=11, bold=True, color=NAVY_DARK), fill=PatternFill("solid", fgColor=LIGHT_GRAY),
                alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=border_all())
    merge_write(ws, f"A{r+2}:L{r+2}", f"상담일 {consultation_date:%Y.%m.%d}    담당자 {consultant or '-'}",
                font=Font(name="맑은 고딕", size=9, color=GRAY_DARK), alignment=Alignment(horizontal="right", vertical="center"))
    ws.print_area = f"A1:L{r+2}"
    page_setup(ws, 1)


def write_two_customer_sheet(ws, customers: list[CustomerData], title: str, consultation_date: date, consultant: str) -> None:
    for col in range(1, 17): ws.column_dimensions[get_column_letter(col)].width = 9.5
    _write_title(ws, title, 16, "필요한 진단비와 주요 치료비 중심으로 보장을 재구성하고, 월 보험료와 총 납입 부담을 함께 낮추는 방향입니다.")
    family_month = sum(c.analysis.monthly_delta for c in customers)
    old_family = sum(c.old_monthly for c in customers)
    family_total = sum((c.analysis.total_delta or 0) for c in customers)
    pct = abs(family_month)/old_family*100 if old_family else 0
    color,bg=excel_direction("감소" if family_month<0 else "증가" if family_month>0 else "동일")
    merge_write(ws, "A4:P5", f"월 {format_change(family_month)}   |   총 납입예정액 {format_change(family_total)}   |   월 보험료 약 {pct:.1f}% {'감소' if family_month<0 else '증가' if family_month>0 else '동일'}",
                font=Font(name="맑은 고딕", size=12, bold=True, color=color), fill=PatternFill("solid", fgColor=bg),
                alignment=Alignment(horizontal="center", vertical="center"), border=border_all())
    for i,c in enumerate(customers):
        a,b=(1,8) if i==0 else (9,16)
        merge_write(ws, f"{get_column_letter(a)}7:{get_column_letter(b)}8", c.name,
                    font=Font(name="맑은 고딕", size=16, bold=True, color=WHITE), fill=PatternFill("solid", fgColor=NAVY),
                    alignment=Alignment(horizontal="center", vertical="center"), border=border_all())
        _write_premium_table(ws,a,b,9,c)
        _write_coverage_list(ws,a,b,12,c,4)
        _write_saving_boxes(ws,a,b,18,c)
    r=22
    headers=["부부 합산 비교","기존","리모델링 후","절감 효과"]
    spans=[(1,3),(4,7),(8,11),(12,16)]
    for (a,b),h in zip(spans,headers):
        merge_write(ws,f"{get_column_letter(a)}{r}:{get_column_letter(b)}{r}",h,font=Font(name="맑은 고딕",size=11,bold=True,color=WHITE),fill=PatternFill("solid",fgColor=NAVY),alignment=Alignment(horizontal="center",vertical="center"),border=border_all())
    old_total_month=sum(c.old_monthly for c in customers); new_total_month=sum(c.new_monthly for c in customers)
    old_total_all=sum(c.old_total or 0 for c in customers); new_total_all=sum(c.new_total or 0 for c in customers)
    rows=[("월 보험료",f"{old_total_month:,}원",f"{new_total_month:,}원",f"{abs(family_month):,}원 {'절감' if family_month<0 else '증가'} (약 {pct:.1f}%)"),
          ("연간 보험료",f"{old_total_month*12:,}원",f"{new_total_month*12:,}원",format_change(family_month*12)),
          ("납입예정 총액",format_compact_won(old_total_all),format_compact_won(new_total_all),format_change(family_total))]
    for rr,row in enumerate(rows,r+1):
        for (a,b),v in zip(spans,row):
            merge_write(ws,f"{get_column_letter(a)}{rr}:{get_column_letter(b)}{rr}",v,font=Font(name="맑은 고딕",size=10.5,bold=(a==1 or a==12),color=GREEN if a==12 and family_month<0 else BLACK),fill=PatternFill("solid",fgColor=LIGHT_BLUE if a==1 else LIGHT_GOLD if a==12 else WHITE),alignment=Alignment(horizontal="center",vertical="center",wrap_text=True),border=border_all())
    merge_write(ws,"A26:P27","단순한 보험료 축소가 아니라, 중복되거나 효율이 낮은 부담을 조정하고 핵심 보장 중심으로 재구성하는 제안입니다.",font=Font(name="맑은 고딕",size=10.5,bold=True,color=NAVY),alignment=Alignment(horizontal="center",vertical="center",wrap_text=True))
    merge_write(ws,"A28:P28",f"상담일 {consultation_date:%Y.%m.%d}    담당자 {consultant or '-'}",font=Font(name="맑은 고딕",size=9,color=GRAY_DARK),alignment=Alignment(horizontal="right",vertical="center"))
    ws.print_area="A1:P28"
    page_setup(ws,1)


def _table_header(ws,row:int,headers:list[str],spans:list[tuple[int,int]]) -> None:
    for h,(a,b) in zip(headers,spans):
        merge_write(ws,f"{get_column_letter(a)}{row}:{get_column_letter(b)}{row}",h,font=Font(name="맑은 고딕",size=9.5,bold=True,color=WHITE),fill=PatternFill("solid",fgColor=NAVY),alignment=Alignment(horizontal="center",vertical="center",wrap_text=True),border=border_all())


def write_detail_page(ws, customers:list[CustomerData], start_row:int, max_col:int, consultation_date:date, consultant:str) -> None:
    last = get_column_letter(max_col)
    merge_write(ws, f"A{start_row}:{last}{start_row+1}", "보장 변경 및 기존 계약 정리",
                font=Font(name="맑은 고딕", size=23, bold=True, color=NAVY_DARK),
                alignment=Alignment(horizontal="center", vertical="center"))
    ws.row_dimensions[start_row].height = 28
    ws.row_dimensions[start_row + 1].height = 28
    row = start_row + 3

    if max_col == 16:
        change_spans = [(1,3),(4,6),(7,9),(10,12),(13,16)]
        contract_spans = [(1,3),(4,6),(7,9),(10,16)]
    else:
        change_spans = [(1,2),(3,4),(5,7),(8,9),(10,12)]
        contract_spans = [(1,3),(4,6),(7,8),(9,12)]

    for c in customers:
        merge_write(ws, f"A{row}:{last}{row}", f"{c.name}님",
                    font=Font(name="맑은 고딕", size=13, bold=True, color=WHITE),
                    fill=PatternFill("solid", fgColor=TEAL),
                    alignment=Alignment(horizontal="left", vertical="center"), border=border_all())
        row += 1

        # 핵심 변경을 2페이지 상단에 배치
        merge_write(ws, f"A{row}:{last}{row}", "핵심 변경 내용",
                    font=Font(name="맑은 고딕", size=11, bold=True, color=NAVY),
                    fill=PatternFill("solid", fgColor=LIGHT_GREEN),
                    alignment=Alignment(horizontal="left", vertical="center"), border=border_all())
        row += 1
        _table_header(ws, row, ["보장·특약명","기존","변경 후","변화 유형","변경 이유"], change_spans)
        row += 1
        changes = c.changes if not c.changes.empty else default_changes_df().iloc[0:0]
        if changes.empty:
            merge_write(ws, f"A{row}:{last}{row}", "입력된 핵심 변경 내용이 없습니다.",
                        font=Font(name="맑은 고딕", size=10, color=GRAY_DARK),
                        alignment=Alignment(horizontal="center", vertical="center"), border=border_all())
            row += 1
        else:
            for _, item in changes.head(8).iterrows():
                vals = [normalize_text(item.get("변경할 보장")), normalize_text(item.get("기존에는")),
                        normalize_text(item.get("변경 후에는")), normalize_text(item.get("어떻게 달라지나요? [목록 선택]")),
                        normalize_text(item.get("왜 바꾸나요?"))]
                for (a,b), value in zip(change_spans, vals):
                    merge_write(ws, f"{get_column_letter(a)}{row}:{get_column_letter(b)}{row}", value,
                                font=Font(name="맑은 고딕", size=9.5, bold=(a==1), color=NAVY if a==1 else BLACK),
                                fill=PatternFill("solid", fgColor=WHITE),
                                alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=border_all())
                ws.row_dimensions[row].height = 34
                row += 1

        row += 1
        merge_write(ws, f"A{row}:{last}{row}", "기존 보험의 변경 방향",
                    font=Font(name="맑은 고딕", size=11, bold=True, color=NAVY),
                    fill=PatternFill("solid", fgColor=LIGHT_BLUE),
                    alignment=Alignment(horizontal="left", vertical="center"), border=border_all())
        row += 1
        _table_header(ws, row, ["보험회사","상품명","처리 방향","구체적인 변경 내용"], contract_spans)
        row += 1
        contracts = c.contracts if not c.contracts.empty else default_contract_df().iloc[0:0]
        if contracts.empty:
            merge_write(ws, f"A{row}:{last}{row}", "입력된 기존 보험 변경 방향이 없습니다.",
                        font=Font(name="맑은 고딕", size=10, color=GRAY_DARK),
                        alignment=Alignment(horizontal="center", vertical="center"), border=border_all())
            row += 1
        else:
            for _, item in contracts.head(8).iterrows():
                vals = [normalize_text(item.get("보험회사")), normalize_text(item.get("상품명")),
                        normalize_text(item.get("처리 방향 [목록 선택]")), normalize_text(item.get("구체적인 변경 내용"))]
                for (a,b), value in zip(contract_spans, vals):
                    merge_write(ws, f"{get_column_letter(a)}{row}:{get_column_letter(b)}{row}", value,
                                font=Font(name="맑은 고딕", size=9.5, bold=(a==7), color=NAVY if a==7 else BLACK),
                                fill=PatternFill("solid", fgColor=WHITE),
                                alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=border_all())
                ws.row_dimensions[row].height = 36
                row += 1
        row += 2

    merge_write(ws, f"A{row}:{last}{row}",
                "※ 기존 계약의 감액·해지 등은 새로운 계약의 승인 조건과 보장 개시 여부를 확인한 후 결정해야 합니다.",
                font=Font(name="맑은 고딕", size=9, color=GRAY_DARK),
                fill=PatternFill("solid", fgColor=LIGHT_ORANGE),
                alignment=Alignment(horizontal="left", vertical="center", wrap_text=True), border=border_all())
    merge_write(ws, f"A{row+2}:{last}{row+2}",
                f"상담일 {consultation_date:%Y.%m.%d}    담당자 {consultant or '-'}",
                font=Font(name="맑은 고딕", size=9, color=GRAY_DARK),
                alignment=Alignment(horizontal="right", vertical="center"))
    ws.print_area = f"A1:{last}{row+2}"
    page_setup(ws, 1)


def create_excel(customers: list[CustomerData], title: str, consultation_date: date, consultant: str) -> BytesIO:
    wb = Workbook()
    summary_ws = wb.active
    summary_ws.title = "리모델링 비교안"
    if len(customers) == 1:
        write_single_customer_sheet(summary_ws, customers[0], title, consultation_date, consultant)
        detail_cols = 12
    else:
        write_two_customer_sheet(summary_ws, customers, title, consultation_date, consultant)
        detail_cols = 16

    detail_ws = wb.create_sheet("보장·계약 변경")
    for col in range(1, detail_cols + 1):
        detail_ws.column_dimensions[get_column_letter(col)].width = 11 if detail_cols == 12 else 9.5
    write_detail_page(detail_ws, customers, 1, detail_cols, consultation_date, consultant)

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
            {"변경할 보장":"암 진단비", "기존에는":"2,000만원", "변경 후에는":"5,000만원", "어떻게 달라지나요? [목록 선택]":"보장금액 증가", "왜 바꾸나요?":"진단 후 필요한 자금 준비"},
            {"변경할 보장":"암 주요치료비", "기존에는":"없음", "변경 후에는":"신규 구성", "어떻게 달라지나요? [목록 선택]":"새로 추가", "왜 바꾸나요?":"치료 과정의 비용 보완"},
            {"변경할 보장":"기존 실손보험", "기존에는":"가입 중", "변경 후에는":"유지", "어떻게 달라지나요? [목록 선택]":"그대로 유지", "왜 바꾸나요?":"기존 가입 조건 유지"},
        ])
        st.session_state[f"rv2_contracts_{i}"] = pd.DataFrame([
            {"보험회사":"기존보험사", "상품명":"기존 종합보험", "처리 방향 [목록 선택]":"일부 특약 조정", "구체적인 변경 내용":"중복 특약은 조정하고 유지 가치가 있는 보장은 남깁니다."},
            {"보험회사":"기존보험사", "상품명":"실손의료보험", "처리 방향 [목록 선택]":"유지", "구체적인 변경 내용":"현재 가입 조건을 유지합니다."},
        ])
        st.session_state[f"rv2_additions_{i}"] = pd.DataFrame([
            {"대표 보장 내용":"암·뇌·심장 진단비", "월납 보험료":"128,589"},
            {"대표 보장 내용":"암 주요치료비", "월납 보험료":"111,255"},
            {"대표 보장 내용":"운전자보험", "월납 보험료":"15,000"},
            {"대표 보장 내용":"순환계 주요치료비", "월납 보험료":"112,863"},
        ])
    st.session_state["rv2_example"] = True


def init_state() -> None:
    for i in [1, 2]:
        st.session_state.setdefault(f"rv2_changes_{i}", default_changes_df())
        st.session_state.setdefault(f"rv2_contracts_{i}", default_contract_df())
        st.session_state.setdefault(f"rv2_additions_{i}", default_additions_df())


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
    additions = clean_additions(st.session_state.get(f"rv2_additions_{i}"))
    priorities = list(st.session_state.get(f"rv2_priorities_{i}", []))[:2]
    analysis = analyze(old_monthly, new_monthly, old_total, new_total, changes)
    proposal = proposal_label_from_state(i, analysis, priorities)
    temp = CustomerData(i, name, proposal, priorities, old_monthly, new_monthly, old_total, new_total, changes, contracts, additions, analysis, "")
    candidates = headline_candidates(temp)
    selected_idx = int(st.session_state.get(f"rv2_headline_idx_{i}", 0))
    headline = normalize_text(st.session_state.get(f"rv2_headline_custom_{i}", "")) or candidates[min(selected_idx, len(candidates)-1)]
    temp.headline = headline
    return temp



def _ensure_change_row_state(i: int) -> int:
    """Prepare plain Streamlit widget state for editable core-change rows."""
    df = st.session_state.get(f"rv2_changes_{i}")
    if not isinstance(df, pd.DataFrame):
        df = default_changes_df()
        st.session_state[f"rv2_changes_{i}"] = df
    count_key = f"rv3_change_count_{i}"
    if count_key not in st.session_state:
        st.session_state[count_key] = max(2, len(df))
    count = int(st.session_state[count_key])
    for r in range(count):
        row = df.iloc[r].to_dict() if r < len(df) else {}
        defaults = {
            "coverage": normalize_text(row.get("변경할 보장", "")),
            "before": normalize_text(row.get("기존에는", "")),
            "after": normalize_text(row.get("변경 후에는", "")),
            "change": normalize_text(row.get("어떻게 달라지나요? [목록 선택]", "")) or "선택하세요 ▼",
            "reason": normalize_text(row.get("왜 바꾸나요?", "")),
        }
        for field, value in defaults.items():
            st.session_state.setdefault(f"rv3_change_{i}_{r}_{field}", value)
    return count


def render_plain_change_inputs(i: int) -> None:
    """Editable core-change inputs using ordinary widgets, not st.data_editor."""
    count = _ensure_change_row_state(i)
    st.caption("보장·특약명, 기존 내용, 변경 후 내용, 변화 유형과 변경 이유만 입력합니다. 입력 내용은 즉시 저장됩니다.")

    for r in range(count):
        coverage_now = normalize_text(st.session_state.get(f"rv3_change_{i}_{r}_coverage", ""))
        label = coverage_now or f"핵심 변경 {r + 1}"
        with st.expander(label, expanded=(r < 2)):
            a, b = st.columns(2)
            with a:
                st.text_input("변경할 보장·특약명", key=f"rv3_change_{i}_{r}_coverage", placeholder="예: 암·뇌·심장 진단비")
                st.text_input("기존에는", key=f"rv3_change_{i}_{r}_before", placeholder="예: 2,000만원 / 없음")
                st.selectbox("어떻게 달라지나요? [목록 선택] ▼", CHANGE_OPTIONS, key=f"rv3_change_{i}_{r}_change")
            with b:
                st.text_input("변경 후에는", key=f"rv3_change_{i}_{r}_after", placeholder="예: 5,000만원 / 신규 구성")
            st.text_area("왜 바꾸나요?", key=f"rv3_change_{i}_{r}_reason", placeholder="예: 핵심 진단비를 보완하고 중복 보장은 조정", height=80)

    rows = []
    for r in range(count):
        rows.append({
            "변경할 보장": st.session_state.get(f"rv3_change_{i}_{r}_coverage", ""),
            "기존에는": st.session_state.get(f"rv3_change_{i}_{r}_before", ""),
            "변경 후에는": st.session_state.get(f"rv3_change_{i}_{r}_after", ""),
            "어떻게 달라지나요? [목록 선택]": st.session_state.get(f"rv3_change_{i}_{r}_change", "선택하세요 ▼"),
            "왜 바꾸나요?": st.session_state.get(f"rv3_change_{i}_{r}_reason", ""),
        })
    st.session_state[f"rv2_changes_{i}"] = pd.DataFrame(rows)

    add_col, remove_col = st.columns(2)
    with add_col:
        if st.button("＋ 핵심 변경 항목 추가", key=f"rv3_add_change_{i}", use_container_width=True):
            st.session_state[f"rv3_change_count_{i}"] = min(10, count + 1)
            st.rerun()
    with remove_col:
        if st.button("－ 마지막 항목 삭제", key=f"rv3_remove_change_{i}", use_container_width=True, disabled=count <= 1):
            last = count - 1
            for field in ("coverage", "before", "after", "change", "reason", "premium", "display"):
                st.session_state.pop(f"rv3_change_{i}_{last}_{field}", None)
            st.session_state[f"rv3_change_count_{i}"] = max(1, count - 1)
            st.rerun()


def _ensure_new_plan_row_state(i: int) -> int:
    """신규 가입 보험 요약 입력값을 일반 위젯용 session_state에 준비한다."""
    data_key = f"rv2_additions_{i}"
    df = st.session_state.get(data_key)
    if not isinstance(df, pd.DataFrame):
        df = default_additions_df()

    count_key = f"rv6_plan_count_{i}"
    st.session_state.setdefault(count_key, max(4, len(df)))
    count = int(st.session_state[count_key])

    for r in range(count):
        row = df.iloc[r].to_dict() if r < len(df) else {}
        st.session_state.setdefault(
            f"rv6_plan_{i}_{r}_content",
            normalize_text(row.get("대표 보장 내용", "")),
        )
        st.session_state.setdefault(
            f"rv6_plan_{i}_{r}_premium",
            normalize_text(row.get("월납 보험료", "")),
        )
    return count


def render_new_plan_inputs(i: int) -> None:
    """대표 보장 내용과 월납 보험료만 일반 입력칸으로 받는다."""
    count = _ensure_new_plan_row_state(i)
    st.caption("1페이지에 표시할 대표 보장 내용과 월납 보험료만 입력합니다. 입력 내용은 즉시 유지됩니다.")

    rows = []
    for r in range(count):
        c_no, c_content, c_premium = st.columns([0.45, 4.2, 1.55], vertical_alignment="bottom")
        with c_no:
            st.markdown(f"**{r + 1}**")
        with c_content:
            content = st.text_input(
                "대표 보장 내용",
                key=f"rv6_plan_{i}_{r}_content",
                placeholder="예: 암·뇌·심장 진단비",
                label_visibility="collapsed",
            )
        with c_premium:
            premium = st.text_input(
                "월납 보험료",
                key=f"rv6_plan_{i}_{r}_premium",
                placeholder="예: 128,589",
                label_visibility="collapsed",
            )
        rows.append({"대표 보장 내용": content, "월납 보험료": premium})

    latest = pd.DataFrame(rows, columns=["대표 보장 내용", "월납 보험료"])
    st.session_state[f"rv2_additions_{i}"] = latest.copy()

    total = sum(parse_money(row["월납 보험료"]) for row in rows if normalize_text(row["대표 보장 내용"]))
    st.markdown(f"**입력된 월납 보험료 합계: {total:,}원**")

    add_col, remove_col = st.columns(2)
    with add_col:
        if st.button("＋ 가입 보험 항목 추가", key=f"rv6_add_plan_{i}", use_container_width=True):
            st.session_state[f"rv6_plan_count_{i}"] = min(10, count + 1)
            st.rerun()
    with remove_col:
        if st.button(
            "－ 마지막 항목 삭제",
            key=f"rv6_remove_plan_{i}",
            use_container_width=True,
            disabled=count <= 1,
        ):
            last = count - 1
            st.session_state.pop(f"rv6_plan_{i}_{last}_content", None)
            st.session_state.pop(f"rv6_plan_{i}_{last}_premium", None)
            st.session_state[f"rv6_plan_count_{i}"] = max(1, count - 1)
            st.rerun()

def _ensure_contract_row_state(i: int) -> int:
    df = st.session_state.get(f"rv2_contracts_{i}")
    if not isinstance(df, pd.DataFrame):
        df = default_contract_df()
    key=f"rv4_contract_count_{i}"
    st.session_state.setdefault(key,max(2,len(df)))
    count=int(st.session_state[key])
    for r in range(count):
        row=df.iloc[r].to_dict() if r < len(df) else {}
        defaults={"company":normalize_text(row.get("보험회사","")),"product":normalize_text(row.get("상품명","")),
                  "direction":normalize_text(row.get("처리 방향 [목록 선택]","")) or "선택하세요 ▼",
                  "detail":normalize_text(row.get("구체적인 변경 내용",""))}
        for field,value in defaults.items(): st.session_state.setdefault(f"rv4_contract_{i}_{r}_{field}",value)
    return count


def render_contract_inputs(i: int) -> None:
    count=_ensure_contract_row_state(i)
    st.caption("기존 보험별 보험회사·상품명·처리 방향·구체적인 변경 내용을 입력합니다.")
    for r in range(count):
        product=normalize_text(st.session_state.get(f"rv4_contract_{i}_{r}_product","")) or f"기존 보험 {r+1}"
        with st.expander(product, expanded=(r<2)):
            c1,c2=st.columns(2)
            with c1:
                st.text_input("보험회사",key=f"rv4_contract_{i}_{r}_company",placeholder="예: KB손해보험")
                st.text_input("상품명",key=f"rv4_contract_{i}_{r}_product",placeholder="예: 종합건강보험")
            with c2:
                st.selectbox("처리 방향 [목록 선택] ▼",CONTRACT_OPTIONS,key=f"rv4_contract_{i}_{r}_direction")
            st.text_area("구체적인 변경 내용",key=f"rv4_contract_{i}_{r}_detail",placeholder="예: 실손은 유지하고 중복된 입원일당 특약은 감액 검토",height=80)
    rows=[]
    for r in range(count):
        rows.append({"보험회사":st.session_state.get(f"rv4_contract_{i}_{r}_company",""),
                     "상품명":st.session_state.get(f"rv4_contract_{i}_{r}_product",""),
                     "처리 방향 [목록 선택]":st.session_state.get(f"rv4_contract_{i}_{r}_direction","선택하세요 ▼"),
                     "구체적인 변경 내용":st.session_state.get(f"rv4_contract_{i}_{r}_detail","")})
    st.session_state[f"rv2_contracts_{i}"] = pd.DataFrame(rows)
    a,b=st.columns(2)
    with a:
        if st.button("＋ 기존 보험 추가",key=f"rv4_add_contract_{i}",use_container_width=True):
            st.session_state[f"rv4_contract_count_{i}"] = min(10,count+1); st.rerun()
    with b:
        if st.button("－ 마지막 기존 보험 삭제",key=f"rv4_remove_contract_{i}",use_container_width=True,disabled=count<=1):
            last=count-1
            for field in ("company","product","direction","detail"):
                st.session_state.pop(f"rv4_contract_{i}_{last}_{field}",None)
            st.session_state[f"rv4_contract_count_{i}"] = max(1,count-1); st.rerun()


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
    page_header("고객 상담", APP_TITLE, "고객정보부터 변경안 비교, 고객 시점 미리보기와 엑셀까지 한 번에 완성합니다.", "RM")
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

    tabs = st.tabs(["1. 공통·고객정보", "2. 보험료 비교", "3. 가입 보험·변경 내용", "4. 미리보기·엑셀"])

    with tabs[0]:
        st.subheader("공통 정보")
        common1, common2 = st.columns(2)
        with common1:
            consultation_date = st.date_input("상담일", value=date.today(), key="rv2_date")
            consultant = st.text_input("담당자", key="rv2_consultant", placeholder="예: 박병선 팀장")
        with common2:
            st.info("엑셀은 시트 1의 고객용 비교안과 시트 2의 보장·기존 계약 변경 상세로 구성됩니다.")
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
        st.info("시트 1에는 신규 가입 보험 요약이 표시되고, 시트 2에는 핵심 변경 내용과 기존 보험 변경 방향이 정리됩니다.")
        for i in range(1, person_count + 1):
            name = normalize_text(st.session_state.get(f"rv2_name_{i}", "")) or f"고객 {i}"
            st.subheader(f"{name} 입력")
            with st.expander("1페이지 · 새롭게 가입하는 보험", expanded=True):
                render_new_plan_inputs(i)
            with st.expander("시트 2 상단 · 핵심 변경 내용", expanded=True):
                render_plain_change_inputs(i)
            with st.expander("시트 2 하단 · 기존 보험의 변경 방향", expanded=True):
                render_contract_inputs(i)
            if i < person_count:
                st.divider()

    customers = [collect_customer(i) for i in range(1, person_count + 1)]
    names = [c.name for c in customers]
    title = normalize_text(st.session_state.get("rv2_title", "")) or auto_title(names)
    consultation_date = st.session_state.get("rv2_date", date.today())
    consultant = normalize_text(st.session_state.get("rv2_consultant", ""))

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
            excel = create_excel(customers, title, consultation_date, consultant)
            filename = safe_filename(f"{title}_{consultation_date:%Y%m%d}.xlsx")
            st.download_button("엑셀 다운로드", data=excel, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)


if __name__ == "__main__":
    run()
