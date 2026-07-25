from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from io import BytesIO
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

APP_TITLE = "보험 리모델링 제안서"

NAVY = "17365D"
NAVY_DARK = "102A43"
BLUE = "2F75B5"
GREEN = "2E7D5B"
ORANGE = "B96B26"
TEAL = "168A8A"
GOLD = "C59A3D"
WHITE = "FFFFFF"
BLACK = "1F2937"
GRAY = "E9EDF2"
GRAY_DARK = "667085"
LIGHT_BLUE = "EAF2F8"
LIGHT_GREEN = "E8F5EE"
LIGHT_ORANGE = "FFF2E6"
LIGHT_GOLD = "FBF4E6"
THIN_GRAY = Side(style="thin", color="D7DDE5")
MEDIUM_NAVY = Side(style="medium", color=NAVY)

CHANGE_OPTIONS = [
    "선택하세요 ▼",
    "새로 추가",
    "보장금액 증가",
    "보장금액 감소",
    "보장 범위 확대",
    "보장 범위 축소",
    "보장기간 연장",
    "보장기간 단축",
    "지급 횟수 증가",
    "그대로 유지",
    "정리 또는 삭제",
    "직접 입력",
]
DISPLAY_OPTIONS = ["선택하세요 ▼", "핵심으로 표시", "상세 페이지에만 표시", "출력하지 않음"]
CONTRACT_OPTIONS = ["선택하세요 ▼", "유지", "감액 검토", "일부 특약 조정", "해지 검토", "신규 승인 후 결정", "추가 확인 필요"]


@dataclass(frozen=True)
class AnalysisResult:
    monthly_delta: int
    annual_delta: int
    total_delta: int | None
    monthly_rate: float | None
    price_direction: str
    coverage_direction: str
    result_type: str
    warnings: tuple[str, ...]


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


def format_money_input(value: Any) -> str:
    amount = parse_money(value)
    return f"{amount:,}" if amount else ""


def format_won(value: int | None) -> str:
    return "-" if value is None else f"{int(value):,}원"


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


def format_change(value: int | None) -> str:
    if value is None:
        return "비교하지 않음"
    if value < 0:
        return f"{abs(value):,}원 절감"
    if value > 0:
        return f"{value:,}원 증가"
    return "변동 없음"


def korean_money_hint(value: int) -> str:
    return f"약 {format_compact_won(value)}" if value else "금액을 입력해 주세요."


def default_changes_df() -> pd.DataFrame:
    return pd.DataFrame([
        {
            "변경할 보장": "",
            "기존에는": "",
            "변경 후에는": "",
            "어떻게 달라지나요? [목록 선택]": "선택하세요 ▼",
            "왜 바꾸나요?": "",
            "변경 후 월 보험료": "",
            "첫 장 표시 [목록 선택]": "핵심으로 표시",
        },
        {
            "변경할 보장": "",
            "기존에는": "",
            "변경 후에는": "",
            "어떻게 달라지나요? [목록 선택]": "선택하세요 ▼",
            "왜 바꾸나요?": "",
            "변경 후 월 보험료": "",
            "첫 장 표시 [목록 선택]": "핵심으로 표시",
        },
    ])


def default_contract_df() -> pd.DataFrame:
    return pd.DataFrame([
        {"기존 계약/보장": "", "처리 방향 [목록 선택]": "선택하세요 ▼", "판단 근거": "", "진행 조건": ""},
    ])


def clean_changes(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return default_changes_df().iloc[0:0]
    result = df.copy()
    key = "변경할 보장"
    result[key] = result[key].map(normalize_text)
    result = result[result[key] != ""].reset_index(drop=True)
    for col in ["어떻게 달라지나요? [목록 선택]", "첫 장 표시 [목록 선택]"]:
        if col in result:
            result[col] = result[col].replace("선택하세요 ▼", "")
    return result


def clean_contracts(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return default_contract_df().iloc[0:0]
    result = df.copy()
    result["기존 계약/보장"] = result["기존 계약/보장"].map(normalize_text)
    result = result[result["기존 계약/보장"] != ""].reset_index(drop=True)
    if "처리 방향 [목록 선택]" in result:
        result["처리 방향 [목록 선택]"] = result["처리 방향 [목록 선택]"].replace("선택하세요 ▼", "")
    return result


def detect_coverage_direction(changes: pd.DataFrame) -> str:
    if changes.empty:
        return "정보 부족"
    types = set(changes["어떻게 달라지나요? [목록 선택]"].dropna().astype(str))
    positive = {"새로 추가", "보장금액 증가", "보장 범위 확대", "보장기간 연장", "지급 횟수 증가"}
    negative = {"보장금액 감소", "보장 범위 축소", "보장기간 단축", "정리 또는 삭제"}
    if types & positive and types & negative:
        return "강화와 조정 혼합"
    if types & positive:
        return "강화"
    if types & negative:
        return "축소/조정"
    return "유지/재배분"


def analyze(old_monthly: int, new_monthly: int, old_total: int | None, new_total: int | None, changes: pd.DataFrame) -> AnalysisResult:
    monthly_delta = new_monthly - old_monthly
    annual_delta = monthly_delta * 12
    total_delta = None if old_total is None or new_total is None else new_total - old_total
    monthly_rate = monthly_delta / old_monthly * 100 if old_monthly else None
    price_direction = "감소" if monthly_delta < 0 else "증가" if monthly_delta > 0 else "동일"
    coverage_direction = detect_coverage_direction(changes)

    if price_direction == "감소" and coverage_direction == "강화":
        result_type = "효율 개선형"
    elif price_direction == "감소":
        result_type = "보험료 절감형" if coverage_direction in {"정보 부족", "유지/재배분"} else "부담 조정형"
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
        warnings.append("핵심 변경 내용이 없습니다. 보장 관련 자동 문구는 제한됩니다.")
    if price_direction == "증가" and coverage_direction not in {"강화", "강화와 조정 혼합"}:
        warnings.append("보험료는 증가하지만 강화 근거가 충분하지 않습니다.")
    if old_total is None or new_total is None:
        warnings.append("총 납입액 비교를 사용하지 않아 관련 카드는 표시되지 않습니다.")
    return AnalysisResult(monthly_delta, annual_delta, total_delta, monthly_rate, price_direction, coverage_direction, result_type, tuple(warnings))


def validate_changes(changes: pd.DataFrame) -> list[str]:
    warnings: list[str] = []
    for idx, row in changes.iterrows():
        old = normalize_text(row.get("기존에는"))
        new = normalize_text(row.get("변경 후에는"))
        change = normalize_text(row.get("어떻게 달라지나요? [목록 선택]"))
        name = normalize_text(row.get("변경할 보장")) or f"{idx + 1}번째 항목"
        if not change:
            warnings.append(f"{name}: 변화 유형을 선택해 주세요.")
        if old and new and old == new and change in {"보장금액 증가", "보장금액 감소", "보장 범위 확대", "보장 범위 축소"}:
            warnings.append(f"{name}: 기존과 변경 후 내용은 같은데 변화 유형은 '{change}'입니다.")
        if old in {"없음", "미가입", "-"} and change == "그대로 유지":
            warnings.append(f"{name}: 기존 내용이 없는데 '그대로 유지'로 선택되어 있습니다.")
        if new in {"없음", "삭제", "-"} and change == "새로 추가":
            warnings.append(f"{name}: 변경 후 내용과 '새로 추가'가 서로 맞지 않습니다.")
    core_count = int((changes.get("첫 장 표시 [목록 선택]", pd.Series(dtype=str)) == "핵심으로 표시").sum())
    if core_count > 4:
        warnings.append(f"첫 장 핵심 항목이 {core_count}개입니다. 고객 이해를 위해 4개 이하를 권장합니다.")
    return warnings


def change_sentence(row: pd.Series) -> str:
    name = normalize_text(row.get("변경할 보장"))
    old = normalize_text(row.get("기존에는"))
    new = normalize_text(row.get("변경 후에는"))
    change = normalize_text(row.get("어떻게 달라지나요? [목록 선택]"))
    if old and new:
        return f"{name}: {old} → {new} ({change or '변경'})"
    if new:
        return f"{name}: {new} ({change or '변경'})"
    return f"{name}: {change or '변경 내용 확인'}"


def highlighted_changes(changes: pd.DataFrame, max_items: int = 4) -> list[str]:
    if changes.empty:
        return []
    rows = changes[changes["첫 장 표시 [목록 선택]"] == "핵심으로 표시"]
    if rows.empty:
        rows = changes
    return [change_sentence(row) for _, row in rows.head(max_items).iterrows()]


def top_names(changes: pd.DataFrame, max_items: int = 3) -> list[str]:
    if changes.empty:
        return []
    rows = changes[changes["첫 장 표시 [목록 선택]"] == "핵심으로 표시"]
    if rows.empty:
        rows = changes
    return [normalize_text(v) for v in rows["변경할 보장"].head(max_items) if normalize_text(v)]


def sentence_candidates(analysis: AnalysisResult, changes: pd.DataFrame, goal: str, priorities: list[str], keep_existing: bool, old_monthly: int, new_monthly: int) -> dict[str, list[str]]:
    names = top_names(changes)
    coverage_phrase = "·".join(names) if names else "핵심 보장"
    delta = abs(analysis.monthly_delta)
    headline: list[str] = []
    recommendation: list[str] = []

    if analysis.price_direction == "감소" and analysis.coverage_direction in {"강화", "강화와 조정 혼합"}:
        headline = [
            f"월 보험료를 {delta:,}원 줄이면서 {coverage_phrase} 중심으로 보장을 다시 구성한 제안입니다.",
            f"월 부담은 낮추고 {coverage_phrase}의 부족한 부분은 보완한 효율 개선안입니다.",
            f"기존 보험료를 낮추면서 필요한 보장에 예산을 집중했습니다.",
        ]
    elif analysis.price_direction == "감소":
        headline = [
            f"월 보험료를 {old_monthly:,}원에서 {new_monthly:,}원으로 조정한 제안입니다.",
            f"월 {delta:,}원의 부담을 줄여 장기적으로 유지하기 쉬운 구조로 조정했습니다.",
        ]
    elif analysis.price_direction == "증가" and analysis.coverage_direction in {"강화", "강화와 조정 혼합"}:
        headline = [
            f"월 {delta:,}원의 추가 부담으로 {coverage_phrase}를 강화하는 제안입니다.",
            f"보험료 절감보다 {coverage_phrase}의 보장 공백을 줄이는 데 초점을 두었습니다.",
        ]
    elif analysis.price_direction == "동일" and analysis.coverage_direction == "강화":
        headline = [
            f"현재와 비슷한 월 보험료 안에서 {coverage_phrase}를 강화한 제안입니다.",
            "추가 부담 없이 보험료가 필요한 보장에 쓰이도록 재배분했습니다.",
        ]
    else:
        headline = ["기존 보험과 변경안을 보험료와 핵심 보장 변화 기준으로 비교한 제안입니다."]

    if priorities:
        recommendation.append(f"고객님께서 중요하게 말씀하신 {'·'.join(priorities[:3])}를 우선 반영했습니다.")
    goal_map = {
        "보험료 부담 완화": "장기적으로 유지 가능한 보험료 수준을 우선해 구성했습니다.",
        "핵심 보장 보완": f"부족했던 {coverage_phrase}를 우선 보완하는 방향입니다.",
        "동일 예산 재구성": "현재 예산을 크게 바꾸지 않으면서 필요한 보장으로 보험료를 재배분했습니다.",
        "보장 강화": f"보험료보다 {coverage_phrase}의 보장 수준을 우선한 구성입니다.",
        "특정 위험 집중": f"전체 보장을 넓게 늘리기보다 {coverage_phrase}에 예산을 집중했습니다.",
        "기존 계약 정리": "중복되거나 우선순위가 낮은 부분을 정리하고 핵심 보장에 집중했습니다.",
        "맞춤 재구성": "보험료와 보장 수준을 함께 고려해 현재 상황에 맞게 균형을 조정했습니다.",
    }
    recommendation.append(goal_map.get(goal, goal_map["맞춤 재구성"]))
    if keep_existing:
        recommendation.append("기존 계약 중 유지 가치가 있는 보장은 남기고 조정이 필요한 부분만 구분했습니다.")

    closing = [
        "새로운 계약의 승인 조건을 확인한 뒤 기존 계약의 유지·감액·해지 여부를 결정합니다.",
        "보험료뿐 아니라 보장 범위, 보험기간, 납입기간과 면책·감액 조건을 함께 확인합니다.",
    ]
    return {
        "제목 아래 한 줄": list(dict.fromkeys(headline)),
        "추천 이유": list(dict.fromkeys(recommendation)),
        "마무리·주의": closing,
    }


def _money_text_input(label: str, key: str, help_text: str | None = None, disabled: bool = False) -> int:
    raw_key = f"{key}_raw"
    if raw_key not in st.session_state:
        st.session_state[raw_key] = ""
    raw = st.text_input(label, key=raw_key, placeholder="예: 694,580", help=help_text, disabled=disabled)
    amount = parse_money(raw)
    if raw and not disabled:
        st.caption(f"입력 금액: **{amount:,}원** · {korean_money_hint(amount)}")
    return amount


def _metric_card(title: str, main: str, sub: str, direction: str) -> str:
    if direction == "감소":
        color, bg = "#2E7D5B", "#E8F5EE"
    elif direction == "증가":
        color, bg = "#B96B26", "#FFF2E6"
    else:
        color, bg = "#17365D", "#EAF2F8"
    return f"""
    <div style="border:1px solid #D7DDE5;border-radius:12px;padding:18px 16px;background:{bg};min-height:130px;box-shadow:0 2px 8px rgba(16,42,67,.06);">
      <div style="font-size:14px;color:#667085;font-weight:700;margin-bottom:8px;">{title}</div>
      <div style="font-size:26px;color:{color};font-weight:800;line-height:1.2;margin-bottom:10px;">{main}</div>
      <div style="font-size:13px;color:#344054;">{sub}</div>
    </div>
    """


def render_preview(title: str, proposal_label: str, headline: str, analysis: AnalysisResult, old_monthly: int, new_monthly: int, old_total: int | None, new_total: int | None, changes: pd.DataFrame, recommendation: str, caution: str) -> None:
    st.markdown(f"## {title}")
    st.caption(f"{proposal_label} · {analysis.result_type}")
    st.info(headline or "입력 내용을 기준으로 핵심 설명 문구가 표시됩니다.")

    direction = analysis.price_direction
    total_direction = "동일" if analysis.total_delta == 0 else "감소" if (analysis.total_delta or 0) < 0 else "증가"
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(_metric_card("월 보험료 변화", format_change(analysis.monthly_delta), f"{old_monthly:,}원 → {new_monthly:,}원", direction), unsafe_allow_html=True)
    with c2:
        st.markdown(_metric_card("연간 보험료 변화", format_change(analysis.annual_delta), "월 차액 × 12개월", direction), unsafe_allow_html=True)
    with c3:
        if analysis.total_delta is None:
            main, sub, total_direction = "비교하지 않음", "총 납입액을 입력하면 표시됩니다.", "동일"
        else:
            main = format_change(analysis.total_delta)
            sub = f"{format_compact_won(old_total)} → {format_compact_won(new_total)}"
        st.markdown(_metric_card("총 납입액 변화", main, sub, total_direction), unsafe_allow_html=True)

    st.markdown("### 핵심 변경 전후")
    rows = []
    for _, row in changes[changes["첫 장 표시 [목록 선택]"] != "출력하지 않음"].head(4).iterrows():
        rows.append({
            "핵심 항목": normalize_text(row.get("변경할 보장")),
            "기존": normalize_text(row.get("기존에는")) or "-",
            "변경 후": normalize_text(row.get("변경 후에는")) or "-",
            "결과": normalize_text(row.get("어떻게 달라지나요? [목록 선택]")) or "변경",
        })
    if rows:
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    else:
        st.caption("핵심 변경 내용을 입력하면 최대 4개가 표시됩니다.")

    st.markdown("### 이번 변경이 필요한 이유")
    reasons = [normalize_text(v) for v in changes["왜 바꾸나요?"].tolist() if normalize_text(v)]
    if reasons:
        for reason in list(dict.fromkeys(reasons))[:4]:
            st.markdown(f"- {reason}")
    else:
        st.write(recommendation or "추천 이유를 선택하거나 수정해 주세요.")

    st.markdown("### 진행 안내")
    st.caption(caution)


def _merge_write(ws, cell_range: str, value: Any, *, font=None, fill=None, alignment=None, border=None):
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(":")[0]]
    cell.value = value
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment
    if border:
        for row in ws[cell_range]:
            for c in row:
                c.border = border
    return cell


def _section_title(ws, row: int, title: str, end_col: int):
    _merge_write(ws, f"A{row}:{get_column_letter(end_col)}{row}", title,
                 font=Font(name="맑은 고딕", size=12, bold=True, color=WHITE),
                 fill=PatternFill("solid", fgColor=NAVY),
                 alignment=Alignment(horizontal="left", vertical="center"),
                 border=Border(bottom=MEDIUM_NAVY))
    ws.row_dimensions[row].height = 25


def _page_setup(ws, fit_height: int = 1):
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_height
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.35, bottom=0.35, header=0.15, footer=0.15)
    ws.print_options.horizontalCentered = True


def _change_color(change_text: str) -> tuple[str, str]:
    text = normalize_text(change_text)
    if any(word in text for word in ["증가", "확대", "추가", "연장"]):
        return GREEN, LIGHT_GREEN
    if any(word in text for word in ["감소", "축소", "단축", "삭제", "정리"]):
        return ORANGE, LIGHT_ORANGE
    return BLUE, LIGHT_BLUE


def create_excel(*, customer_name: str, title: str, proposal_label: str, consultation_date: date, consultant: str,
                 analysis: AnalysisResult, old_monthly: int, new_monthly: int, old_total: int | None, new_total: int | None,
                 changes: pd.DataFrame, contracts: pd.DataFrame, headline: str, recommendation: str, caution: str,
                 output_type: str) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "01_한눈에보기"
    for col, width in {"A": 17, "B": 17, "C": 18, "D": 18, "E": 3, "F": 17, "G": 17, "H": 18, "I": 18, "J": 3}.items():
        ws.column_dimensions[col].width = width

    _merge_write(ws, "A1:J2", title, font=Font(name="맑은 고딕", size=22, bold=True, color=NAVY_DARK), alignment=Alignment(horizontal="left", vertical="center"))
    _merge_write(ws, "A3:H3", headline, font=Font(name="맑은 고딕", size=11, bold=True, color=BLACK), fill=PatternFill("solid", fgColor=LIGHT_BLUE), alignment=Alignment(horizontal="left", vertical="center", wrap_text=True), border=Border(left=MEDIUM_NAVY))
    _merge_write(ws, "I3:J3", proposal_label, font=Font(name="맑은 고딕", size=10, bold=True, color=WHITE), fill=PatternFill("solid", fgColor=TEAL), alignment=Alignment(horizontal="center", vertical="center"))
    ws.row_dimensions[3].height = 38

    cards = [
        ("월 보험료 변화", format_change(analysis.monthly_delta), f"{old_monthly:,}원 → {new_monthly:,}원", analysis.price_direction),
        ("연간 보험료 변화", format_change(analysis.annual_delta), "월 차액 × 12개월", analysis.price_direction),
        ("총 납입액 변화", format_change(analysis.total_delta), f"{format_compact_won(old_total)} → {format_compact_won(new_total)}", "감소" if (analysis.total_delta or 0) < 0 else "증가" if (analysis.total_delta or 0) > 0 else "동일"),
    ]
    ranges = ["A5:C7", "D5:F7", "G5:J7"]
    for (label, main, sub, direction), rng in zip(cards, ranges):
        font_color = GREEN if direction == "감소" else ORANGE if direction == "증가" else NAVY
        fill = LIGHT_GREEN if direction == "감소" else LIGHT_ORANGE if direction == "증가" else LIGHT_BLUE
        _merge_write(ws, rng, f"{label}\n{main}\n{sub}", font=Font(name="맑은 고딕", size=14, bold=True, color=font_color), fill=PatternFill("solid", fgColor=fill), alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY))
    for row in range(5, 8):
        ws.row_dimensions[row].height = 25

    _section_title(ws, 9, "핵심 변경 전후", 10)
    headers = [("A10:B10", "핵심 항목"), ("C10:D10", "기존"), ("E10:G10", "변경 후"), ("H10:J10", "결과")]
    for rng, text in headers:
        _merge_write(ws, rng, text, font=Font(name="맑은 고딕", size=10, bold=True, color=WHITE), fill=PatternFill("solid", fgColor=NAVY_DARK), alignment=Alignment(horizontal="center", vertical="center"), border=Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY))

    core = changes[changes["첫 장 표시 [목록 선택]"] == "핵심으로 표시"]
    if core.empty:
        core = changes[changes["첫 장 표시 [목록 선택]"] != "출력하지 않음"]
    r = 11
    for _, item in core.head(4).iterrows():
        color, fill = _change_color(item.get("어떻게 달라지나요? [목록 선택]"))
        cells = [
            ("A", "B", normalize_text(item.get("변경할 보장")), GRAY, BLACK),
            ("C", "D", normalize_text(item.get("기존에는")) or "-", WHITE, BLACK),
            ("E", "G", normalize_text(item.get("변경 후에는")) or "-", LIGHT_BLUE, BLUE),
            ("H", "J", normalize_text(item.get("어떻게 달라지나요? [목록 선택]")) or "변경", fill, color),
        ]
        for start, end, value, fill_color, font_color in cells:
            _merge_write(ws, f"{start}{r}:{end}{r}", value, font=Font(name="맑은 고딕", size=10, bold=start in {"A", "E", "H"}, color=font_color), fill=PatternFill("solid", fgColor=fill_color), alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY))
        ws.row_dimensions[r].height = 32
        r += 1
    while r <= 14:
        _merge_write(ws, f"A{r}:J{r}", "", fill=PatternFill("solid", fgColor=WHITE), border=Border(bottom=THIN_GRAY))
        ws.row_dimensions[r].height = 25
        r += 1

    _section_title(ws, 16, "이번 변경이 필요한 이유", 10)
    reasons = [normalize_text(v) for v in changes["왜 바꾸나요?"].tolist() if normalize_text(v)]
    if not reasons:
        reasons = [recommendation]
    r = 17
    for idx, reason in enumerate(list(dict.fromkeys(reasons))[:4], 1):
        _merge_write(ws, f"A{r}:J{r}", f"{idx}. {reason}", font=Font(name="맑은 고딕", size=10.5, bold=idx <= 2, color=BLACK), fill=PatternFill("solid", fgColor=WHITE if idx % 2 else "F7F9FC"), alignment=Alignment(horizontal="left", vertical="center", wrap_text=True), border=Border(bottom=THIN_GRAY))
        ws.row_dimensions[r].height = 29
        r += 1
    while r <= 20:
        _merge_write(ws, f"A{r}:J{r}", "", fill=PatternFill("solid", fgColor=WHITE), border=Border(bottom=THIN_GRAY))
        ws.row_dimensions[r].height = 24
        r += 1

    _section_title(ws, 22, "진행 안내", 10)
    _merge_write(ws, "A23:J24", caution, font=Font(name="맑은 고딕", size=9, color=GRAY_DARK), fill=PatternFill("solid", fgColor="F7F8FA"), alignment=Alignment(horizontal="left", vertical="center", wrap_text=True), border=Border(left=MEDIUM_NAVY))
    _merge_write(ws, "A25:F25", f"상담일: {consultation_date:%Y-%m-%d}", font=Font(name="맑은 고딕", size=8.5, color=GRAY_DARK), alignment=Alignment(horizontal="left"))
    _merge_write(ws, "G25:J25", f"담당: {consultant or '-'}", font=Font(name="맑은 고딕", size=8.5, color=GRAY_DARK), alignment=Alignment(horizontal="right"))
    ws.print_area = "A1:J25"
    _page_setup(ws, 1)

    if output_type in {"표준형(2장)", "상세형(3장)"}:
        ws2 = wb.create_sheet("02_변경상세")
        widths = [24, 20, 20, 19, 31, 16]
        for i, width in enumerate(widths, 1):
            ws2.column_dimensions[get_column_letter(i)].width = width
        _merge_write(ws2, "A1:F2", f"{customer_name}님 핵심 변경 상세", font=Font(name="맑은 고딕", size=20, bold=True, color=NAVY_DARK), alignment=Alignment(horizontal="left", vertical="center"))
        _merge_write(ws2, "A3:F3", "고객에게 설명할 중요한 변경 내용만 정리했습니다.", font=Font(name="맑은 고딕", size=10, color=BLACK), fill=PatternFill("solid", fgColor=LIGHT_BLUE), alignment=Alignment(horizontal="left", vertical="center"))
        headers2 = ["변경할 보장", "기존에는", "변경 후에는", "변화", "왜 바꾸나요?", "변경 후 월 보험료"]
        for c, header in enumerate(headers2, 1):
            cell = ws2.cell(5, c, header)
            cell.font = Font(name="맑은 고딕", size=10, bold=True, color=WHITE)
            cell.fill = PatternFill("solid", fgColor=NAVY_DARK)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
        r2 = 6
        for _, item in changes[changes["첫 장 표시 [목록 선택]"] != "출력하지 않음"].iterrows():
            values = [
                normalize_text(item.get("변경할 보장")), normalize_text(item.get("기존에는")), normalize_text(item.get("변경 후에는")),
                normalize_text(item.get("어떻게 달라지나요? [목록 선택]")), normalize_text(item.get("왜 바꾸나요?")), parse_money(item.get("변경 후 월 보험료")),
            ]
            change_color, change_fill = _change_color(values[3])
            for c, value in enumerate(values, 1):
                cell = ws2.cell(r2, c, value)
                cell.font = Font(name="맑은 고딕", size=9.5, bold=c in {1, 3, 4}, color=change_color if c == 4 else BLUE if c == 3 else BLACK)
                cell.fill = PatternFill("solid", fgColor=change_fill if c == 4 else LIGHT_BLUE if c == 3 else WHITE)
                cell.alignment = Alignment(horizontal="right" if c == 6 else "left", vertical="center", wrap_text=True)
                cell.border = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
                if c == 6 and value:
                    cell.number_format = '#,##0"원"'
            ws2.row_dimensions[r2].height = 42
            r2 += 1
        ws2.print_area = f"A1:F{max(r2, 12)}"
        _page_setup(ws2, 0)

    if output_type == "상세형(3장)":
        ws3 = wb.create_sheet("03_계약정리")
        widths3 = [26, 22, 36, 36]
        for i, width in enumerate(widths3, 1):
            ws3.column_dimensions[get_column_letter(i)].width = width
        _merge_write(ws3, "A1:D2", f"{customer_name}님 기존 계약 정리 및 확인사항", font=Font(name="맑은 고딕", size=20, bold=True, color=NAVY_DARK), alignment=Alignment(horizontal="left", vertical="center"))
        _merge_write(ws3, "A3:D3", "새로운 계약의 승인 조건을 확인한 뒤 기존 계약의 처리 방향을 결정합니다.", font=Font(name="맑은 고딕", size=10, bold=True, color=BLACK), fill=PatternFill("solid", fgColor=LIGHT_GOLD), alignment=Alignment(horizontal="left", vertical="center"))
        headers3 = ["기존 계약/보장", "처리 방향", "판단 근거", "진행 조건"]
        for c, header in enumerate(headers3, 1):
            cell = ws3.cell(5, c, header)
            cell.font = Font(name="맑은 고딕", size=10, bold=True, color=WHITE)
            cell.fill = PatternFill("solid", fgColor=NAVY_DARK)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
        r3 = 6
        for _, item in contracts.iterrows():
            values = [normalize_text(item.get("기존 계약/보장")), normalize_text(item.get("처리 방향 [목록 선택]")), normalize_text(item.get("판단 근거")), normalize_text(item.get("진행 조건"))]
            for c, value in enumerate(values, 1):
                cell = ws3.cell(r3, c, value)
                cell.font = Font(name="맑은 고딕", size=9.5, bold=c in {1, 2}, color=BLACK)
                cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE if c == 2 else WHITE)
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                cell.border = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
            ws3.row_dimensions[r3].height = 40
            r3 += 1
        _section_title(ws3, max(r3 + 1, 10), "반드시 확인할 내용", 4)
        r3 = max(r3 + 2, 11)
        checks = [
            "새로운 계약의 정상 승인 여부와 조건부 승인 내용을 먼저 확인합니다.",
            "기존 계약의 해지환급금, 납입 경과기간과 유지 가치를 함께 확인합니다.",
            "면책기간, 감액기간, 갱신 여부, 보험기간과 납입기간을 비교합니다.",
            "기존 계약을 먼저 해지해 보장 공백이 생기지 않도록 진행 순서를 확인합니다.",
        ]
        for idx, text in enumerate(checks, 1):
            _merge_write(ws3, f"A{r3}:D{r3}", f"{idx}. {text}", font=Font(name="맑은 고딕", size=10, color=BLACK), fill=PatternFill("solid", fgColor=WHITE if idx % 2 else "F7F9FC"), alignment=Alignment(horizontal="left", vertical="center", wrap_text=True), border=Border(bottom=THIN_GRAY))
            ws3.row_dimensions[r3].height = 30
            r3 += 1
        ws3.print_area = f"A1:D{r3}"
        _page_setup(ws3, 1)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def _safe_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip() or "보험리모델링_제안서"


def _load_example() -> None:
    st.session_state["rp_customer"] = "홍길동"
    st.session_state["rp_old_month_raw"] = "694,580"
    st.session_state["rp_new_month_raw"] = "367,707"
    st.session_state["rp_old_total_raw"] = "144,940,000"
    st.session_state["rp_new_total_raw"] = "88,240,000"
    st.session_state["rp_changes_df"] = pd.DataFrame([
        {"변경할 보장": "암 진단비", "기존에는": "2,000만원", "변경 후에는": "5,000만원", "어떻게 달라지나요? [목록 선택]": "보장금액 증가", "왜 바꾸나요?": "암 진단 시 필요한 치료비와 생활자금을 더 충분히 준비하기 위해", "변경 후 월 보험료": "80,000", "첫 장 표시 [목록 선택]": "핵심으로 표시"},
        {"변경할 보장": "암 주요치료비", "기존에는": "없음", "변경 후에는": "신규 구성", "어떻게 달라지나요? [목록 선택]": "새로 추가", "왜 바꾸나요?": "진단 이후 반복될 수 있는 치료비를 준비하기 위해", "변경 후 월 보험료": "50,000", "첫 장 표시 [목록 선택]": "핵심으로 표시"},
        {"변경할 보장": "기존 실손보험", "기존에는": "가입 중", "변경 후에는": "유지", "어떻게 달라지나요? [목록 선택]": "그대로 유지", "왜 바꾸나요?": "기존 가입 조건의 유지 가치를 고려하기 위해", "변경 후 월 보험료": "", "첫 장 표시 [목록 선택]": "상세 페이지에만 표시"},
    ])
    st.session_state["rp_example_loaded"] = True


def run():
    st.title(f"🔁 {APP_TITLE}")
    st.caption("누구의 보험인지 → 돈이 얼마나 달라지는지 → 무엇을 왜 바꾸는지 → 고객에게 어떻게 보여줄지 순서로 작성합니다.")

    if "rp_changes_df" not in st.session_state:
        st.session_state.rp_changes_df = default_changes_df()
    if "rp_contract_df" not in st.session_state:
        st.session_state.rp_contract_df = default_contract_df()

    top1, top2 = st.columns([1, 4])
    with top1:
        if st.button("예시로 먼저 보기", use_container_width=True):
            _load_example()
            st.rerun()
    with top2:
        if st.session_state.get("rp_example_loaded"):
            st.info("현재 홍길동 예시 데이터가 입력되어 있습니다. 실제 고객 정보로 바꾸어 사용하세요.")

    tabs = st.tabs(["1. 기본정보", "2. 보험료 비교", "3. 핵심 변경 내용", "4. 미리보기·엑셀"])

    with tabs[0]:
        st.subheader("누구에게 어떤 방향으로 제안하나요?")
        c1, c2 = st.columns(2)
        with c1:
            customer_name = st.text_input("고객명", placeholder="예: 홍길동", key="rp_customer")
            goal = st.selectbox("설계 목적 [목록 선택] ▼", ["보험료 부담 완화", "핵심 보장 보완", "동일 예산 재구성", "보장 강화", "특정 위험 집중", "기존 계약 정리", "맞춤 재구성"], key="rp_goal")
            proposal_label = st.text_input("제안 표시", value="균형 보장형 · 추천안", key="rp_label")
        with c2:
            consultation_date = st.date_input("상담일", value=date.today(), key="rp_date")
            consultant = st.text_input("담당자", placeholder="예: 박병선 팀장", key="rp_consultant")
            output_type = st.radio("엑셀 출력", ["간단형(1장)", "표준형(2장)", "상세형(3장)"], horizontal=True, key="rp_output")
        title = st.text_input("자료 제목", placeholder="비워두면 '홍길동님 보험 리모델링 비교안' 형식으로 자동 생성됩니다.", key="rp_title")
        priorities = st.multiselect("고객이 중요하게 생각하는 부분 [목록 선택] ▼", ["월 보험료 부담", "암 보장", "뇌·심장 보장", "치료비", "간병비", "사망보장", "노후 의료비", "납입기간", "갱신 부담", "기존 계약 최대한 유지"], key="rp_priorities")
        keep_existing = st.checkbox("유지 가치가 있는 기존 보장은 남기는 방향", value=True, key="rp_keep")

    with tabs[1]:
        st.subheader("돈이 얼마나 달라지나요?")
        st.caption("숫자만 입력해도 계산되며, 입력 금액은 아래에 쉼표와 한글 단위로 확인할 수 있습니다.")
        c1, c2 = st.columns(2)
        with c1:
            old_monthly = _money_text_input("기존 월 보험료", "rp_old_month")
            use_total = st.checkbox("총 납입액도 비교하기", value=True, key="rp_use_total")
            old_total = _money_text_input("기존 납입예정 총액", "rp_old_total", disabled=not use_total)
        with c2:
            new_monthly = _money_text_input("변경 후 월 보험료", "rp_new_month")
            new_total = _money_text_input("변경 후 납입예정 총액", "rp_new_total", disabled=not use_total)
        if not use_total:
            old_total = new_total = None
        elif not old_total and not new_total:
            old_total = new_total = None
        if old_monthly or new_monthly:
            delta = new_monthly - old_monthly
            st.success(f"월 보험료 {format_change(delta)} · 연간 {format_change(delta * 12)}")

    with tabs[2]:
        st.subheader("무엇을 왜 바꾸나요?")
        st.info("고객에게 설명할 중요한 변경만 입력하세요. 모든 담보를 입력할 필요는 없습니다.")
        st.markdown("**직접 입력:** 흰색 칸 · **목록 선택:** 제목에 `[목록 선택]`과 ▼가 있는 칸")
        st.caption("예: 암 진단비 | 2,000만원 | 5,000만원 | 보장금액 증가 | 진단 후 필요한 자금 준비")

        with st.form("rp_changes_form", clear_on_submit=False):
            edited = st.data_editor(
                st.session_state.rp_changes_df,
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                column_config={
                    "변경할 보장": st.column_config.TextColumn("변경할 보장", required=True, width="medium", help="고객 자료에는 입력한 담보명을 그대로 사용합니다."),
                    "기존에는": st.column_config.TextColumn("기존에는", width="medium", help="예: 2,000만원 / 없음 / 가입 중"),
                    "변경 후에는": st.column_config.TextColumn("변경 후에는", width="medium", help="예: 5,000만원 / 신규 구성 / 유지"),
                    "어떻게 달라지나요? [목록 선택]": st.column_config.SelectboxColumn("어떻게 달라지나요? [목록 선택] ▼", options=CHANGE_OPTIONS, required=True, width="medium"),
                    "왜 바꾸나요?": st.column_config.TextColumn("왜 바꾸나요?", width="large"),
                    "변경 후 월 보험료": st.column_config.TextColumn("변경 후 월 보험료", width="small", help="선택 입력입니다. 숫자만 입력해도 됩니다."),
                    "첫 장 표시 [목록 선택]": st.column_config.SelectboxColumn("첫 장 표시 [목록 선택] ▼", options=DISPLAY_OPTIONS, required=True, width="medium"),
                },
                key="rp_changes_editor",
            )
            submitted = st.form_submit_button("핵심 변경 내용 적용", type="primary", use_container_width=True)
        if submitted:
            st.session_state.rp_changes_df = edited
            st.success("핵심 변경 내용을 적용했습니다.")

        changes = clean_changes(st.session_state.rp_changes_df)
        if not changes.empty:
            core_count = int((changes["첫 장 표시 [목록 선택]"] == "핵심으로 표시").sum())
            st.caption(f"현재 입력: {len(changes)}건 · 첫 장 핵심 표시: {core_count}건")
            for warning in validate_changes(changes):
                st.warning(warning)

        with st.expander("기존 계약 처리 방향 입력 · 상세형(3장)에서 사용"):
            st.caption("▼ 표시가 있는 칸은 목록에서 선택하세요.")
            with st.form("rp_contract_form", clear_on_submit=False):
                contract_edited = st.data_editor(
                    st.session_state.rp_contract_df,
                    num_rows="dynamic",
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "기존 계약/보장": st.column_config.TextColumn(required=True, width="large"),
                        "처리 방향 [목록 선택]": st.column_config.SelectboxColumn("처리 방향 [목록 선택] ▼", options=CONTRACT_OPTIONS, required=True),
                        "판단 근거": st.column_config.TextColumn(width="large"),
                        "진행 조건": st.column_config.TextColumn(width="large"),
                    },
                    key="rp_contract_editor",
                )
                contract_submit = st.form_submit_button("계약 처리 내용 적용", use_container_width=True)
            if contract_submit:
                st.session_state.rp_contract_df = contract_edited
                st.success("계약 처리 내용을 적용했습니다.")

    customer_name = st.session_state.get("rp_customer", "")
    title = st.session_state.get("rp_title", "") or (f"{customer_name}님 보험 리모델링 비교안" if customer_name else "보험 리모델링 비교안")
    goal = st.session_state.get("rp_goal", "맞춤 재구성")
    proposal_label = st.session_state.get("rp_label", "균형 보장형 · 추천안")
    consultation_date = st.session_state.get("rp_date", date.today())
    consultant = st.session_state.get("rp_consultant", "")
    output_type = st.session_state.get("rp_output", "간단형(1장)")
    priorities = st.session_state.get("rp_priorities", [])
    keep_existing = bool(st.session_state.get("rp_keep", True))
    old_monthly = parse_money(st.session_state.get("rp_old_month_raw", ""))
    new_monthly = parse_money(st.session_state.get("rp_new_month_raw", ""))
    use_total = bool(st.session_state.get("rp_use_total", True))
    old_total = parse_money(st.session_state.get("rp_old_total_raw", "")) if use_total else None
    new_total = parse_money(st.session_state.get("rp_new_total_raw", "")) if use_total else None
    if use_total and old_total == 0 and new_total == 0:
        old_total = new_total = None
    changes = clean_changes(st.session_state.rp_changes_df)
    contracts = clean_contracts(st.session_state.rp_contract_df)
    analysis = analyze(old_monthly, new_monthly, old_total, new_total, changes)
    candidates = sentence_candidates(analysis, changes, goal, priorities, keep_existing, old_monthly, new_monthly)

    with tabs[3]:
        st.subheader("고객에게 어떻게 보여줄까요?")
        a1, a2, a3, a4 = st.columns(4)
        a1.metric("실제 결과", analysis.result_type)
        a2.metric("보험료 방향", analysis.price_direction)
        a3.metric("보장 방향", analysis.coverage_direction)
        a4.metric("핵심 변경", f"{len(changes)}건")
        for warning in analysis.warnings:
            st.warning(warning)
        for warning in validate_changes(changes):
            st.warning(warning)

        headline_options = candidates["제목 아래 한 줄"] or ["입력 내용을 기준으로 핵심 설명 문구가 표시됩니다."]
        headline_choice = st.selectbox("첫 장 핵심 결론 [목록 선택] ▼", headline_options, key="rp_headline_choice")
        headline = st.text_area("최종 핵심 결론", value=headline_choice, height=80, key="rp_headline_final")

        recommendation_options = candidates["추천 이유"] or ["보험료와 보장 수준을 함께 고려한 제안입니다."]
        recommendation_choice = st.selectbox("추천 이유 [목록 선택] ▼", recommendation_options, key="rp_recommend_choice")
        recommendation = st.text_area("최종 추천 이유", value=recommendation_choice, height=90, key="rp_recommend_final")

        caution_base = " ".join(candidates["마무리·주의"])
        caution = st.text_area("진행 안내·주의사항", value=caution_base + " 본 자료는 입력된 계약정보를 기준으로 작성되며 실제 보장내용과 가입 조건은 보험회사·상품·담보·심사 결과 및 약관에 따라 달라질 수 있습니다.", height=100, key="rp_caution")

        st.divider()
        st.subheader("고객 시점 미리보기")
        render_preview(title, proposal_label, headline, analysis, old_monthly, new_monthly, old_total, new_total, changes, recommendation, caution)

        st.divider()
        st.subheader("출력 전 최종 점검")
        checks = {
            "고객명": bool(customer_name),
            "보험료 비교": bool(old_monthly and new_monthly),
            "핵심 변경 내용": not changes.empty,
            "첫 장 핵심 항목": bool((changes.get("첫 장 표시 [목록 선택]", pd.Series(dtype=str)) == "핵심으로 표시").any()),
        }
        for label, okay in checks.items():
            st.markdown(f"{'✅' if okay else '⚠️'} {label}")

        if customer_name and old_monthly and new_monthly and not changes.empty:
            excel = create_excel(
                customer_name=customer_name,
                title=title,
                proposal_label=proposal_label,
                consultation_date=consultation_date,
                consultant=consultant,
                analysis=analysis,
                old_monthly=old_monthly,
                new_monthly=new_monthly,
                old_total=old_total,
                new_total=new_total,
                changes=changes,
                contracts=contracts,
                headline=headline,
                recommendation=recommendation,
                caution=caution,
                output_type=output_type,
            )
            filename = _safe_filename(f"{customer_name}님_보험리모델링_{output_type}_{consultation_date:%Y%m%d}.xlsx")
            st.download_button("엑셀 다운로드", data=excel, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
        else:
            st.info("고객명, 기존·변경 월 보험료와 핵심 변경 내용을 입력하면 엑셀을 다운로드할 수 있습니다.")
