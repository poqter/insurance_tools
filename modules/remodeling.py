from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from io import BytesIO
from typing import Any, Iterable

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins


APP_TITLE = "보험 리모델링 전/후 비교"

# 고객용 표기는 사용자가 입력한 담보명에 실제 포함된 단어만 사용합니다.
# 임의의 쉬운 명칭 변환 사전은 사용하지 않습니다.

NAVY = "17365D"
NAVY_DARK = "102A43"
BLUE = "2F75B5"
TEAL = "168A8A"
GOLD = "C59A3D"
GRAY = "E9EDF2"
GRAY_DARK = "667085"
LIGHT_BLUE = "EAF2F8"
LIGHT_TEAL = "E8F5F3"
LIGHT_GOLD = "FBF4E6"
WHITE = "FFFFFF"
BLACK = "1F2937"
RED_SOFT = "FCE8E6"
ORANGE = "B96B26"
GREEN = "2E7D5B"

THIN_GRAY = Side(style="thin", color="D7DDE5")
MEDIUM_NAVY = Side(style="medium", color=NAVY)


@dataclass(frozen=True)
class AnalysisResult:
    monthly_delta: int
    annual_delta: int
    total_delta: int | None
    monthly_rate: float | None
    total_rate: float | None
    price_direction: str
    total_direction: str
    coverage_direction: str
    result_type: str
    warnings: tuple[str, ...]


def _money(value: Any) -> int:
    try:
        if value is None or pd.isna(value):
            return 0
        return int(round(float(value)))
    except (TypeError, ValueError):
        return 0


def format_won(value: int | float | None) -> str:
    if value is None:
        return "-"
    return f"{int(round(value)):,}원"


def format_compact_won(value: int | float | None, approximate: bool = False) -> str:
    if value is None:
        return "-"
    amount = int(round(value))
    prefix = "약 " if approximate else ""
    abs_amount = abs(amount)
    sign = "-" if amount < 0 else ""

    if abs_amount >= 100_000_000:
        eok = abs_amount // 100_000_000
        man = (abs_amount % 100_000_000) // 10_000
        if man:
            return f"{prefix}{sign}{eok}억 {man:,}만원"
        return f"{prefix}{sign}{eok}억원"
    if abs_amount >= 10_000:
        return f"{prefix}{sign}{abs_amount / 10_000:,.0f}만원"
    return f"{prefix}{sign}{abs_amount:,}원"


def format_change(value: int | None, decrease_word: str = "절감", increase_word: str = "증가") -> str:
    if value is None:
        return "비교하지 않음"
    if value < 0:
        return f"{format_won(abs(value))} {decrease_word}"
    if value > 0:
        return f"{format_won(value)} {increase_word}"
    return "변동 없음"


def normalize_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def contains_only_words_from_name(display_text: str, coverage_name: str) -> bool:
    """고객용 표시가 담보명에 포함된 단어만으로 구성됐는지 보수적으로 확인합니다."""
    display = normalize_text(display_text)
    name = normalize_text(coverage_name)
    if not display:
        return True
    tokens = [t for t in re.split(r"[\s·/,+()\-]+", display) if len(t) >= 2]
    return all(token in name for token in tokens)


def default_rebuild_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"보장 항목(담보명 포함 표현)": "", "월 보험료": 0, "보험회사": "", "구성 목적": "", "첫 장 강조": True},
            {"보장 항목(담보명 포함 표현)": "", "월 보험료": 0, "보험회사": "", "구성 목적": "", "첫 장 강조": True},
            {"보장 항목(담보명 포함 표현)": "", "월 보험료": 0, "보험회사": "", "구성 목적": "", "첫 장 강조": False},
        ]
    )


def default_change_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "담보명": "",
                "기존 보장": "",
                "변경 후 보장": "",
                "변화 유형": "가입금액 증액",
                "변화 설명": "",
                "구성 목적": "",
                "첫 장 강조": True,
            },
            {
                "담보명": "",
                "기존 보장": "",
                "변경 후 보장": "",
                "변화 유형": "신규 보장 추가",
                "변화 설명": "",
                "구성 목적": "",
                "첫 장 강조": True,
            },
        ]
    )


def default_contract_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"기존 계약/보장": "", "처리 방향": "유지", "판단 근거": "", "진행 조건": ""},
            {"기존 계약/보장": "", "처리 방향": "신규 승인 후 결정", "판단 근거": "", "진행 조건": ""},
        ]
    )


def clean_df(df: pd.DataFrame, required_col: str) -> pd.DataFrame:
    if df is None or df.empty or required_col not in df.columns:
        return pd.DataFrame(columns=df.columns if df is not None else [])
    result = df.copy()
    result[required_col] = result[required_col].map(normalize_text)
    return result[result[required_col] != ""].reset_index(drop=True)


def price_bucket(rate: float | None, delta: int) -> str:
    if rate is None:
        return "동일"
    abs_rate = abs(rate)
    if delta < 0:
        return "큰 폭 감소" if abs_rate >= 20 else "소폭 감소"
    if delta > 0:
        return "큰 폭 증가" if abs_rate >= 20 else "소폭 증가"
    return "동일"


def detect_coverage_direction(changes: pd.DataFrame) -> str:
    if changes.empty:
        return "정보 부족"
    types = set(changes["변화 유형"].dropna().astype(str))
    positive = {
        "가입금액 증액", "보장 범위 확대", "신규 보장 추가", "보장기간 연장",
        "지급 횟수 확대", "특정 보장 집중 보완", "갱신 부담 조정"
    }
    negative = {"가입금액 축소", "보장 범위 축소", "보장기간 단축", "보장 삭제/정리"}
    has_positive = bool(types & positive)
    has_negative = bool(types & negative)
    if has_positive and has_negative:
        return "강화와 조정 혼합"
    if has_positive:
        return "강화"
    if has_negative:
        return "축소/조정"
    return "유사/재배분"


def analyze(
    old_monthly: int,
    new_monthly: int,
    old_total: int | None,
    new_total: int | None,
    changes: pd.DataFrame,
    rebuilt_sum: int,
) -> AnalysisResult:
    monthly_delta = new_monthly - old_monthly
    annual_delta = monthly_delta * 12
    monthly_rate = (monthly_delta / old_monthly * 100) if old_monthly else None

    if old_total is not None and new_total is not None:
        total_delta = new_total - old_total
        total_rate = (total_delta / old_total * 100) if old_total else None
    else:
        total_delta = None
        total_rate = None

    p_bucket = price_bucket(monthly_rate, monthly_delta)
    price_direction = "감소" if monthly_delta < 0 else "증가" if monthly_delta > 0 else "동일"
    total_direction = (
        "비교 안 함" if total_delta is None else "감소" if total_delta < 0 else "증가" if total_delta > 0 else "동일"
    )
    coverage_direction = detect_coverage_direction(changes)

    if price_direction == "감소" and coverage_direction == "강화":
        result_type = "효율 개선형"
    elif price_direction == "감소" and coverage_direction in {"유사/재배분", "정보 부족"}:
        result_type = "보험료 절감형"
    elif price_direction == "감소" and coverage_direction == "축소/조정":
        result_type = "부담 조정형"
    elif price_direction == "동일" and coverage_direction == "강화":
        result_type = "동일 예산 강화형"
    elif price_direction == "증가" and coverage_direction == "강화":
        result_type = "보장 강화형"
    elif coverage_direction == "강화와 조정 혼합":
        result_type = "맞춤 재구성형"
    else:
        result_type = "맞춤 재구성형"

    warnings: list[str] = []
    if old_monthly <= 0:
        warnings.append("기존 월 보험료가 입력되지 않아 증감률 문구를 제한합니다.")
    if new_monthly <= 0:
        warnings.append("리모델링 후 월 보험료를 확인해 주세요.")
    if rebuilt_sum and new_monthly and abs(rebuilt_sum - new_monthly) >= 1:
        warnings.append(
            f"보장 재구성 항목 합계({format_won(rebuilt_sum)})와 변경 월 보험료({format_won(new_monthly)})가 "
            f"{format_won(abs(rebuilt_sum - new_monthly))} 차이 납니다."
        )
    if price_direction == "증가" and coverage_direction not in {"강화", "강화와 조정 혼합"}:
        warnings.append("보험료가 증가하지만 강화 근거가 충분하지 않습니다. 강화 내역을 확인해 주세요.")
    if old_total is None or new_total is None:
        warnings.append("총 납입예정액을 입력하지 않아 관련 문구는 자동으로 제외됩니다.")
    if changes.empty:
        warnings.append("보장 변화가 입력되지 않아 보장 강화·범위 확대 문구를 자동 생성하지 않습니다.")

    return AnalysisResult(
        monthly_delta=monthly_delta,
        annual_delta=annual_delta,
        total_delta=total_delta,
        monthly_rate=monthly_rate,
        total_rate=total_rate,
        price_direction=price_direction,
        total_direction=total_direction,
        coverage_direction=coverage_direction,
        result_type=result_type,
        warnings=tuple(warnings),
    )


def change_sentence(row: pd.Series) -> str:
    name = normalize_text(row.get("담보명"))
    old = normalize_text(row.get("기존 보장"))
    new = normalize_text(row.get("변경 후 보장"))
    change_type = normalize_text(row.get("변화 유형"))
    manual = normalize_text(row.get("변화 설명"))
    if manual:
        return manual
    if not name:
        return ""
    if old and new:
        if change_type == "신규 보장 추가" and old in {"없음", "미가입", "-"}:
            return f"{name}: 기존에 없던 보장을 {new}로 새로 구성"
        return f"{name}: {old}에서 {new}로 {change_type}"
    if new:
        return f"{name}: {new}로 {change_type}"
    return f"{name}: {change_type}"


def highlighted_changes(changes: pd.DataFrame, max_items: int = 4) -> list[str]:
    if changes.empty:
        return []
    rows = changes.copy()
    if "첫 장 강조" in rows.columns and rows["첫 장 강조"].fillna(False).any():
        rows = rows[rows["첫 장 강조"].fillna(False)]
    return [s for s in rows.apply(change_sentence, axis=1).tolist() if s][:max_items]


def top_coverage_names(changes: pd.DataFrame, max_items: int = 3) -> list[str]:
    if changes.empty:
        return []
    rows = changes.copy()
    if "첫 장 강조" in rows.columns and rows["첫 장 강조"].fillna(False).any():
        rows = rows[rows["첫 장 강조"].fillna(False)]
    return [normalize_text(v) for v in rows["담보명"].tolist() if normalize_text(v)][:max_items]


def sentence_candidates(
    customer_name: str,
    analysis: AnalysisResult,
    changes: pd.DataFrame,
    design_goal: str,
    priorities: list[str],
    keep_existing: bool,
    old_monthly: int,
    new_monthly: int,
) -> dict[str, list[str]]:
    names = top_coverage_names(changes)
    coverage_phrase = "·".join(names) if names else "입력된 핵심 보장"
    delta = abs(analysis.monthly_delta)

    headline: list[str] = []
    recommendation: list[str] = []
    closing: list[str] = []

    if analysis.price_direction == "감소" and analysis.coverage_direction == "강화":
        headline.extend([
            f"월 보험료 부담은 낮추고 {coverage_phrase} 보장은 보완한 제안입니다.",
            f"월 {format_won(delta)}을 줄이면서 필요한 보장을 강화한 효율 개선안입니다.",
            f"기존 보험료는 낮추고, 부족했던 {coverage_phrase} 보장을 채우는 방향으로 재구성했습니다.",
        ])
    elif analysis.price_direction == "감소":
        headline.extend([
            f"월 보험료를 {format_won(delta)} 낮춰 장기적으로 유지하기 쉬운 구조로 조정한 제안입니다.",
            f"현재 보험료 부담을 줄이고 필요한 보장 중심으로 다시 정리한 제안입니다.",
            f"월 보험료를 {format_won(old_monthly)}에서 {format_won(new_monthly)}으로 조정한 절감형 제안입니다.",
        ])
    elif analysis.price_direction == "증가" and analysis.coverage_direction in {"강화", "강화와 조정 혼합"}:
        headline.extend([
            f"월 보험료는 {format_won(delta)} 늘어나지만, {coverage_phrase} 보장을 강화한 제안입니다.",
            f"월 {format_won(delta)}의 추가 예산으로 부족했던 {coverage_phrase} 보장을 보완합니다.",
            f"보험료 절감보다 {coverage_phrase}의 보장 공백을 줄이는 데 초점을 둔 제안입니다.",
        ])
    elif analysis.price_direction == "동일" and analysis.coverage_direction == "강화":
        headline.extend([
            f"현재와 비슷한 월 보험료 안에서 {coverage_phrase} 보장을 강화한 제안입니다.",
            "보험료 수준은 유지하면서 필요한 보장으로 배분을 바꾼 제안입니다.",
            "추가 부담 없이 핵심 보장의 구성을 개선하는 데 초점을 두었습니다.",
        ])
    else:
        headline.extend([
            "기존 보험과 변경안을 보험료와 보장 변화 기준으로 비교한 제안입니다.",
            "유지할 부분과 조정할 부분을 구분해 현재 상황에 맞게 재구성한 제안입니다.",
        ])

    if design_goal == "보험료 부담 완화":
        recommendation.extend([
            "현재 가장 중요한 기준이 월 보험료 부담이므로, 장기적으로 유지 가능한 수준을 우선해 구성했습니다.",
            "보험은 가입보다 유지가 중요하므로, 필요한 보장을 남기면서 월 부담을 조정하는 방향을 추천드립니다.",
        ])
    elif design_goal == "보장 강화":
        recommendation.extend([
            f"보험료 절감보다 부족했던 {coverage_phrase} 보장을 채우는 것이 중요한 상황이라 이 안을 추천드립니다.",
            f"추가 보험료가 발생하더라도 {coverage_phrase} 보장을 우선 보완하는 방향입니다.",
        ])
    elif design_goal == "동일 예산 재구성":
        recommendation.extend([
            "현재 예산을 크게 바꾸지 않으면서 보험료가 필요한 보장에 쓰이도록 재배분한 안입니다.",
            "월 부담은 유지하되 보장의 우선순위를 다시 정리하는 방향이 적절합니다.",
        ])
    elif design_goal == "특정 위험 집중":
        recommendation.extend([
            f"고객님이 중요하게 생각하신 {coverage_phrase} 영역을 우선해 집중적으로 구성했습니다.",
            f"전체 보장을 넓게 늘리기보다 {coverage_phrase}에 예산을 집중한 안입니다.",
        ])
    else:
        recommendation.extend([
            "보험료와 보장 수준을 함께 비교했을 때, 필요한 보장을 유지하면서 장기적으로 관리하기 쉬운 구성입니다.",
            "무조건 보험료를 줄이거나 보장을 늘리는 것이 아니라, 현재 필요한 부분에 맞춰 균형을 조정한 안입니다.",
        ])

    if priorities:
        priority_text = "·".join(priorities[:3])
        recommendation.insert(0, f"고객님께서 중요하게 말씀하신 {priority_text}를 우선 반영한 구성입니다.")

    if keep_existing:
        closing.append("기존 계약 중 유지 가치가 있는 보장은 남기고, 조정이 필요한 부분만 구분해 검토합니다.")
    closing.extend([
        "새로운 계약의 승인 조건을 확인한 뒤 기존 계약의 유지·감액·해지 여부를 결정하는 것이 안전합니다.",
        "보험료뿐 아니라 보장 범위, 보험기간, 납입기간과 면책·감액 조건을 함께 확인한 후 결정해야 합니다.",
        "이 안은 정답을 정해 놓은 자료가 아니라, 고객님의 예산과 우선순위에 맞는 선택을 돕기 위한 비교안입니다.",
    ])

    # 근거가 없는 단정 표현 방지
    if not names:
        headline = [s.replace(f"{coverage_phrase} 보장", "핵심 보장") for s in headline]
        recommendation = [s.replace(coverage_phrase, "핵심 보장") for s in recommendation]

    return {
        "제목 아래 한 줄": list(dict.fromkeys(headline))[:5],
        "추천 이유": list(dict.fromkeys(recommendation))[:5],
        "마무리·주의": list(dict.fromkeys(closing))[:5],
    }


def select_default(candidates: dict[str, list[str]]) -> dict[str, str]:
    return {key: values[0] if values else "" for key, values in candidates.items()}


def _set_cell(ws, cell: str, value: Any, *, font=None, fill=None, alignment=None, border=None, number_format=None):
    c = ws[cell]
    c.value = value
    if font:
        c.font = font
    if fill:
        c.fill = fill
    if alignment:
        c.alignment = alignment
    if border:
        c.border = border
    if number_format:
        c.number_format = number_format
    return c


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


def _style_all(ws, max_row: int, max_col: int):
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for c in row:
            c.font = c.font.copy(name="맑은 고딕")
            if c.alignment is None:
                c.alignment = Alignment(vertical="center")
    ws.sheet_view.showGridLines = False


def _page_setup(ws, fit_height: int = 1):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_height
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.4, bottom=0.4, header=0.15, footer=0.15)
    ws.print_options.horizontalCentered = True


def _section_title(ws, row: int, title: str, end_col: int = 10):
    _merge_write(
        ws,
        f"A{row}:{get_column_letter(end_col)}{row}",
        title,
        font=Font(name="맑은 고딕", size=12, bold=True, color=WHITE),
        fill=PatternFill("solid", fgColor=NAVY),
        alignment=Alignment(horizontal="left", vertical="center"),
        border=Border(bottom=MEDIUM_NAVY),
    )
    ws.row_dimensions[row].height = 24


def _write_table_header(ws, row: int, headers: list[str], widths: list[int] | None = None):
    for idx, header in enumerate(headers, start=1):
        c = ws.cell(row=row, column=idx, value=header)
        c.font = Font(name="맑은 고딕", size=10, bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=NAVY_DARK)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
        if widths:
            ws.column_dimensions[get_column_letter(idx)].width = widths[idx - 1]
    ws.row_dimensions[row].height = 25


def create_excel(
    *,
    customer_name: str,
    title: str,
    proposal_label: str,
    consultation_date: date,
    consultant: str,
    analysis: AnalysisResult,
    old_monthly: int,
    new_monthly: int,
    old_total: int | None,
    new_total: int | None,
    rebuild_df: pd.DataFrame,
    changes_df: pd.DataFrame,
    contract_df: pd.DataFrame,
    headline: str,
    recommendation: str,
    caution: str,
    output_type: str,
    core_rows: list[dict[str, str]],
) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "01_한눈에보기"

    # 1장: 고객 설득용 핵심 비교
    for col, width in {"A": 17, "B": 18, "C": 18, "D": 18, "E": 3, "F": 17, "G": 18, "H": 18, "I": 18, "J": 3}.items():
        ws.column_dimensions[col].width = width

    _merge_write(
        ws, "A1:J2", title,
        font=Font(name="맑은 고딕", size=22, bold=True, color=NAVY_DARK),
        alignment=Alignment(horizontal="left", vertical="center"),
    )
    ws.row_dimensions[1].height = 29
    ws.row_dimensions[2].height = 22

    _merge_write(
        ws, "A3:H3", headline,
        font=Font(name="맑은 고딕", size=11, bold=True, color=BLACK),
        fill=PatternFill("solid", fgColor=LIGHT_BLUE),
        alignment=Alignment(horizontal="left", vertical="center", wrap_text=True),
        border=Border(left=MEDIUM_NAVY),
    )
    _merge_write(
        ws, "I3:J3", proposal_label,
        font=Font(name="맑은 고딕", size=11, bold=True, color=WHITE),
        fill=PatternFill("solid", fgColor=TEAL if "강화" in analysis.result_type else GOLD),
        alignment=Alignment(horizontal="center", vertical="center"),
    )
    ws.row_dimensions[3].height = 36

    # 핵심 카드 3개
    cards: list[tuple[str, str, str]] = []
    if analysis.result_type in {"보장 강화형", "동일 예산 강화형"}:
        highlighted = highlighted_changes(changes_df, 2)
        cards.append(("월 보험료 변화", format_change(analysis.monthly_delta), LIGHT_GOLD))
        cards.append(("가장 큰 보장 변화", highlighted[0] if highlighted else "강화 내역 확인", LIGHT_TEAL))
        cards.append(("추가 변화", highlighted[1] if len(highlighted) > 1 else "상세표에서 확인", LIGHT_BLUE))
    else:
        cards.append(("월 보험료 변화", format_change(analysis.monthly_delta), LIGHT_TEAL))
        cards.append(("연간 보험료 변화", format_change(analysis.annual_delta), LIGHT_BLUE))
        cards.append(("총 납입예정액 변화", format_change(analysis.total_delta), LIGHT_GOLD))

    ranges = [("A5:C7"), ("D5:F7"), ("G5:J7")]
    for (label, value, fill_color), cell_range in zip(cards, ranges):
        start = cell_range.split(":")[0]
        _merge_write(
            ws, cell_range, f"{label}\n{value}",
            font=Font(name="맑은 고딕", size=14, bold=True, color=NAVY_DARK),
            fill=PatternFill("solid", fgColor=fill_color),
            alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
            border=Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY),
        )
        ws[start].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(5, 8):
        ws.row_dimensions[r].height = 24

    _section_title(ws, 9, "기존 보험과 리모델링 후 한눈에 비교", 10)
    headers = ["구분", "기존", "리모델링 후", "변화"]
    # 4개 열을 각각 2~3열씩 병합해 사용
    merges = [("A10:B10", headers[0]), ("C10:D10", headers[1]), ("E10:G10", headers[2]), ("H10:J10", headers[3])]
    for rng, text in merges:
        _merge_write(
            ws, rng, text,
            font=Font(name="맑은 고딕", size=10, bold=True, color=WHITE),
            fill=PatternFill("solid", fgColor=NAVY_DARK),
            alignment=Alignment(horizontal="center", vertical="center"),
            border=Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY),
        )
    ws.row_dimensions[10].height = 24

    comparison_rows = [
        {"구분": "월 보험료", "기존": format_won(old_monthly), "변경": format_won(new_monthly), "변화": format_change(analysis.monthly_delta)},
    ]
    if old_total is not None and new_total is not None:
        comparison_rows.append({
            "구분": "납입예정 총액",
            "기존": format_compact_won(old_total),
            "변경": format_compact_won(new_total),
            "변화": format_change(analysis.total_delta),
        })
    comparison_rows.extend(core_rows[: max(0, 5 - len(comparison_rows))])

    row = 11
    for item in comparison_rows[:5]:
        values = [
            ("A", "B", item.get("구분", ""), GRAY),
            ("C", "D", item.get("기존", ""), GRAY),
            ("E", "G", item.get("변경", item.get("리모델링 후", "")), LIGHT_BLUE),
            ("H", "J", item.get("변화", ""), LIGHT_TEAL),
        ]
        for start_col, end_col, value, fill_color in values:
            _merge_write(
                ws, f"{start_col}{row}:{end_col}{row}", value,
                font=Font(name="맑은 고딕", size=10, bold=start_col in {"E", "H"}, color=BLACK),
                fill=PatternFill("solid", fgColor=fill_color),
                alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
                border=Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY),
            )
        ws.row_dimensions[row].height = 28
        row += 1

    row += 1
    _section_title(ws, row, "이번에 좋아지거나 달라지는 점", 10)
    row += 1
    highlights = highlighted_changes(changes_df, 4)
    if not highlights:
        highlights = ["입력된 내용을 기준으로 보험료와 보장 구조를 비교했습니다."]
    for idx, text in enumerate(highlights, start=1):
        _merge_write(
            ws, f"A{row}:J{row}", f"{idx}. {text}",
            font=Font(name="맑은 고딕", size=10.5, bold=True if idx <= 2 else False, color=BLACK),
            fill=PatternFill("solid", fgColor=WHITE if idx % 2 else "F7F9FC"),
            alignment=Alignment(horizontal="left", vertical="center", wrap_text=True),
            border=Border(bottom=THIN_GRAY),
        )
        ws.row_dimensions[row].height = 27
        row += 1

    row += 1
    _section_title(ws, row, "이 안을 추천드리는 이유", 10)
    row += 1
    _merge_write(
        ws, f"A{row}:J{row+1}", recommendation,
        font=Font(name="맑은 고딕", size=10.5, color=BLACK),
        fill=PatternFill("solid", fgColor=LIGHT_BLUE),
        alignment=Alignment(horizontal="left", vertical="center", wrap_text=True),
        border=Border(left=MEDIUM_NAVY),
    )
    ws.row_dimensions[row].height = 28
    ws.row_dimensions[row + 1].height = 28
    row += 3

    _merge_write(
        ws, f"A{row}:J{row+1}", caution,
        font=Font(name="맑은 고딕", size=8.5, color=GRAY_DARK),
        fill=PatternFill("solid", fgColor="F7F8FA"),
        alignment=Alignment(horizontal="left", vertical="center", wrap_text=True),
        border=Border(top=THIN_GRAY),
    )
    ws.row_dimensions[row].height = 23
    ws.row_dimensions[row + 1].height = 23
    row += 2

    _merge_write(
        ws, f"A{row}:F{row}", f"상담일: {consultation_date.strftime('%Y-%m-%d')}",
        font=Font(name="맑은 고딕", size=8.5, color=GRAY_DARK),
        alignment=Alignment(horizontal="left"),
    )
    _merge_write(
        ws, f"G{row}:J{row}", f"담당: {consultant or '-'}",
        font=Font(name="맑은 고딕", size=8.5, color=GRAY_DARK),
        alignment=Alignment(horizontal="right"),
    )

    ws.print_area = f"A1:J{row}"
    _style_all(ws, row, 10)
    _page_setup(ws, 1)

    # 2장: 보장 변화 상세
    if output_type in {"표준형(2장)", "상세형(3장)"}:
        ws2 = wb.create_sheet("02_보장상세")
        widths = [25, 20, 20, 18, 31, 34]
        for i, width in enumerate(widths, 1):
            ws2.column_dimensions[get_column_letter(i)].width = width
        _merge_write(
            ws2, "A1:F2", f"{customer_name}님 보장 변화 상세",
            font=Font(name="맑은 고딕", size=20, bold=True, color=NAVY_DARK),
            alignment=Alignment(horizontal="left", vertical="center"),
        )
        _merge_write(
            ws2, "A3:F3", "정확한 담보명과 입력된 보장 내용을 기준으로 작성했습니다.",
            font=Font(name="맑은 고딕", size=10, color=BLACK),
            fill=PatternFill("solid", fgColor=LIGHT_BLUE),
            alignment=Alignment(horizontal="left", vertical="center"),
        )
        _write_table_header(ws2, 5, ["담보명", "기존 보장", "변경 후 보장", "변화 유형", "변화 설명", "구성 목적"], widths)
        r = 6
        if changes_df.empty:
            _merge_write(ws2, f"A{r}:F{r}", "입력된 보장 변화가 없습니다.", alignment=Alignment(horizontal="center"))
            r += 1
        else:
            for _, item in changes_df.iterrows():
                vals = [
                    normalize_text(item.get("담보명")), normalize_text(item.get("기존 보장")),
                    normalize_text(item.get("변경 후 보장")), normalize_text(item.get("변화 유형")),
                    change_sentence(item), normalize_text(item.get("구성 목적")),
                ]
                for cidx, value in enumerate(vals, 1):
                    c = ws2.cell(r, cidx, value)
                    c.font = Font(name="맑은 고딕", size=9.5, bold=cidx in {1, 3})
                    c.fill = PatternFill("solid", fgColor=LIGHT_BLUE if cidx == 3 else WHITE)
                    c.alignment = Alignment(horizontal="left" if cidx in {1, 5, 6} else "center", vertical="center", wrap_text=True)
                    c.border = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
                ws2.row_dimensions[r].height = 39
                r += 1

        r += 2
        _section_title(ws2, r, "리모델링 후 보험료 구성", 6)
        r += 1
        _write_table_header(ws2, r, ["보장 항목", "월 보험료", "보험회사", "구성 목적", "", ""], widths)
        # 마지막 두 헤더는 병합하여 표 폭 유지
        ws2.unmerge_cells(start_row=r, start_column=5, end_row=r, end_column=6) if False else None
        r += 1
        for _, item in rebuild_df.iterrows():
            vals = [normalize_text(item.get("보장 항목(담보명 포함 표현)")), _money(item.get("월 보험료")), normalize_text(item.get("보험회사")), normalize_text(item.get("구성 목적"))]
            merges2 = [(1, 1, vals[0]), (2, 2, vals[1]), (3, 3, vals[2]), (4, 6, vals[3])]
            for c1, c2, value in merges2:
                if c2 > c1:
                    ws2.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
                c = ws2.cell(r, c1, value)
                c.font = Font(name="맑은 고딕", size=9.5, bold=c1 in {1, 2})
                c.fill = PatternFill("solid", fgColor=LIGHT_TEAL if c1 == 2 else WHITE)
                c.alignment = Alignment(horizontal="right" if c1 == 2 else "left", vertical="center", wrap_text=True)
                for cell in ws2.iter_cols(min_col=c1, max_col=c2, min_row=r, max_row=r):
                    for cc in cell:
                        cc.border = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
                if c1 == 2:
                    c.number_format = '#,##0"원"'
            ws2.row_dimensions[r].height = 30
            r += 1
        if not rebuild_df.empty:
            ws2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=1)
            ws2.cell(r, 1, "합계").font = Font(name="맑은 고딕", size=10, bold=True, color=WHITE)
            ws2.cell(r, 1).fill = PatternFill("solid", fgColor=NAVY)
            ws2.cell(r, 2, sum(_money(v) for v in rebuild_df["월 보험료"])).number_format = '#,##0"원"'
            ws2.cell(r, 2).font = Font(name="맑은 고딕", size=10, bold=True, color=WHITE)
            ws2.cell(r, 2).fill = PatternFill("solid", fgColor=NAVY)
            ws2.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
            ws2.cell(r, 3, "").fill = PatternFill("solid", fgColor=NAVY)
            for c in range(1, 7):
                ws2.cell(r, c).border = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
            r += 1
        ws2.print_area = f"A1:F{r}"
        ws2.freeze_panes = "A6"
        _style_all(ws2, r, 6)
        _page_setup(ws2, 0)

    # 3장: 계약 처리와 확인사항
    if output_type == "상세형(3장)":
        ws3 = wb.create_sheet("03_계약정리")
        widths3 = [26, 21, 35, 35]
        for i, width in enumerate(widths3, 1):
            ws3.column_dimensions[get_column_letter(i)].width = width
        _merge_write(
            ws3, "A1:D2", f"{customer_name}님 기존 계약 정리 및 확인사항",
            font=Font(name="맑은 고딕", size=20, bold=True, color=NAVY_DARK),
            alignment=Alignment(horizontal="left", vertical="center"),
        )
        _merge_write(
            ws3, "A3:D3", "기존 계약의 조정은 새로운 계약의 승인 조건을 확인한 뒤 결정합니다.",
            font=Font(name="맑은 고딕", size=10, bold=True, color=BLACK),
            fill=PatternFill("solid", fgColor=LIGHT_GOLD),
            alignment=Alignment(horizontal="left", vertical="center"),
        )
        _write_table_header(ws3, 5, ["기존 계약/보장", "처리 방향", "판단 근거", "진행 조건"], widths3)
        r3 = 6
        if contract_df.empty:
            _merge_write(ws3, f"A{r3}:D{r3}", "입력된 계약 처리 내역이 없습니다.", alignment=Alignment(horizontal="center"))
            r3 += 1
        else:
            for _, item in contract_df.iterrows():
                vals = [normalize_text(item.get(c)) for c in ["기존 계약/보장", "처리 방향", "판단 근거", "진행 조건"]]
                for cidx, value in enumerate(vals, 1):
                    c = ws3.cell(r3, cidx, value)
                    c.font = Font(name="맑은 고딕", size=9.5, bold=cidx in {1, 2})
                    c.fill = PatternFill("solid", fgColor=LIGHT_BLUE if cidx == 2 else WHITE)
                    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                    c.border = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
                ws3.row_dimensions[r3].height = 38
                r3 += 1
        r3 += 2
        _section_title(ws3, r3, "반드시 확인할 내용", 4)
        r3 += 1
        checks = [
            "새로운 계약의 정상 승인 여부와 조건부 승인 내용을 먼저 확인합니다.",
            "기존 계약의 해지환급금, 납입 경과기간과 유지 가치를 함께 확인합니다.",
            "면책기간, 감액기간, 갱신 여부, 보험기간과 납입기간을 비교합니다.",
            "기존 계약을 먼저 해지해 보장 공백이 생기지 않도록 진행 순서를 확인합니다.",
        ]
        for idx, text in enumerate(checks, 1):
            _merge_write(
                ws3, f"A{r3}:D{r3}", f"{idx}. {text}",
                font=Font(name="맑은 고딕", size=10, color=BLACK),
                fill=PatternFill("solid", fgColor="F7F9FC" if idx % 2 == 0 else WHITE),
                alignment=Alignment(horizontal="left", vertical="center", wrap_text=True),
                border=Border(bottom=THIN_GRAY),
            )
            ws3.row_dimensions[r3].height = 29
            r3 += 1
        ws3.print_area = f"A1:D{r3}"
        ws3.freeze_panes = "A6"
        _style_all(ws3, r3, 4)
        _page_setup(ws3, 1)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def _safe_filename(name: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", name).strip()
    return cleaned or "보험리모델링_비교안"


def render_preview(
    customer_name: str,
    title: str,
    proposal_label: str,
    headline: str,
    analysis: AnalysisResult,
    old_monthly: int,
    new_monthly: int,
    old_total: int | None,
    new_total: int | None,
    changes_df: pd.DataFrame,
    recommendation: str,
):
    st.markdown(f"## {title}")
    st.caption(f"{proposal_label} · 자동 판정: {analysis.result_type}")
    st.info(headline or "입력 내용을 바탕으로 핵심 설명 문구가 표시됩니다.")

    c1, c2, c3 = st.columns(3)
    c1.metric("월 보험료", format_won(new_monthly), delta=format_change(analysis.monthly_delta))
    if analysis.price_direction == "증가" and analysis.coverage_direction in {"강화", "강화와 조정 혼합"}:
        highlights = highlighted_changes(changes_df, 2)
        c2.metric("핵심 보장 변화", highlights[0] if highlights else "입력 필요")
        c3.metric("추가 변화", highlights[1] if len(highlights) > 1 else "상세표 확인")
    else:
        c2.metric("연간 변화", format_change(analysis.annual_delta))
        c3.metric("총 납입액 변화", format_change(analysis.total_delta))

    rows = [{"구분": "월 보험료", "기존": format_won(old_monthly), "리모델링 후": format_won(new_monthly), "변화": format_change(analysis.monthly_delta)}]
    if old_total is not None and new_total is not None:
        rows.append({"구분": "납입예정 총액", "기존": format_compact_won(old_total), "리모델링 후": format_compact_won(new_total), "변화": format_change(analysis.total_delta)})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.markdown("### 이번에 좋아지거나 달라지는 점")
    highlights = highlighted_changes(changes_df, 4)
    if highlights:
        for text in highlights:
            st.markdown(f"- {text}")
    else:
        st.caption("보장 변화 내역을 입력하면 최대 4개가 표시됩니다.")

    st.markdown("### 이 안을 추천드리는 이유")
    st.write(recommendation or "추천 문구를 선택하거나 직접 수정해 주세요.")


def run():
    st.title(f"🔁 {APP_TITLE}")
    st.caption("고객이 30초 안에 보험료 변화, 보장 변화와 추천 이유를 이해하도록 만드는 비교표입니다.")

    if "remodel_rebuild_df" not in st.session_state:
        st.session_state.remodel_rebuild_df = default_rebuild_df()
    if "remodel_change_df" not in st.session_state:
        st.session_state.remodel_change_df = default_change_df()
    if "remodel_contract_df" not in st.session_state:
        st.session_state.remodel_contract_df = default_contract_df()

    tabs = st.tabs(["1. 기본정보", "2. 보험료 비교", "3. 보장 재구성", "4. 보장 변화", "5. 문구·미리보기·엑셀"])

    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1:
            customer_name = st.text_input("고객명", placeholder="예: 유민재", key="remodel_customer")
            proposal_type = st.selectbox(
                "설계 목적",
                ["보험료 부담 완화", "핵심 보장 보완", "동일 예산 재구성", "보장 강화", "특정 위험 집중", "기존 계약 정리", "맞춤 재구성"],
                key="remodel_goal",
            )
            proposal_label = st.text_input("제안 표시", value="균형 보장형 · 추천안", key="remodel_label")
        with c2:
            consultation_date = st.date_input("상담일", value=date.today(), key="remodel_date")
            consultant = st.text_input("담당자", placeholder="예: 박병선 팀장", key="remodel_consultant")
            output_type = st.radio("엑셀 출력", ["간단형(1장)", "표준형(2장)", "상세형(3장)"], horizontal=True, key="remodel_output")

        title_default = f"{customer_name}님 보험 리모델링 비교안" if customer_name else "보험 리모델링 비교안"
        title = st.text_input("자료 제목", value=title_default, key="remodel_title")
        priorities = st.multiselect(
            "고객이 중요하게 생각하는 부분",
            ["월 보험료 부담", "암 보장", "뇌·심장 보장", "치료비", "간병비", "사망보장", "노후 의료비", "납입기간", "갱신 부담", "기존 계약 최대한 유지"],
            key="remodel_priorities",
        )
        keep_existing = st.checkbox("기존 계약 중 유지 가치가 있는 보장을 남기는 방향", value=True, key="remodel_keep")

    with tabs[1]:
        st.subheader("보험료와 총 납입예정액")
        c1, c2 = st.columns(2)
        with c1:
            old_monthly = st.number_input("기존 월 보험료", min_value=0, step=1000, format="%d", key="remodel_old_month")
            use_total = st.checkbox("총 납입예정액 비교", value=True, key="remodel_use_total")
            old_total_raw = st.number_input("기존 납입예정 총액", min_value=0, step=100_000, format="%d", key="remodel_old_total", disabled=not use_total)
        with c2:
            new_monthly_manual = st.number_input("리모델링 후 월 보험료", min_value=0, step=1000, format="%d", key="remodel_new_month")
            new_total_raw = st.number_input("변경 납입예정 총액", min_value=0, step=100_000, format="%d", key="remodel_new_total", disabled=not use_total)
        st.caption("보장 재구성 항목 합계와 최종 월 보험료가 다를 수 있으므로 직접 입력값을 기준으로 비교하며, 차액은 자동 점검합니다.")

    with tabs[2]:
        st.subheader("리모델링 후 보험료 구성")
        st.caption("고객용 항목명은 정확한 담보명에 실제 포함된 단어만 사용해 주세요.")
        rebuild_edited = st.data_editor(
            st.session_state.remodel_rebuild_df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "보장 항목(담보명 포함 표현)": st.column_config.TextColumn(required=True, width="large"),
                "월 보험료": st.column_config.NumberColumn(min_value=0, step=1000, format="%d원"),
                "보험회사": st.column_config.TextColumn(width="medium"),
                "구성 목적": st.column_config.TextColumn(width="large"),
                "첫 장 강조": st.column_config.CheckboxColumn(),
            },
            key="remodel_rebuild_editor",
        )
        st.session_state.remodel_rebuild_df = rebuild_edited
        rebuild_df = clean_df(rebuild_edited, "보장 항목(담보명 포함 표현)")
        rebuilt_sum = sum(_money(v) for v in rebuild_df.get("월 보험료", []))
        st.metric("보장 항목 합계", format_won(rebuilt_sum))

    with tabs[3]:
        st.subheader("무엇이 달라지는지")
        st.caption("첫 장에는 중요 표시한 항목 중 최대 4개만 보여주고, 나머지는 상세 시트로 넘깁니다.")
        change_types = [
            "가입금액 증액", "보장 범위 확대", "신규 보장 추가", "보장기간 연장", "지급 횟수 확대",
            "특정 보장 집중 보완", "갱신 부담 조정", "유지", "가입금액 축소", "보장 범위 축소",
            "보장기간 단축", "보장 삭제/정리", "기타 직접 입력",
        ]
        change_edited = st.data_editor(
            st.session_state.remodel_change_df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "담보명": st.column_config.TextColumn(required=True, width="large"),
                "기존 보장": st.column_config.TextColumn(width="medium"),
                "변경 후 보장": st.column_config.TextColumn(width="medium"),
                "변화 유형": st.column_config.SelectboxColumn(options=change_types, required=True),
                "변화 설명": st.column_config.TextColumn(help="비워두면 담보명·기존·변경·변화 유형으로 자동 작성합니다.", width="large"),
                "구성 목적": st.column_config.TextColumn(width="large"),
                "첫 장 강조": st.column_config.CheckboxColumn(),
            },
            key="remodel_change_editor",
        )
        st.session_state.remodel_change_df = change_edited
        changes_df = clean_df(change_edited, "담보명")

        if not changes_df.empty:
            invalid_manual = []
            for idx, row in changes_df.iterrows():
                manual = normalize_text(row.get("변화 설명"))
                name = normalize_text(row.get("담보명"))
                if manual and not contains_only_words_from_name(manual, name):
                    # 전체 문장은 숫자/조사도 포함하므로 경고만 제공하고 차단하지 않음.
                    invalid_manual.append(idx + 1)
            if invalid_manual:
                st.info("변화 설명은 완성 문장이므로 담보명 외 단어가 포함될 수 있습니다. 다만 보장 명칭 자체는 입력한 담보명을 그대로 사용합니다.")

        st.divider()
        st.subheader("기존 계약 처리 방향 · 상세형에서 사용")
        contract_edited = st.data_editor(
            st.session_state.remodel_contract_df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "기존 계약/보장": st.column_config.TextColumn(required=True, width="large"),
                "처리 방향": st.column_config.SelectboxColumn(options=["유지", "감액 검토", "일부 특약 조정", "해지 검토", "신규 승인 후 결정", "추가 확인 필요"]),
                "판단 근거": st.column_config.TextColumn(width="large"),
                "진행 조건": st.column_config.TextColumn(width="large"),
            },
            key="remodel_contract_editor",
        )
        st.session_state.remodel_contract_df = contract_edited

    # 탭 밖에서도 안전하게 값 확보
    rebuild_df = clean_df(st.session_state.remodel_rebuild_df, "보장 항목(담보명 포함 표현)")
    changes_df = clean_df(st.session_state.remodel_change_df, "담보명")
    contract_df = clean_df(st.session_state.remodel_contract_df, "기존 계약/보장")
    rebuilt_sum = sum(_money(v) for v in rebuild_df.get("월 보험료", []))
    new_monthly = _money(st.session_state.get("remodel_new_month", 0))
    old_monthly = _money(st.session_state.get("remodel_old_month", 0))
    use_total = bool(st.session_state.get("remodel_use_total", True))
    old_total = _money(st.session_state.get("remodel_old_total", 0)) if use_total else None
    new_total = _money(st.session_state.get("remodel_new_total", 0)) if use_total else None
    if use_total and old_total == 0 and new_total == 0:
        old_total = None
        new_total = None

    analysis = analyze(old_monthly, new_monthly, old_total, new_total, changes_df, rebuilt_sum)
    candidates = sentence_candidates(
        customer_name=st.session_state.get("remodel_customer", ""),
        analysis=analysis,
        changes=changes_df,
        design_goal=st.session_state.get("remodel_goal", "맞춤 재구성"),
        priorities=st.session_state.get("remodel_priorities", []),
        keep_existing=bool(st.session_state.get("remodel_keep", True)),
        old_monthly=old_monthly,
        new_monthly=new_monthly,
    )
    defaults = select_default(candidates)

    with tabs[4]:
        st.subheader("자동 판정과 문구 후보")
        cols = st.columns(4)
        cols[0].metric("실제 결과", analysis.result_type)
        cols[1].metric("보험료 방향", price_bucket(analysis.monthly_rate, analysis.monthly_delta))
        cols[2].metric("보장 방향", analysis.coverage_direction)
        cols[3].metric("월 변화", format_change(analysis.monthly_delta))

        for warning in analysis.warnings:
            st.warning(warning)

        st.markdown("#### 제목 아래 한 줄")
        headline_choice = st.radio("후보 선택", candidates["제목 아래 한 줄"], key="remodel_headline_choice", label_visibility="collapsed") if candidates["제목 아래 한 줄"] else ""
        headline = st.text_area("최종 한 줄", value=headline_choice or defaults["제목 아래 한 줄"], height=80, key="remodel_headline_final")

        st.markdown("#### 추천 이유")
        recommendation_choice = st.radio("추천 이유 후보", candidates["추천 이유"], key="remodel_recommend_choice", label_visibility="collapsed") if candidates["추천 이유"] else ""
        recommendation = st.text_area("최종 추천 이유", value=recommendation_choice or defaults["추천 이유"], height=100, key="remodel_recommend_final")

        st.markdown("#### 마무리·주의")
        caution_choice = st.radio("주의 문구 후보", candidates["마무리·주의"], key="remodel_caution_choice", label_visibility="collapsed") if candidates["마무리·주의"] else ""
        caution_default = (
            f"{caution_choice or defaults['마무리·주의']} 본 자료는 입력된 계약정보를 기준으로 작성되며, "
            "실제 보장내용과 가입 조건은 보험회사·상품·담보·심사 결과 및 약관에 따라 달라질 수 있습니다."
        )
        caution = st.text_area("최종 주의 문구", value=caution_default, height=100, key="remodel_caution_final")

        st.divider()
        st.subheader("첫 장 비교표 추가 행")
        core_df_default = pd.DataFrame([
            {"구분": "", "기존": "", "리모델링 후": "", "변화": ""},
            {"구분": "", "기존": "", "리모델링 후": "", "변화": ""},
            {"구분": "", "기존": "", "리모델링 후": "", "변화": ""},
        ])
        core_df = st.data_editor(core_df_default, num_rows="dynamic", use_container_width=True, hide_index=True, key="remodel_core_rows")
        core_rows = [
            {k: normalize_text(v) for k, v in row.items()}
            for row in core_df.to_dict("records")
            if normalize_text(row.get("구분"))
        ]

        st.divider()
        st.subheader("고객 시점 미리보기")
        render_preview(
            customer_name=st.session_state.get("remodel_customer", ""),
            title=st.session_state.get("remodel_title", "보험 리모델링 비교안"),
            proposal_label=st.session_state.get("remodel_label", ""),
            headline=headline,
            analysis=analysis,
            old_monthly=old_monthly,
            new_monthly=new_monthly,
            old_total=old_total,
            new_total=new_total,
            changes_df=changes_df,
            recommendation=recommendation,
        )

        can_download = bool(st.session_state.get("remodel_customer", "").strip()) and old_monthly > 0 and new_monthly > 0
        if not can_download:
            st.error("엑셀을 만들려면 고객명, 기존 월 보험료와 리모델링 후 월 보험료를 입력해 주세요.")
        else:
            excel_bytes = create_excel(
                customer_name=st.session_state.get("remodel_customer", ""),
                title=st.session_state.get("remodel_title", "보험 리모델링 비교안"),
                proposal_label=st.session_state.get("remodel_label", ""),
                consultation_date=st.session_state.get("remodel_date", date.today()),
                consultant=st.session_state.get("remodel_consultant", ""),
                analysis=analysis,
                old_monthly=old_monthly,
                new_monthly=new_monthly,
                old_total=old_total,
                new_total=new_total,
                rebuild_df=rebuild_df,
                changes_df=changes_df,
                contract_df=contract_df,
                headline=headline,
                recommendation=recommendation,
                caution=caution,
                output_type=st.session_state.get("remodel_output", "간단형(1장)"),
                core_rows=core_rows,
            )
            output_name = st.session_state.get("remodel_output", "간단형(1장)").split("(")[0]
            filename = _safe_filename(
                f"{st.session_state.get('remodel_customer', '')}님_보험리모델링_{output_name}_{date.today().strftime('%Y%m%d')}.xlsx"
            )
            st.download_button(
                "📥 고객용 엑셀 다운로드",
                data=excel_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

        with st.expander("app.py 연결 확인"):
            st.code(
                'from modules import remodeling\n\nall_apps = {\n    "🔁 보험 리모델링 전/후 비교": remodeling.run,\n}',
                language="python",
            )
            st.caption("확인한 기존 app.py에는 위 연결이 이미 포함되어 있습니다. remodeling.py를 modules 폴더에 교체하면 됩니다.")
