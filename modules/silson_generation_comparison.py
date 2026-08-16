from __future__ import annotations

from datetime import date
from io import BytesIO
from pathlib import Path
from typing import Dict, Tuple

import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

try:
    from .ui_components import page_header, section_intro
except ImportError:  # 단독 점검용
    from ui_components import page_header, section_intro


STANDARD_DATE = "2026.08"
FIFTH_RATES = {"급여": 20.0, "중증 비급여": 30.0, "비중증 비급여": 50.0}
CLAIM_LEVELS = ["없음", "50만원 미만", "50~100만원", "100~300만원", "300만원 이상", "직접 입력"]
PREMIUM_MODES = ["연령별 예상 보험료", "가입제안서 직접 입력"]

# 공개된 대표 보험료 예시를 상담용 곡선으로 환산하기 위한 기준값입니다.
# 보험료 환경이 바뀌면 이 값만 수정할 수 있도록 한곳에 모았습니다.
REFERENCE_PREMIUM = {
    "남성": {"age_40_full": 17_000, "annual_factor": 1.038},
    "여성": {"age_40_full": 19_000, "annual_factor": 1.038},
}
CURRENT_PREMIUM_REFERENCE = {
    "남성": {"1세대": 54_300, "2세대": 33_700, "3세대": 22_000, "4세대": 14_600},
    "여성": {"1세대": 60_000, "2세대": 40_000, "3세대": 24_000, "4세대": 16_000},
}
CURRENT_PREMIUM_AGE_FACTOR = {"1세대": 1.048, "2세대": 1.047, "3세대": 1.043, "4세대": 1.040}


def won(value: float) -> str:
    return f"{int(round(value)):,}원"


def compact_manwon(value: float) -> str:
    amount = value / 10_000
    return f"{amount:,.1f}".rstrip("0").rstrip(".") + "만원"


def safe_filename(value: str) -> str:
    cleaned = "".join("_" if ch in '\\/:*?\"<>|' else ch for ch in value.strip())
    return cleaned or "OOO"


def current_rates(generation: str, option: str) -> Dict[str, float]:
    if generation == "1세대":
        rate = float(option.replace("%", ""))
        return {"급여": rate, "중증 비급여": rate, "비중증 비급여": rate}
    if generation == "2세대":
        rate = float(option.replace("형", "").replace("%", ""))
        return {"급여": rate, "중증 비급여": rate, "비중증 비급여": rate}
    if generation == "3세대":
        salary = float(option.replace("급여 ", "").replace("형", "").replace("%", ""))
        return {"급여": salary, "중증 비급여": 20.0, "비중증 비급여": 20.0}
    return {"급여": 20.0, "중증 비급여": 30.0, "비중증 비급여": 30.0}


def calculate(medical: Dict[str, float], rates: Dict[str, float]) -> Tuple[float, float]:
    covered = sum(medical[key] for key in ("급여", "중증 비급여", "비중증 비급여"))
    burden = sum(medical[key] * rates[key] / 100 for key in rates)
    excluded = medical["보상 제외 가능 비급여"]
    return max(0.0, covered - burden), burden + excluded


def estimate_fifth_premium(age: int, gender: str) -> int:
    """공개된 대표 연령 보험료를 바탕으로 1세 단위 상담용 참고값을 계산합니다."""
    reference = REFERENCE_PREMIUM[gender]
    full_premium = reference["age_40_full"] * (reference["annual_factor"] ** (age - 40))
    return max(1_000, int(round(full_premium / 100) * 100))


def estimate_current_premium(age: int, gender: str, generation: str, option: str) -> int:
    """세대·연령·성별·계약유형을 반영한 현재 실손의 상담용 참고 보험료입니다."""
    base = CURRENT_PREMIUM_REFERENCE[gender][generation]
    age_adjusted = base * (CURRENT_PREMIUM_AGE_FACTOR[generation] ** (age - 40))
    option_factor = 1.0
    if generation == "1세대":
        option_factor = {"0%": 1.0, "10%": 0.92, "20%": 0.85}.get(option, 1.0)
    elif generation == "2세대" and option == "20%형":
        option_factor = 0.85
    elif generation == "3세대" and option == "급여 20%형":
        option_factor = 0.90
    return max(1_000, int(round(age_adjusted * option_factor / 100) * 100))


def _mark_reference_premium_modified() -> None:
    st.session_state["sc_reference_modified"] = True


def _mark_current_premium_modified() -> None:
    st.session_state["sc_current_modified"] = True


def inject_styles() -> None:
    st.markdown(
        """
        <style>
        .sc-card{padding:1.15rem 1.25rem;border:1px solid #DCE6EE;border-radius:16px;background:rgba(255,255,255,.94);box-shadow:0 11px 30px rgba(37,72,98,.055)}
        .sc-rate-grid{display:grid;grid-template-columns:1.3fr 1fr 1fr;border:1px solid #DCE6EE;border-radius:14px;overflow:hidden;background:#fff}
        .sc-rate-grid>div{padding:.78rem .9rem;border-bottom:1px solid #E8EEF3;text-align:center;font-size:.88rem}
        .sc-rate-grid>div:nth-last-child(-n+3){border-bottom:0}.sc-rate-grid .head{background:#F2F7FB;color:#516A7E;font-size:.76rem;font-weight:800}
        .sc-rate-grid .label{text-align:left;color:#40596D;font-weight:700}.sc-current{color:#1769DC;font-weight:850}.sc-fifth{color:#119B98;font-weight:850}
        .sc-bars{display:grid;gap:.8rem}.sc-bar-row{display:grid;grid-template-columns:8.2rem 1fr 6.2rem;align-items:center;gap:.75rem;font-size:.82rem;color:#536B7E}
        .sc-track{height:1rem;border-radius:999px;background:#EAF0F4;overflow:hidden}.sc-fill-blue{height:100%;background:linear-gradient(90deg,#1769DC,#5B96E8);border-radius:999px}.sc-fill-teal{height:100%;background:linear-gradient(90deg,#119B98,#56BFBA);border-radius:999px}
        .sc-value{text-align:right;color:#203B50;font-weight:800}.sc-diff{margin-top:1rem;padding:.9rem 1rem;text-align:center;border-radius:12px;background:#EDF5FF;color:#174D91;font-weight:850}
        .sc-total{display:flex;align-items:end;justify-content:space-between;gap:1rem;margin:.9rem 0 1rem;padding:.85rem 1rem;border-radius:12px;background:#F3F7FA;color:#5B7182;font-size:.78rem}
        .sc-total strong{color:#17364E;font-size:1.18rem}.sc-stack-wrap{display:grid;gap:1.15rem}.sc-stack-row{display:grid;grid-template-columns:7.2rem 1fr;align-items:center;gap:.8rem}
        .sc-stack-name{color:#506A7D;font-size:.8rem;font-weight:750}.sc-stack-content{min-width:0}.sc-stack-values{display:flex;align-items:center;justify-content:space-between;gap:.65rem;margin-bottom:.38rem;font-size:.72rem;font-weight:850;white-space:nowrap}.sc-stack-insurance{color:#1769DC}.sc-stack-insurance-fifth{color:#119B98}.sc-stack-self{color:#D96C2D}.sc-stack-excluded-text{color:#71808C}
        .sc-stack{display:flex;height:1.7rem;overflow:hidden;border-radius:9px;background:#E8EEF2;box-shadow:inset 0 0 0 1px rgba(61,91,112,.06)}
        .sc-stack-part{height:100%}.sc-stack-current{background:linear-gradient(90deg,#1769DC,#4B8AE5)}.sc-stack-fifth{background:linear-gradient(90deg,#119B98,#4EBAB5)}.sc-stack-burden{background:linear-gradient(90deg,#E88B3D,#D96C2D)}.sc-stack-excluded{background:linear-gradient(90deg,#A6B0B8,#7D8993)}
        .sc-example-note{display:flex;align-items:center;min-height:2.75rem;padding:.55rem .8rem;border:1px solid #DDE9F1;border-radius:10px;background:#F8FBFD;color:#60788A;font-size:.78rem}.sc-example-note b{margin-left:.25rem;color:#244C68;font-size:.86rem}
        .sc-basis{margin-top:.6rem;padding:.7rem .85rem;border:1px solid #DDE9F1;border-radius:10px;background:#F8FBFD;color:#687F91;font-size:.74rem;line-height:1.5}
        .sc-input-title{display:flex;align-items:center;justify-content:space-between;gap:.7rem;margin:0 0 .9rem;color:#17364E;font-size:1rem;font-weight:850}.sc-input-title:before{content:"";width:4px;height:18px;border-radius:99px;background:linear-gradient(#1769DC,#119B98)}
        .sc-input-title span:first-child{flex:1}.sc-subtitle{display:flex;align-items:center;justify-content:space-between;margin:.1rem 0 .75rem;color:#36566E;font-size:.84rem;font-weight:850}
        .sc-fixed-badge{display:inline-flex;align-items:center;padding:.25rem .55rem;border:1px solid #BCE1DE;border-radius:999px;background:#EAF8F7;color:#147C79;font-size:.65rem;font-weight:850}
        .sc-divider{height:1px;margin:1.1rem 0;background:linear-gradient(90deg,transparent,#D7E4EC 10%,#D7E4EC 90%,transparent)}
        .sc-cumulative-card{padding:1.05rem 1.1rem;border:1px solid #DCE6EE;border-radius:15px;background:#fff;box-shadow:0 10px 28px rgba(37,72,98,.055)}
        .sc-cumulative-year{margin-bottom:.8rem;color:#29485F;font-size:.92rem;font-weight:850}.sc-cumulative-row{display:flex;align-items:center;justify-content:space-between;gap:.65rem;padding:.48rem .62rem;border-radius:9px;font-size:.78rem;font-weight:750}
        .sc-cumulative-row+.sc-cumulative-row{margin-top:.4rem}.sc-cumulative-current{background:#EDF4FF;color:#175FBF}.sc-cumulative-fifth{background:#EAF8F7;color:#117C79}.sc-cumulative-row strong{font-size:1.04rem;letter-spacing:-.025em;white-space:nowrap}.sc-cumulative-diff{display:flex;align-items:center;justify-content:space-between;gap:.6rem;margin-top:.65rem;padding:.55rem .65rem;border-top:1px solid #E1EAF0;color:#556F82;font-size:.76rem;font-weight:750}.sc-cumulative-diff strong{color:#173E62;font-size:.96rem;white-space:nowrap}
        .sc-cum-chart{margin-top:1rem;padding:1.15rem 1.25rem;border:1px solid #DCE6EE;border-radius:16px;background:#fff;box-shadow:0 11px 30px rgba(37,72,98,.055)}.sc-cum-chart-head{display:flex;align-items:center;justify-content:space-between;gap:1rem;margin-bottom:1rem}.sc-cum-chart-title{color:#29485F;font-size:.92rem;font-weight:850}.sc-cum-legend{display:flex;gap:1rem;color:#63798A;font-size:.75rem;font-weight:750}.sc-cum-dot{display:inline-block;width:.55rem;height:.55rem;margin-right:.3rem;border-radius:50%}.sc-cum-dot-current{background:#1769DC}.sc-cum-dot-fifth{background:#119B98}
        .sc-cum-plot{display:flex;align-items:end;gap:1.4rem;min-height:16rem;padding:.8rem .6rem .2rem;border-bottom:1px solid #DCE6EE}.sc-cum-group{flex:1;display:flex;align-items:end;justify-content:center;gap:1rem;height:14.5rem;position:relative;padding-bottom:3rem}.sc-cum-bar-col{width:min(5.2rem,38%);height:11rem;display:flex;flex-direction:column;justify-content:end;align-items:center}.sc-cum-bar-value{margin-bottom:.35rem;color:#29485F;font-size:.79rem;font-weight:850;white-space:nowrap}.sc-cum-bar{width:100%;min-height:.28rem;border-radius:9px 9px 3px 3px}.sc-cum-bar-current{background:linear-gradient(180deg,#4C8CE8,#1769DC)}.sc-cum-bar-fifth{background:linear-gradient(180deg,#50BDB8,#119B98)}.sc-cum-year{position:absolute;bottom:1.45rem;color:#3F5C71;font-size:.85rem;font-weight:850}.sc-cum-difference{position:absolute;bottom:.05rem;padding:.28rem .55rem;border-radius:999px;background:#EFF4F8;color:#244C6B;font-size:.75rem;font-weight:850;white-space:nowrap}
        .sc-note{margin-top:.65rem;color:#718393;font-size:.74rem;line-height:1.55}
        @media(max-width:700px){.sc-bar-row{grid-template-columns:6.5rem 1fr}.sc-value{grid-column:2}.sc-rate-grid>div{padding:.65rem .4rem;font-size:.76rem}.sc-stack-row{grid-template-columns:1fr}.sc-stack-values{font-size:.66rem;gap:.35rem}.sc-cum-chart{overflow-x:auto}.sc-cum-plot{min-width:42rem}.sc-cum-chart-head{min-width:42rem}}
        </style>
        """,
        unsafe_allow_html=True,
    )


def rate_table(generation: str, rates: Dict[str, float]) -> None:
    labels = ["급여 입원 자기부담", "중증 비급여 자기부담", "비중증 비급여 자기부담"]
    keys = ["급여", "중증 비급여", "비중증 비급여"]
    cells = ['<div class="head">비교 항목</div>', f'<div class="head">현재 {generation}</div>', '<div class="head">5세대 실손</div>']
    for label, key in zip(labels, keys):
        old_prefix = "구분 없음 · " if generation in ("1세대", "2세대", "3세대") and key != "급여" else ""
        cells.extend([
            f'<div class="label">{label}</div>',
            f'<div class="sc-current">{old_prefix}{rates[key]:g}%</div>',
            f'<div class="sc-fifth">{FIFTH_RATES[key]:g}%</div>',
        ])
    st.markdown('<div class="sc-rate-grid">' + "".join(cells) + '</div>', unsafe_allow_html=True)


def comparison_bars(title: str, current_value: float, fifth_value: float, diff_label: str) -> None:
    maximum = max(current_value, fifth_value, 1)
    rows = [
        ("현재 실손", current_value, "sc-fill-blue"),
        ("5세대 실손", fifth_value, "sc-fill-teal"),
    ]
    html_rows = "".join(
        f'<div class="sc-bar-row"><span>{name}</span><div class="sc-track"><div class="{css}" style="width:{value / maximum * 100:.1f}%"></div></div><span class="sc-value">{won(value)}</span></div>'
        for name, value, css in rows
    )
    st.markdown(f'<div class="sc-card"><b>{title}</b><div class="sc-bars" style="margin-top:1rem">{html_rows}</div><div class="sc-diff">{diff_label}</div></div>', unsafe_allow_html=True)


def medical_stacked_bars(
    total: float,
    generation: str,
    current_payout: float,
    current_burden: float,
    fifth_payout: float,
    fifth_burden: float,
    excluded: float,
) -> None:
    denominator = max(total, 1)
    rows = []
    for name, payout, burden, css in (
        (f"현재 {generation}", current_payout, current_burden, "sc-stack-current"),
        ("5세대 실손", fifth_payout, fifth_burden, "sc-stack-fifth"),
    ):
        payout_width = max(0.0, min(100.0, payout / denominator * 100))
        covered_burden = max(0.0, burden - excluded)
        burden_width = max(0.0, min(100.0, covered_burden / denominator * 100))
        excluded_width = max(0.0, min(100.0, excluded / denominator * 100))
        insurance_class = "sc-stack-insurance-fifth" if css == "sc-stack-fifth" else "sc-stack-insurance"
        rows.append(
            f'<div class="sc-stack-row"><span class="sc-stack-name">{name}</span>'
            f'<div class="sc-stack-content"><div class="sc-stack-values">'
            f'<span class="{insurance_class}">예상 보험금 {won(payout)}</span>'
            f'<span class="sc-stack-self">본인 부담 {won(covered_burden)}</span>'
            f'<span class="sc-stack-excluded-text">보상 제외 {won(excluded)}</span></div>'
            f'<div class="sc-stack"><div class="sc-stack-part {css}" style="width:{payout_width:.2f}%"></div>'
            f'<div class="sc-stack-part sc-stack-burden" style="width:{burden_width:.2f}%"></div>'
            f'<div class="sc-stack-part sc-stack-excluded" style="width:{excluded_width:.2f}%"></div></div></div></div>'
        )
    st.markdown(
        '<div class="sc-card"><b>총 의료비와 고객 부담 비교</b>'
        f'<div class="sc-total"><span>동일한 입원·수술 사례의 총 의료비</span><strong>{won(total)}</strong></div>'
        f'<div class="sc-stack-wrap">{"".join(rows)}</div>'
        f'<div class="sc-diff">고객 부담 차이 {won(abs(current_burden-fifth_burden))}</div></div>',
        unsafe_allow_html=True,
    )


def cumulative_premium_chart(current_premium: float, fifth_premium: float) -> None:
    values = [(years, current_premium * 12 * years, fifth_premium * 12 * years) for years in range(1, 6)]
    maximum = max((max(current, fifth) for _, current, fifth in values), default=1) or 1
    groups = []
    for years, current, fifth in values:
        current_height = max(2.0, current / maximum * 100)
        fifth_height = max(2.0, fifth / maximum * 100)
        groups.append(
            '<div class="sc-cum-group">'
            f'<div class="sc-cum-bar-col"><span class="sc-cum-bar-value">{won(current)}</span><div class="sc-cum-bar sc-cum-bar-current" style="height:{current_height:.2f}%"></div></div>'
            f'<div class="sc-cum-bar-col"><span class="sc-cum-bar-value">{won(fifth)}</span><div class="sc-cum-bar sc-cum-bar-fifth" style="height:{fifth_height:.2f}%"></div></div>'
            f'<span class="sc-cum-year">{years}년</span><span class="sc-cum-difference">차이 {won(abs(current-fifth))}</span></div>'
        )
    st.markdown(
        '<div class="sc-cum-chart"><div class="sc-cum-chart-head"><span class="sc-cum-chart-title">누적 보험료 비교 그래프</span>'
        '<span class="sc-cum-legend"><span><i class="sc-cum-dot sc-cum-dot-current"></i>현재 실손</span><span><i class="sc-cum-dot sc-cum-dot-fifth"></i>5세대 실손</span></span></div>'
        f'<div class="sc-cum-plot">{"".join(groups)}</div></div>',
        unsafe_allow_html=True,
    )


def build_pdf(data: dict) -> bytes:
    output = BytesIO()
    page_w, page_h = landscape(A4)
    c = canvas.Canvas(output, pagesize=(page_w, page_h))
    font_paths = [
        Path(__file__).resolve().parent.parent / "assets" / "fonts" / "PretendardVariable.ttf",
        Path(__file__).resolve().parent / "assets" / "fonts" / "PretendardVariable.ttf",
    ]
    font_path = next((path for path in font_paths if path.is_file()), None)
    try:
        if font_path:
            pdfmetrics.registerFont(TTFont("PretendardPDF", str(font_path)))
            font = "PretendardPDF"
        else:
            pdfmetrics.registerFont(UnicodeCIDFont("HYSMyeongJo-Medium"))
            font = "HYSMyeongJo-Medium"
    except Exception:
        font = "Helvetica"

    navy, blue, teal, muted, line = colors.HexColor("#16324F"), colors.HexColor("#2D6EAD"), colors.HexColor("#2A918C"), colors.HexColor("#687F91"), colors.HexColor("#DCE6EE")

    def text(x, y, value, size=9, color=navy, bold=False):
        c.setFillColor(color); c.setFont(font, size); c.drawString(x, y, str(value))

    c.setFillColor(colors.HexColor("#F6F9FC")); c.rect(0, 0, page_w, page_h, fill=1, stroke=0)
    c.setFillColor(colors.white); c.roundRect(13*mm, 11*mm, page_w-26*mm, page_h-22*mm, 5*mm, fill=1, stroke=0)
    text(21*mm, page_h-24*mm, f"{data['customer']}님 실손보험 세대 비교 안내", 22, navy)
    text(21*mm, page_h-33*mm, f"현재 {data['generation']} 실손과 5세대 실손의 보험료·입원 보장을 간단히 비교했습니다.", 11, muted)
    if data["consultant"]:
        c.drawRightString(page_w-21*mm, page_h-25*mm, f"담당자  {data['consultant']}")

    # 월 보험료 비교: 막대 길이와 정확한 금액을 함께 보여주는 메인 차트
    premium_x, premium_y, premium_w, premium_h = 21*mm, page_h-62*mm, 237*mm, 27*mm
    c.setFillColor(colors.HexColor("#F7FAFD"))
    c.roundRect(premium_x, premium_y, premium_w, premium_h, 4*mm, fill=1, stroke=0)
    text(premium_x+5*mm, premium_y+19*mm, "월 보험료 비교", 12.5, navy)

    premium_max = max(data["current_premium"], data["fifth_premium"], 1)
    # 막대 영역과 금액 영역을 분리해 긴 막대가 숫자와 겹치지 않도록 합니다.
    bar_x, bar_w, bar_h = premium_x+31*mm, 112*mm, 3.2*mm
    premium_rows = [
        ("현재 실손", data["current_premium"], blue, premium_y+12.5*mm),
        ("5세대", data["fifth_premium"], teal, premium_y+5*mm),
    ]
    for label, value, color, y in premium_rows:
        text(premium_x+5*mm, y+.5*mm, label, 9.5, muted)
        c.setFillColor(colors.HexColor("#E7EEF3"))
        c.roundRect(bar_x, y, bar_w, bar_h, bar_h/2, fill=1, stroke=0)
        c.setFillColor(color)
        c.roundRect(bar_x, y, max(bar_w*value/premium_max, 1.2*mm), bar_h, bar_h/2, fill=1, stroke=0)
        c.setFillColor(navy); c.setFont(font, 11)
        c.drawRightString(premium_x+176*mm, y+.35*mm, won(value))

    if data["premium_diff"] > 0:
        difference_title = "월 절감 예상액"
    elif data["premium_diff"] < 0:
        difference_title = "월 추가 예상액"
    else:
        difference_title = "월 보험료 차이"
    badge_x = premium_x+184*mm
    c.setFillColor(colors.HexColor("#E8F6F5") if data["premium_diff"] >= 0 else colors.HexColor("#FFF3E8"))
    c.roundRect(badge_x, premium_y+4*mm, 47*mm, 19*mm, 3.5*mm, fill=1, stroke=0)
    text(badge_x+4*mm, premium_y+16.5*mm, difference_title, 9, muted)
    text(badge_x+4*mm, premium_y+8*mm, won(abs(data["premium_diff"])), 14, teal if data["premium_diff"] >= 0 else colors.HexColor("#C66A24"))

    # rates table
    x0, y0, widths, rh = 21*mm, page_h-75*mm, [54*mm, 40*mm, 40*mm], 10*mm
    headers = ["핵심 비교", f"현재 {data['generation']}", "5세대 실손"]
    for col, width in enumerate(widths):
        x = x0 + sum(widths[:col]); c.setFillColor(colors.HexColor("#EDF3F7")); c.rect(x, y0, width, rh, fill=1, stroke=0); text(x+3*mm, y0+3.2*mm, headers[col], 9.5, muted)
    rate_rows = [("급여 입원 자기부담", "급여"), ("중증 비급여 자기부담", "중증 비급여"), ("비중증 비급여 자기부담", "비중증 비급여")]
    for row, (label, key) in enumerate(rate_rows, 1):
        y = y0-row*rh; c.setStrokeColor(line); c.line(x0, y, x0+sum(widths), y)
        prefix = "구분 없음 · " if data['generation'] in ("1세대", "2세대", "3세대") and key != "급여" else ""
        vals = [label, f"{prefix}{data['current_rates'][key]:g}%", f"{FIFTH_RATES[key]:g}%"]
        for col, value in enumerate(vals): text(x0+sum(widths[:col])+3*mm, y+3.2*mm, value, 9.3, blue if col == 1 else teal if col == 2 else navy)

    # 동일한 총 의료비 안에서 예상 보험금과 본인 부담을 나누어 표시
    chart_x, chart_y, chart_w = 166*mm, page_h-75*mm, 92*mm
    text(chart_x, chart_y+6*mm, "입원·수술 예시 결과", 13.5, navy)
    text(chart_x, chart_y-2*mm, "총 의료비", 9.5, muted)
    c.setFillColor(navy); c.setFont(font, 13)
    c.drawRightString(chart_x+chart_w, chart_y-2*mm, won(data["total_medical"]))
    total = max(data["total_medical"], 1)
    orange, gray = colors.HexColor("#DD762F"), colors.HexColor("#84909A")
    excluded = data["excluded"]
    result_rows = [
        (f"현재 {data['generation']}", data["current_payout"], data["current_burden"], blue),
        ("5세대 실손", data["fifth_payout"], data["fifth_burden"], teal),
    ]
    for idx, (label, payout, burden, payout_color) in enumerate(result_rows):
        y = chart_y-21*mm-idx*20*mm
        covered_burden = max(0, burden-excluded)
        text(chart_x, y+12.5*mm, label, 9.5, muted)
        text(chart_x, y+7.5*mm, f"보험금 {compact_manwon(payout)}", 8.5, payout_color)
        c.setFillColor(orange); c.setFont(font, 8.5)
        c.drawCentredString(chart_x+chart_w*.60, y+7.5*mm, f"본인 부담 {compact_manwon(covered_burden)}")
        c.setFillColor(gray); c.setFont(font, 8.5)
        c.drawRightString(chart_x+chart_w, y+7.5*mm, f"보상 제외 {compact_manwon(excluded)}")
        payout_w = chart_w * max(0, min(payout / total, 1))
        burden_w = chart_w * max(0, min(covered_burden / total, 1))
        excluded_w = chart_w * max(0, min(excluded / total, 1))
        c.setFillColor(colors.HexColor("#E8EEF2")); c.roundRect(chart_x, y, chart_w, 6*mm, 2*mm, fill=1, stroke=0)
        if payout_w > 0:
            c.setFillColor(payout_color); c.roundRect(chart_x, y, payout_w, 6*mm, 2*mm, fill=1, stroke=0)
        if burden_w > 0:
            # 중간 구간도 부드러운 모서리를 적용하되 미세하게 겹쳐 구간 사이의 흰 틈을 방지합니다.
            overlap = .45*mm
            c.setFillColor(orange)
            c.roundRect(
                chart_x+payout_w-overlap,
                y,
                burden_w+overlap*2,
                6*mm,
                1.15*mm,
                fill=1,
                stroke=0,
            )
        if excluded_w > 0:
            c.setFillColor(gray); c.roundRect(chart_x+chart_w-excluded_w, y, excluded_w, 6*mm, 2*mm, fill=1, stroke=0)
    c.setFillColor(colors.HexColor("#EEF5FB")); c.roundRect(chart_x, chart_y-63*mm, chart_w, 11*mm, 3*mm, fill=1, stroke=0)
    text(chart_x+4*mm, chart_y-59.5*mm, "고객 부담 차이", 10, muted)
    c.setFillColor(navy); c.setFont(font, 13)
    c.drawRightString(chart_x+chart_w-4*mm, chart_y-59.5*mm, won(abs(data['burden_diff'])))

    # 누적 보험료 비교: 1~5년을 같은 축에서 비교하는 그룹 막대 차트
    base_y = 16*mm
    chart_left, chart_bottom, chart_width, chart_height = 21*mm, base_y, 237*mm, 30*mm
    text(chart_left, chart_bottom+56*mm, "누적 보험료 비교", 13.5, navy)
    text(chart_left+43*mm, chart_bottom+56*mm, "● 현재 실손", 10, blue)
    text(chart_left+76*mm, chart_bottom+56*mm, "● 5세대", 10, teal)
    cumulative = [(years, data['current_premium']*12*years, data['fifth_premium']*12*years) for years in range(1, 6)]
    cumulative_max = max((max(current, fifth) for _, current, fifth in cumulative), default=1) or 1
    # 차분한 보고서형 차트: 옅은 패널, 보조선, 슬림한 막대로 정보 위계를 정리합니다.
    c.setFillColor(colors.HexColor("#F8FAFC"))
    c.roundRect(chart_left, chart_bottom-1.5*mm, chart_width, 48*mm, 3*mm, fill=1, stroke=0)
    baseline = chart_bottom+14*mm
    c.setStrokeColor(colors.HexColor("#E4EAF0")); c.setLineWidth(.45)
    for ratio in (.25, .5, .75, 1.0):
        grid_y = baseline+chart_height*ratio
        c.line(chart_left+3*mm, grid_y, chart_left+chart_width-3*mm, grid_y)
    c.setStrokeColor(colors.HexColor("#C9D5DF")); c.setLineWidth(.7)
    c.line(chart_left, baseline, chart_left+chart_width, baseline)
    group_gap, bar_width = 45*mm, 7*mm
    for idx, (years, current, fifth) in enumerate(cumulative):
        group_x = chart_left+8*mm+idx*group_gap
        for offset, value, color in ((0, current, blue), (9*mm, fifth, teal)):
            height = max(chart_height*value/cumulative_max, 1*mm)
            c.setFillColor(color)
            c.roundRect(group_x+offset, baseline, bar_width, height, .7*mm, fill=1, stroke=0)
            c.setFillColor(color); c.setFont(font, 8.5)
            c.drawCentredString(group_x+offset+bar_width/2, baseline+height+2*mm, compact_manwon(value))
        center_x = group_x+8*mm
        c.setFillColor(navy); c.setFont(font, 9.5)
        c.drawCentredString(center_x, chart_bottom+9.5*mm, f"{years}년")
        diff_text = f"차이 {compact_manwon(abs(current-fifth))}"
        badge_w = 28*mm
        c.setFillColor(colors.HexColor("#EAF0F5"))
        c.roundRect(center_x-badge_w/2, chart_bottom+2.2*mm, badge_w, 5.5*mm, 2.2*mm, fill=1, stroke=0)
        c.setFillColor(colors.HexColor("#365D78")); c.setFont(font, 8.2)
        c.drawCentredString(center_x, chart_bottom+4*mm, diff_text)
    c.showPage(); c.save(); output.seek(0)
    return output.getvalue()


def run() -> None:
    inject_styles()
    page_header("고객 상담", "실손보험 세대 비교 도우미", "현재 가입 실손과 5세대 실손의 보험료와 입원 보장 차이를 한눈에 비교합니다.", "🩺")

    with st.expander("✦ 사용 방법 및 비교 기준", expanded=False):
        st.markdown("✦ 현재 실손 세대와 보험료를 입력합니다.\n\n✦ 입원·수술 예시 금액을 확인하거나 수정합니다.\n\n✦ 화면 결과를 확인한 뒤 고객용 PDF를 내려받습니다.")
        st.caption("세대별 대표 자기부담률을 적용하는 상담용 간단 비교이며, 실제 계약의 약관과 공제금액이 우선합니다.")

    section_intro("INPUT", "기본 정보", "고객 정보와 비교할 실손 세대를 입력해 주세요.")
    customer_column, insurance_column = st.columns([0.4, 0.6], gap="medium")
    with customer_column:
        with st.container(border=True):
            st.markdown('<div class="sc-input-title"><span>고객·상담 정보</span></div>', unsafe_allow_html=True)
            customer = st.text_input("고객명 (선택)", placeholder="예: 홍길동", key="sc_customer")
            consultant = st.text_input("담당자 (선택)", placeholder="예: 박병선", key="sc_consultant")
            age = int(st.number_input("실제 만 나이", min_value=0, max_value=100, value=40, step=1, key="sc_age"))
            gender = st.selectbox("성별", ["남성", "여성"], key="sc_gender")

    with insurance_column:
        with st.container(border=True):
            st.markdown('<div class="sc-input-title"><span>실손 비교 정보</span></div>', unsafe_allow_html=True)
            st.markdown('<div class="sc-subtitle"><span>현재 가입 실손</span></div>', unsafe_allow_html=True)
            generation = st.selectbox("현재 실손 세대", ["1세대", "2세대", "3세대", "4세대"], index=1, key="sc_generation")
            if generation == "1세대":
                option = st.selectbox("현재 계약 자기부담률", ["0%", "10%", "20%"], help="1세대는 계약별 차이가 커 실제 증권에 맞게 선택해 주세요.", key="sc_contract_option_1")
            elif generation == "2세대":
                option = st.selectbox("현재 계약 유형", ["10%형", "20%형"], key="sc_contract_option_2")
            elif generation == "3세대":
                option = st.selectbox("급여 자기부담 유형", ["급여 10%형", "급여 20%형"], key="sc_contract_option_3")
            else:
                option = "4세대 대표 기준"
                st.text_input("현재 계약 유형", value=option, disabled=True, key="sc_contract_option_4")

            estimated_current = estimate_current_premium(age, gender, generation, option)
            current_signature = (age, gender, generation, option)
            if st.session_state.get("sc_current_signature") != current_signature:
                st.session_state["sc_current_signature"] = current_signature
                st.session_state["sc_current_premium"] = estimated_current
                st.session_state["sc_current_modified"] = False
            current_premium = float(
                st.number_input(
                    "현재 실손 월 보험료",
                    min_value=0,
                    step=100,
                    format="%d",
                    key="sc_current_premium",
                    on_change=_mark_current_premium_modified,
                    help="처음에는 세대·연령·성별 참고값이 표시되며 실제 납부 보험료를 알고 있다면 수정할 수 있습니다.",
                )
            )
            current_modified = bool(st.session_state.get("sc_current_modified", False))
            current_premium_basis = "실제 납부 보험료" if current_modified else f"{generation} · 만 {age}세 {gender} 예상값"
            current_basis_title = "실제 납부 보험료" if current_modified else "세대·연령 기준 예상금액"
            st.markdown(
                f'<div class="sc-basis"><b>{current_basis_title}</b> · {current_premium_basis}<br>'
                '보험회사와 갱신 이력에 따라 달라질 수 있으며 금액을 수정하면 실제 납부 보험료로 반영됩니다.</div>',
                unsafe_allow_html=True,
            )

            claim_level = st.selectbox("최근 1년 실손보험금 수령 수준", CLAIM_LEVELS, key="sc_claim_level")
            if claim_level == "직접 입력":
                st.number_input("최근 1년 수령 보험금", min_value=0, step=100_000, format="%d", key="sc_claim_exact")

            st.markdown('<div class="sc-divider"></div>', unsafe_allow_html=True)
            st.markdown('<div class="sc-subtitle"><span>5세대 비교 실손</span><span class="sc-fixed-badge">전체 보장형 고정</span></div>', unsafe_allow_html=True)
            premium_mode = st.radio("5세대 보험료 입력 방식", PREMIUM_MODES, index=0, horizontal=True, key="sc_premium_mode")

            if premium_mode == "연령별 예상 보험료":
                estimated_premium = estimate_fifth_premium(age, gender)
                reference_signature = (age, gender)
                if st.session_state.get("sc_reference_signature") != reference_signature:
                    st.session_state["sc_reference_signature"] = reference_signature
                    st.session_state["sc_reference_premium"] = estimated_premium
                    st.session_state["sc_reference_modified"] = False
                fifth_premium = float(
                    st.number_input(
                        "비교에 적용할 5세대 월 보험료",
                        min_value=0,
                        step=100,
                        format="%d",
                        key="sc_reference_premium",
                        on_change=_mark_reference_premium_modified,
                    )
                )
                reference_modified = bool(st.session_state.get("sc_reference_modified", False))
                premium_basis = "연령 기준 예상값을 사용자가 수정한 금액 · 전체 보장형" if reference_modified else f"만 {age}세 {gender} · 전체 보장형 예상값"
                st.markdown(
                    f'<div class="sc-basis"><b>공개 보험료 예시 기반 상담용 추정값</b> · 만 {age}세 {gender} · 전체 보장형<br>'
                    f'실제 보험료는 보험회사, 직업, 가입조건에 따라 달라질 수 있으며 현재 표시된 금액을 직접 수정할 수 있습니다.</div>',
                    unsafe_allow_html=True,
                )
            else:
                fifth_premium = float(
                    st.number_input("가입제안서의 5세대 월 보험료", min_value=0, value=30_000, step=100, format="%d", key="sc_direct_premium")
                )
                premium_basis = "가입제안서 직접 입력 금액 · 전체 보장형"

    rates = current_rates(generation, option)
    section_intro("COMPARE", "한눈에 보는 핵심 차이", "선택한 현재 실손의 대표 기준과 5세대 기준을 비교합니다.")
    rate_table(generation, rates)
    if generation == "1세대":
        st.caption("※ 1세대는 표준화 이전 상품으로 계약별 차이가 큽니다. 선택한 비율이 실제 증권과 일치하는지 확인해 주세요.")
    elif generation == "3세대":
        st.caption("※ 3세대의 도수치료·비급여 주사·비급여 MRI 등 3대 비급여 특약은 30%가 적용될 수 있습니다.")

    comparison_bars("월 보험료 비교", current_premium, fifth_premium, f"월 보험료 차이 {won(abs(current_premium-fifth_premium))}")

    section_intro("CASE", "입원·수술 사례 비교", "금액은 만원 단위로 입력하며 총 의료비는 자동 계산됩니다.")
    example_button_column, example_note_column = st.columns([0.34, 0.66], gap="small")
    with example_button_column:
        if st.button("일반 예시 금액 입력", key="sc_example", use_container_width=True):
            st.session_state.update(sc_salary_manwon=150, sc_severe_manwon=0, sc_nonsevere_manwon=50, sc_excluded_manwon=30)
            st.rerun()
    with example_note_column:
        st.markdown('<div class="sc-example-note">일반 예시 · 총 의료비 <b>230만원</b> 기준</div>', unsafe_allow_html=True)
    m1, m2, m3, m4 = st.columns(4)
    salary = float(m1.number_input("급여 의료비 (만원)", min_value=0, value=150, step=1, format="%d", key="sc_salary_manwon")) * 10_000
    severe = float(m2.number_input("중증 비급여 (만원)", min_value=0, value=0, step=1, format="%d", key="sc_severe_manwon")) * 10_000
    nonsevere = float(m3.number_input("비중증 비급여 (만원)", min_value=0, value=50, step=1, format="%d", key="sc_nonsevere_manwon")) * 10_000
    excluded = float(m4.number_input("보상 제외 (만원)", min_value=0, value=30, step=1, format="%d", key="sc_excluded_manwon")) * 10_000
    medical = {"급여": salary, "중증 비급여": severe, "비중증 비급여": nonsevere, "보상 제외 가능 비급여": excluded}
    total_medical = sum(medical.values())
    current_payout, current_burden = calculate(medical, rates)
    fifth_payout, fifth_burden = calculate(medical, FIFTH_RATES)

    k1, k2, k3 = st.columns(3)
    k1.metric("총 의료비", won(total_medical))
    k2.metric(f"현재 {generation} 예상 보험금", won(current_payout))
    k3.metric("5세대 예상 보험금", won(fifth_payout))
    medical_stacked_bars(total_medical, generation, current_payout, current_burden, fifth_payout, fifth_burden, excluded)

    section_intro("RESULT", "누적 보험료", "현재 월 보험료가 동일하게 유지된다는 단순 가정입니다.")
    cols = st.columns(5)
    for col, years in zip(cols, range(1, 6)):
        with col:
            st.markdown(
                f'<div class="sc-cumulative-card"><div class="sc-cumulative-year">{years}년 누적</div>'
                f'<div class="sc-cumulative-row sc-cumulative-current"><span>현재 실손</span><strong>{won(current_premium*12*years)}</strong></div>'
                f'<div class="sc-cumulative-row sc-cumulative-fifth"><span>5세대 실손</span><strong>{won(fifth_premium*12*years)}</strong></div>'
                f'<div class="sc-cumulative-diff"><span>차이</span><strong>{won(abs((current_premium-fifth_premium)*12*years))}</strong></div></div>',
                unsafe_allow_html=True,
            )
    cumulative_premium_chart(current_premium, fifth_premium)

    saving = current_premium - fifth_premium
    burden_gap = fifth_burden - current_burden
    if saving > 0 and saving * 60 > max(burden_gap, 0):
        result_text = "보험료 절감 효과가 큼"
    elif burden_gap > saving * 60 and burden_gap > 0:
        result_text = "현재 실손 유지의 보장 가치가 큼"
    else:
        result_text = "보험료와 보장 차이를 함께 비교할 필요가 있음"
    st.info(f"비교 요약 · **{result_text}**")

    data = {
        "customer": customer.strip() or "OOO", "consultant": consultant.strip(), "generation": generation,
        "current_rates": rates, "current_premium": current_premium, "fifth_premium": fifth_premium,
        "premium_diff": current_premium-fifth_premium, "current_payout": current_payout, "fifth_payout": fifth_payout,
        "current_burden": current_burden, "fifth_burden": fifth_burden, "burden_diff": fifth_burden-current_burden,
        "total_medical": total_medical, "excluded": excluded, "premium_basis": premium_basis, "current_premium_basis": current_premium_basis,
    }
    pdf = build_pdf(data)
    filename = f"{safe_filename(data['customer'])}님_실손보험_세대비교_{date.today():%Y%m%d}.pdf"
    st.download_button("고객용 비교안 PDF 다운로드", pdf, filename, "application/pdf", type="primary", use_container_width=True)
    st.caption("본 자료는 간단 비교용이며 실제 보험금은 가입 상품의 약관, 공제금액, 보상한도 및 보험회사 심사에 따라 달라질 수 있습니다.")
