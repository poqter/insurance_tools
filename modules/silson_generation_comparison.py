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


def won(value: float) -> str:
    return f"{int(round(value)):,}원"


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
        .sc-note{margin-top:.65rem;color:#718393;font-size:.74rem;line-height:1.55}
        @media(max-width:700px){.sc-bar-row{grid-template-columns:6.5rem 1fr}.sc-value{grid-column:2}.sc-rate-grid>div{padding:.65rem .4rem;font-size:.76rem}}
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

    navy, blue, teal, muted, line = colors.HexColor("#16324F"), colors.HexColor("#1769DC"), colors.HexColor("#119B98"), colors.HexColor("#687F91"), colors.HexColor("#DCE6EE")

    def text(x, y, value, size=9, color=navy, bold=False):
        c.setFillColor(color); c.setFont(font, size); c.drawString(x, y, str(value))

    c.setFillColor(colors.HexColor("#F6F9FC")); c.rect(0, 0, page_w, page_h, fill=1, stroke=0)
    c.setFillColor(colors.white); c.roundRect(13*mm, 11*mm, page_w-26*mm, page_h-22*mm, 5*mm, fill=1, stroke=0)
    text(21*mm, page_h-25*mm, f"{data['customer']}님 실손보험 세대 비교 안내", 18, navy)
    text(21*mm, page_h-33*mm, f"현재 {data['generation']} 실손과 5세대 실손의 보험료·입원 보장을 간단히 비교했습니다.", 8.5, muted)
    if data["consultant"]:
        c.drawRightString(page_w-21*mm, page_h-25*mm, f"담당자  {data['consultant']}")

    # 월 보험료 비교: 막대 길이와 정확한 금액을 함께 보여주는 메인 차트
    premium_x, premium_y, premium_w, premium_h = 21*mm, page_h-60*mm, 237*mm, 24*mm
    c.setFillColor(colors.HexColor("#F7FAFD"))
    c.roundRect(premium_x, premium_y, premium_w, premium_h, 4*mm, fill=1, stroke=0)
    text(premium_x+5*mm, premium_y+17*mm, "월 보험료 비교", 9.5, navy)

    premium_max = max(data["current_premium"], data["fifth_premium"], 1)
    bar_x, bar_w, bar_h = premium_x+31*mm, 128*mm, 3.7*mm
    premium_rows = [
        ("현재 실손", data["current_premium"], blue, premium_y+11.5*mm),
        ("5세대", data["fifth_premium"], teal, premium_y+5*mm),
    ]
    for label, value, color, y in premium_rows:
        text(premium_x+5*mm, y+.5*mm, label, 7.2, muted)
        c.setFillColor(colors.HexColor("#E7EEF3"))
        c.roundRect(bar_x, y, bar_w, bar_h, bar_h/2, fill=1, stroke=0)
        c.setFillColor(color)
        c.roundRect(bar_x, y, max(bar_w*value/premium_max, 1.2*mm), bar_h, bar_h/2, fill=1, stroke=0)
        c.setFillColor(navy); c.setFont(font, 8)
        c.drawRightString(premium_x+174*mm, y+.5*mm, won(value))

    if data["premium_diff"] > 0:
        difference_title = "월 절감 예상액"
    elif data["premium_diff"] < 0:
        difference_title = "월 추가 예상액"
    else:
        difference_title = "월 보험료 차이"
    badge_x = premium_x+184*mm
    c.setFillColor(colors.HexColor("#E8F6F5") if data["premium_diff"] >= 0 else colors.HexColor("#FFF3E8"))
    c.roundRect(badge_x, premium_y+4*mm, 47*mm, 16*mm, 3.5*mm, fill=1, stroke=0)
    text(badge_x+4*mm, premium_y+14*mm, difference_title, 7, muted)
    text(badge_x+4*mm, premium_y+7*mm, won(abs(data["premium_diff"])), 11, teal if data["premium_diff"] >= 0 else colors.HexColor("#C66A24"))

    # rates table
    x0, y0, widths, rh = 21*mm, page_h-72*mm, [54*mm, 40*mm, 40*mm], 8*mm
    headers = ["핵심 비교", f"현재 {data['generation']}", "5세대 실손"]
    for col, width in enumerate(widths):
        x = x0 + sum(widths[:col]); c.setFillColor(colors.HexColor("#EDF3F7")); c.rect(x, y0, width, rh, fill=1, stroke=0); text(x+3*mm, y0+2.7*mm, headers[col], 7.5, muted)
    rate_rows = [("급여 입원 자기부담", "급여"), ("중증 비급여 자기부담", "중증 비급여"), ("비중증 비급여 자기부담", "비중증 비급여")]
    for row, (label, key) in enumerate(rate_rows, 1):
        y = y0-row*rh; c.setStrokeColor(line); c.line(x0, y, x0+sum(widths), y)
        prefix = "구분 없음 · " if data['generation'] in ("1세대", "2세대", "3세대") and key != "급여" else ""
        vals = [label, f"{prefix}{data['current_rates'][key]:g}%", f"{FIFTH_RATES[key]:g}%"]
        for col, value in enumerate(vals): text(x0+sum(widths[:col])+3*mm, y+2.7*mm, value, 7.5, blue if col == 1 else teal if col == 2 else navy)

    # result bars
    chart_x, chart_y, chart_w = 166*mm, page_h-72*mm, 92*mm
    text(chart_x, chart_y+5*mm, "입원·수술 예시 결과", 11, navy)
    max_burden = max(data["current_burden"], data["fifth_burden"], 1)
    for idx, (label, val, color) in enumerate([("현재 고객 부담", data["current_burden"], blue), ("5세대 고객 부담", data["fifth_burden"], teal)]):
        y = chart_y-8*mm-idx*14*mm; text(chart_x, y+4*mm, label, 7.5, muted)
        c.setFillColor(colors.HexColor("#EAF0F4")); c.roundRect(chart_x, y, chart_w, 3.8*mm, 1.9*mm, fill=1, stroke=0)
        c.setFillColor(color); c.roundRect(chart_x, y, chart_w*val/max_burden, 3.8*mm, 1.9*mm, fill=1, stroke=0)
        c.drawRightString(chart_x+chart_w, y+5*mm, won(val))
    text(chart_x, chart_y-39*mm, f"고객 부담 차이  {won(abs(data['burden_diff']))}", 10, navy)

    # 누적 보험료 비교: 1·5·10년을 같은 축에서 비교하는 그룹 막대 차트
    base_y = 25*mm
    chart_left, chart_bottom, chart_width, chart_height = 21*mm, base_y, 132*mm, 40*mm
    text(chart_left, chart_bottom+54*mm, "누적 보험료 비교", 11, navy)
    text(chart_left+38*mm, chart_bottom+54*mm, "● 현재 실손", 6.8, blue)
    text(chart_left+62*mm, chart_bottom+54*mm, "● 5세대", 6.8, teal)
    cumulative = [(years, data['current_premium']*12*years, data['fifth_premium']*12*years) for years in (1, 5, 10)]
    cumulative_max = max((max(current, fifth) for _, current, fifth in cumulative), default=1) or 1
    c.setStrokeColor(colors.HexColor("#E3EBF1")); c.setLineWidth(.5)
    c.line(chart_left, chart_bottom+5*mm, chart_left+chart_width, chart_bottom+5*mm)
    group_gap, bar_width = 43*mm, 8*mm
    for idx, (years, current, fifth) in enumerate(cumulative):
        group_x = chart_left+12*mm+idx*group_gap
        for offset, value, color in ((0, current, blue), (9.5*mm, fifth, teal)):
            height = max(chart_height*value/cumulative_max, 1*mm)
            c.setFillColor(color)
            c.roundRect(group_x+offset, chart_bottom+5*mm, bar_width, height, 1.5*mm, fill=1, stroke=0)
        c.setFillColor(navy); c.setFont(font, 6.5)
        c.drawCentredString(group_x+8.75*mm, chart_bottom+1.2*mm, f"{years}년")
        c.setFillColor(muted); c.setFont(font, 5.8)
        c.drawCentredString(group_x+4*mm, chart_bottom+6*mm+max(chart_height*current/cumulative_max, 1*mm), won(current))
        c.drawCentredString(group_x+13.5*mm, chart_bottom+6*mm+max(chart_height*fifth/cumulative_max, 1*mm), won(fifth))
    text(166*mm, base_y+54*mm, "안내", 10, navy)
    notes = ["본 자료는 입력값과 대표 자기부담률을 이용한 간단 비교입니다.", "실제 지급액은 약관, 공제금액, 보상한도와 심사 결과에 따라 달라질 수 있습니다.", f"기준일 {STANDARD_DATE} · 보험료는 변동 없이 유지된다고 가정했습니다."]
    for i, note in enumerate(notes): text(166*mm, base_y+44*mm-i*7*mm, f"- {note}", 7, muted)
    c.showPage(); c.save(); output.seek(0)
    return output.getvalue()


def run() -> None:
    inject_styles()
    page_header("고객 상담", "실손보험 세대 비교 도우미", "현재 가입 실손과 5세대 실손의 보험료와 입원 보장 차이를 한눈에 비교합니다.", "🩺")

    with st.expander("✦ 사용 방법 및 비교 기준", expanded=False):
        st.markdown("1. 현재 실손 세대와 보험료를 입력합니다.\n2. 입원·수술 예시 금액을 확인하거나 수정합니다.\n3. 화면 결과를 확인한 뒤 고객용 PDF를 내려받습니다.")
        st.caption("세대별 대표 자기부담률을 적용하는 상담용 간단 비교이며, 실제 계약의 약관과 공제금액이 우선합니다.")

    section_intro("INPUT", "기본 정보", "고객 정보와 비교할 실손 세대를 입력해 주세요.")
    with st.container(border=True):
        c1, c2, c3 = st.columns(3)
        customer = c1.text_input("고객명 (선택)", placeholder="예: 홍길동", key="sc_customer")
        consultant = c2.text_input("담당자 (선택)", placeholder="예: 박병선", key="sc_consultant")
        generation = c3.selectbox("현재 실손 세대", ["1세대", "2세대", "3세대", "4세대"], index=1, key="sc_generation")

        o1, o2 = st.columns(2)
        if generation == "1세대":
            option = o1.selectbox("현재 계약 자기부담률", ["0%", "10%", "20%"], help="1세대는 계약별 차이가 커 실제 증권에 맞게 선택해 주세요.")
        elif generation == "2세대":
            option = o1.selectbox("현재 계약 유형", ["10%형", "20%형"])
        elif generation == "3세대":
            option = o1.selectbox("급여 자기부담 유형", ["급여 10%형", "급여 20%형"])
        else:
            option = "4세대 대표 기준"
            o1.text_input("현재 계약 유형", value=option, disabled=True)
        claim_level = o2.selectbox("최근 1년 실손보험금 수령 수준", CLAIM_LEVELS)
        if claim_level == "직접 입력":
            st.number_input("최근 1년 수령 보험금", min_value=0, step=100_000, format="%d", key="sc_claim_exact")

        p1, p2 = st.columns(2)
        current_premium = float(p1.number_input("현재 월 보험료", min_value=0, value=60_000, step=1_000, format="%d"))
        fifth_premium = float(p2.number_input("5세대 제안 월 보험료", min_value=0, value=30_000, step=1_000, format="%d", help="실제 가입제안서의 보험료를 입력해 주세요."))

    rates = current_rates(generation, option)
    section_intro("COMPARE", "한눈에 보는 핵심 차이", "선택한 현재 실손의 대표 기준과 5세대 기준을 비교합니다.")
    rate_table(generation, rates)
    if generation == "1세대":
        st.caption("※ 1세대는 표준화 이전 상품으로 계약별 차이가 큽니다. 선택한 비율이 실제 증권과 일치하는지 확인해 주세요.")
    elif generation == "3세대":
        st.caption("※ 3세대의 도수치료·비급여 주사·비급여 MRI 등 3대 비급여 특약은 30%가 적용될 수 있습니다.")

    comparison_bars("월 보험료 비교", current_premium, fifth_premium, f"월 보험료 차이 {won(abs(current_premium-fifth_premium))}")

    section_intro("CASE", "입원·수술 사례 비교", "총 의료비는 아래 항목의 합계로 자동 계산됩니다.")
    if st.button("일반 예시 금액 입력", key="sc_example"):
        st.session_state.update(sc_salary=2_000_000, sc_severe=2_000_000, sc_nonsevere=500_000, sc_excluded=0)
        st.rerun()
    m1, m2, m3, m4 = st.columns(4)
    salary = float(m1.number_input("급여 본인부담금", min_value=0, value=2_000_000, step=100_000, format="%d", key="sc_salary"))
    severe = float(m2.number_input("중증 비급여", min_value=0, value=2_000_000, step=100_000, format="%d", key="sc_severe"))
    nonsevere = float(m3.number_input("비중증 비급여", min_value=0, value=500_000, step=100_000, format="%d", key="sc_nonsevere"))
    excluded = float(m4.number_input("보상 제외 가능 비급여", min_value=0, value=0, step=100_000, format="%d", key="sc_excluded"))
    medical = {"급여": salary, "중증 비급여": severe, "비중증 비급여": nonsevere, "보상 제외 가능 비급여": excluded}
    total_medical = sum(medical.values())
    current_payout, current_burden = calculate(medical, rates)
    fifth_payout, fifth_burden = calculate(medical, FIFTH_RATES)

    k1, k2, k3 = st.columns(3)
    k1.metric("총 의료비", won(total_medical))
    k2.metric(f"현재 {generation} 예상 보험금", won(current_payout))
    k3.metric("5세대 예상 보험금", won(fifth_payout))
    comparison_bars("고객 부담 비교", current_burden, fifth_burden, f"고객 부담 차이 {won(abs(current_burden-fifth_burden))}")

    section_intro("RESULT", "누적 보험료", "현재 월 보험료가 동일하게 유지된다는 단순 가정입니다.")
    cols = st.columns(3)
    for col, years in zip(cols, (1, 5, 10)):
        with col:
            st.markdown(f"**{years}년 누적**")
            st.caption(f"현재 {won(current_premium*12*years)} · 5세대 {won(fifth_premium*12*years)}")

    saving = current_premium - fifth_premium
    burden_gap = fifth_burden - current_burden
    if saving > 0 and saving * 120 > max(burden_gap, 0):
        result_text = "보험료 절감 효과가 큼"
    elif burden_gap > saving * 120 and burden_gap > 0:
        result_text = "현재 실손 유지의 보장 가치가 큼"
    else:
        result_text = "보험료와 보장 차이를 함께 비교할 필요가 있음"
    st.info(f"비교 요약 · **{result_text}**")

    data = {
        "customer": customer.strip() or "OOO", "consultant": consultant.strip(), "generation": generation,
        "current_rates": rates, "current_premium": current_premium, "fifth_premium": fifth_premium,
        "premium_diff": current_premium-fifth_premium, "current_payout": current_payout, "fifth_payout": fifth_payout,
        "current_burden": current_burden, "fifth_burden": fifth_burden, "burden_diff": fifth_burden-current_burden,
    }
    pdf = build_pdf(data)
    filename = f"{safe_filename(data['customer'])}님_실손보험_세대비교_{date.today():%Y%m%d}.pdf"
    st.download_button("고객용 비교안 PDF 다운로드", pdf, filename, "application/pdf", type="primary", use_container_width=True)
    st.caption("본 자료는 간단 비교용이며 실제 보험금은 가입 상품의 약관, 공제금액, 보상한도 및 보험회사 심사에 따라 달라질 수 있습니다.")
