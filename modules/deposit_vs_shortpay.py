import time
from datetime import datetime
from io import BytesIO

import streamlit as st

from .ui_components import page_header


TAX_RATE = 0.154
DEPOSIT_REPEAT_YEARS = 10


def format_currency(value_manwon: float) -> str:
    """만원 단위 값을 자연스러운 원화 문자열로 변환합니다."""
    won = int(round(value_manwon * 10_000))
    sign = "-" if won < 0 else ""
    won = abs(won)

    if won >= 100_000_000 and won % 1_000_000 == 0:
        eok = won / 100_000_000
        text = f"{eok:,.2f}".rstrip("0").rstrip(".")
        return f"{sign}{text}억원"
    if won % 10_000 == 0:
        return f"{sign}{won // 10_000:,}만원"
    return f"{sign}{won:,}원"


def calculate_deposit(monthly_manwon: float, annual_rate: float) -> dict:
    """기존 방식대로 1년 적금의 단리 세후이자를 10회 합산합니다."""
    monthly_rate = annual_rate / 100 / 12
    interest_weight = sum(12 - month for month in range(12))  # 12+...+1 = 78
    one_year_principal = monthly_manwon * 12
    pretax_interest = monthly_manwon * monthly_rate * interest_weight
    tax = pretax_interest * TAX_RATE
    aftertax_interest = pretax_interest - tax
    ten_year_interest = aftertax_interest * DEPOSIT_REPEAT_YEARS

    return {
        "one_year_principal": one_year_principal,
        "pretax_interest": pretax_interest,
        "tax": tax,
        "aftertax_interest": aftertax_interest,
        "ten_year_interest": ten_year_interest,
        "ten_year_total_paid": monthly_manwon * 12 * DEPOSIT_REPEAT_YEARS,
        "interest_weight": interest_weight,
    }


def calculate_shortpay(
    monthly_manwon: float,
    pay_years: int,
    refund_rate: float,
) -> dict:
    total_premium = monthly_manwon * 12 * pay_years
    refund_amount = total_premium * refund_rate / 100
    refund_gain = refund_amount - total_premium

    return {
        "total_premium": total_premium,
        "refund_amount": refund_amount,
        "refund_gain": refund_gain,
    }


def calculate_required_deposit_rate(
    monthly_manwon: float,
    target_gain_manwon: float,
    interest_weight: int = 78,
) -> float:
    """기존 단리 방식에서 목표 이익에 도달하기 위한 적금 연이율입니다."""
    if monthly_manwon <= 0 or target_gain_manwon <= 0:
        return 0.0
    monthly_rate = (
        (target_gain_manwon / DEPOSIT_REPEAT_YEARS)
        / (monthly_manwon * interest_weight * (1 - TAX_RATE))
    )
    return monthly_rate * 12 * 100


def calculate_required_monthly_payment(
    annual_rate: float,
    target_gain_manwon: float,
    interest_weight: int = 78,
) -> float:
    """현재 금리에서 목표 이익에 도달하기 위한 적금 월납입액입니다."""
    monthly_rate = annual_rate / 100 / 12
    denominator = monthly_rate * interest_weight * (1 - TAX_RATE) * DEPOSIT_REPEAT_YEARS
    if denominator <= 0:
        return 0.0
    return target_gain_manwon / denominator


def create_result_pdf(
    monthly: float,
    annual_rate: float,
    pay_years: int,
    refund_rate: float,
    deposit: dict,
    shortpay: dict,
    advantage: float,
    required_rate: float,
) -> bytes:
    """현재 상담 결과를 화랑 WORKSPACE 스타일의 A4 한 장 PDF로 생성합니다."""
    from reportlab.lib.colors import HexColor
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfgen import canvas

    font_name = "HYSMyeongJo-Medium"
    if font_name not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(UnicodeCIDFont(font_name))

    navy = HexColor("#16324F")
    blue = HexColor("#2F6FA3")
    blue_soft = HexColor("#EAF2F8")
    gold = HexColor("#C9963D")
    gold_deep = HexColor("#A87422")
    gold_soft = HexColor("#FBF5E9")
    text = HexColor("#203247")
    muted = HexColor("#6E7E90")
    line = HexColor("#DCE4EC")
    white = HexColor("#FFFFFF")

    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4, pageCompression=1)
    width, height = A4
    margin = 34

    def set_font(size: float) -> None:
        pdf.setFont(font_name, size)

    def rounded_box(x, y, w, h, fill, stroke=line, radius=9):
        pdf.setFillColor(fill)
        pdf.setStrokeColor(stroke)
        pdf.setLineWidth(0.7)
        pdf.roundRect(x, y, w, h, radius, fill=1, stroke=1)

    def draw_right(value: str, x: float, y: float, size: float = 8.4, color=text):
        set_font(size)
        pdf.setFillColor(color)
        pdf.drawRightString(x, y, value)

    def draw_wrapped(value: str, x: float, y: float, max_width: float, size: float, leading: float, color=muted):
        set_font(size)
        pdf.setFillColor(color)
        words = value.split()
        lines = []
        current = ""
        for word in words:
            candidate = word if not current else f"{current} {word}"
            if pdfmetrics.stringWidth(candidate, font_name, size) <= max_width:
                current = candidate
            else:
                if current:
                    lines.append(current)
                current = word
        if current:
            lines.append(current)
        for index, line_text in enumerate(lines):
            pdf.drawString(x, y - index * leading, line_text)
        return y - len(lines) * leading

    # 상단 브랜드와 제목
    set_font(8.5)
    pdf.setFillColor(gold_deep)
    pdf.drawString(margin, height - 31, "화랑 WORKSPACE")
    set_font(17)
    pdf.setFillColor(navy)
    pdf.drawString(margin, height - 53, "적금 vs 단기납 10년 예상 이익 비교")
    set_font(7.5)
    pdf.setFillColor(muted)
    pdf.drawRightString(width - margin, height - 50, datetime.now().strftime("%Y.%m.%d"))
    pdf.setStrokeColor(line)
    pdf.line(margin, height - 64, width - margin, height - 64)

    # 핵심 결론
    hero_y, hero_h = height - 136, 57
    rounded_box(margin, hero_y, width - margin * 2, hero_h, gold_soft, HexColor("#E8D5AE"), 11)
    set_font(8.2)
    pdf.setFillColor(muted)
    pdf.drawCentredString(width / 2, hero_y + 39, f"같은 월 {format_currency(monthly)}을 활용했을 때")
    set_font(14.2)
    pdf.setFillColor(gold_deep if advantage >= 0 else navy)
    if advantage >= 0:
        hero_text = f"단기납 예상 환급차익이 {format_currency(advantage)} 더 큽니다"
    else:
        hero_text = f"현재 조건에서는 적금 세후이자가 {format_currency(abs(advantage))} 더 큽니다"
    pdf.drawCentredString(width / 2, hero_y + 19, hero_text)
    set_font(6.9)
    pdf.setFillColor(muted)
    pdf.drawCentredString(
        width / 2,
        hero_y + 7,
        f"적금 1년 만기 10회 반복 · 단기납 {pay_years}년납 후 10년 시점",
    )

    # 예상 이익 비교 막대그래프
    chart_bottom = 463
    chart_top = 672
    chart_max = max(deposit["ten_year_interest"], shortpay["refund_gain"], 1)
    available_bar_height = 142
    deposit_h = max(31, deposit["ten_year_interest"] / chart_max * available_bar_height)
    shortpay_h = max(31, shortpay["refund_gain"] / chart_max * available_bar_height)
    bar_w = 105
    deposit_x = 121
    shortpay_x = width - 121 - bar_w

    pdf.setStrokeColor(line)
    pdf.setLineWidth(0.45)
    for offset in (0, 47, 94, 141):
        pdf.line(margin + 25, chart_bottom + offset, width - margin - 25, chart_bottom + offset)

    def draw_bar(x, bar_h, fill, amount, inside_label, name, detail):
        pdf.setFillColor(fill)
        pdf.setStrokeColor(fill)
        pdf.roundRect(x, chart_bottom, bar_w, bar_h, 5, fill=1, stroke=0)
        set_font(8.5)
        pdf.setFillColor(navy)
        pdf.drawCentredString(x + bar_w / 2, chart_bottom + bar_h + 10, amount)
        set_font(7.3)
        pdf.setFillColor(white)
        label_lines = inside_label.split("|")
        label_y = chart_bottom + max(10, bar_h / 2 + 2)
        for index, label in enumerate(label_lines):
            pdf.drawCentredString(x + bar_w / 2, label_y - index * 9, label)
        set_font(9.2)
        pdf.setFillColor(navy)
        pdf.drawCentredString(x + bar_w / 2, chart_bottom - 14, name)
        set_font(6.4)
        pdf.setFillColor(muted)
        pdf.drawCentredString(x + bar_w / 2, chart_bottom - 25, detail)

    draw_bar(
        deposit_x,
        deposit_h,
        blue,
        format_currency(deposit["ten_year_interest"]),
        "10년 누적|세후이자",
        "적금",
        f"월 {format_currency(monthly)} · 1년 적금 10회",
    )
    draw_bar(
        shortpay_x,
        shortpay_h,
        gold_deep,
        format_currency(shortpay["refund_gain"]),
        "10년 시점|예상 환급차익",
        "단기납",
        f"월 {format_currency(monthly)} · {pay_years}년납 후 유지",
    )

    if advantage >= 0:
        badge_text = f"적금 대비 +{format_currency(advantage)}"
        badge_w = pdfmetrics.stringWidth(badge_text, font_name, 7.2) + 18
        badge_x = min(width - margin - badge_w, shortpay_x + bar_w + 7)
        badge_y = chart_bottom + shortpay_h - 5
        rounded_box(badge_x, badge_y, badge_w, 19, gold_soft, gold, 9)
        set_font(7.2)
        pdf.setFillColor(gold_deep)
        pdf.drawCentredString(badge_x + badge_w / 2, badge_y + 6, badge_text)

    # 단기납 타임라인
    timeline_y = 409
    start_x = margin + 31
    end_x = width - margin - 31
    segment_w = (end_x - start_x) / 3
    pdf.setStrokeColor(line)
    pdf.setLineWidth(1.1)
    pdf.line(start_x, timeline_y + 26, end_x, timeline_y + 26)
    phases = [
        (f"1~{pay_years}년 납입", "보험료 납입"),
        (f"{pay_years + 1}~9년 유지", "추가납입 없이 유지"),
        ("10년 주요 시점", "환급률·비과세 요건 확인"),
        ("해지 또는 계속 유지", "환급금 추가 증가 가능"),
    ]
    for index, (main, sub) in enumerate(phases):
        x = start_x + segment_w * index
        pdf.setFillColor(gold if index == 2 else line)
        pdf.circle(x, timeline_y + 26, 4.2, fill=1, stroke=0)
        set_font(7.2)
        pdf.setFillColor(navy)
        pdf.drawCentredString(x, timeline_y + 11, main)
        set_font(5.9)
        pdf.setFillColor(muted)
        pdf.drawCentredString(x, timeline_y + 1, sub)

    # 계산 내역 카드
    card_y, card_h = 210, 174
    gap = 14
    card_w = (width - margin * 2 - gap) / 2
    left_x = margin
    right_x = margin + card_w + gap

    def draw_calc_card(x, title_text, rows, accent):
        rounded_box(x, card_y, card_w, card_h, white, line, 9)
        pdf.setFillColor(accent)
        pdf.roundRect(x, card_y + card_h - 29, card_w, 29, 9, fill=1, stroke=0)
        pdf.rect(x, card_y + card_h - 29, card_w, 9, fill=1, stroke=0)
        set_font(9.1)
        pdf.setFillColor(white)
        pdf.drawString(x + 13, card_y + card_h - 19, title_text)
        row_y = card_y + card_h - 48
        for index, (label, value) in enumerate(rows):
            set_font(7.6)
            pdf.setFillColor(muted)
            pdf.drawString(x + 13, row_y, label)
            draw_right(value, x + card_w - 13, row_y, 7.8, navy if index == len(rows) - 1 else text)
            if index < len(rows) - 1:
                pdf.setStrokeColor(line)
                pdf.setLineWidth(0.35)
                pdf.line(x + 13, row_y - 7, x + card_w - 13, row_y - 7)
            row_y -= 24

    deposit_rows = [
        ("1년 납입원금", format_currency(deposit["one_year_principal"])),
        ("1년 세전이자", format_currency(deposit["pretax_interest"])),
        ("이자소득세 15.4%", format_currency(deposit["tax"])),
        ("1년 세후이자", format_currency(deposit["aftertax_interest"])),
        ("10년 누적 세후이자", format_currency(deposit["ten_year_interest"])),
    ]
    shortpay_rows = [
        ("납입기간", f"{pay_years}년"),
        ("총납입보험료", format_currency(shortpay["total_premium"])),
        ("10년 예상 환급률", f"{refund_rate:,.1f}%"),
        ("10년 예상 해지환급금", format_currency(shortpay["refund_amount"])),
        ("예상 환급차익", format_currency(shortpay["refund_gain"])),
    ]
    draw_calc_card(left_x, "적금 계산 내역", deposit_rows, blue)
    draw_calc_card(right_x, "단기납 계산 내역", shortpay_rows, gold_deep)

    # 필요 적금금리
    rate_y, rate_h = 145, 47
    rounded_box(margin, rate_y, width - margin * 2, rate_h, blue_soft, HexColor("#C9DBE9"), 9)
    set_font(8.1)
    pdf.setFillColor(text)
    pdf.drawCentredString(width / 2, rate_y + 29, "단기납과 같은 예상 이익을 내려면")
    set_font(13)
    pdf.setFillColor(navy)
    pdf.drawCentredString(width / 2, rate_y + 12, f"적금금리 연 {required_rate:,.2f}% 필요")

    # 하단 안내
    note = (
        "적금은 1년 만기 상품을 동일 조건으로 10회 반복한 단리 계산이며 원금 재예치에 따른 복리는 반영하지 않습니다. "
        "단기납은 10년에 반드시 해지해야 하는 상품이 아니며, 계속 유지할 경우 상품의 해지환급금 예시표에 따라 환급금이 추가로 증가할 수 있습니다. "
        "실제 해지환급금과 비과세 적용 여부는 상품의 설계서, 계약조건 및 관련 요건에 따라 달라질 수 있습니다."
    )
    draw_wrapped(note, margin, 119, width - margin * 2, 6.3, 9.2, muted)

    pdf.setStrokeColor(line)
    pdf.line(margin, 42, width - margin, 42)
    set_font(6.3)
    pdf.setFillColor(muted)
    pdf.drawString(margin, 29, "비전본부 드림지점 박병선 팀장")
    pdf.drawRightString(width - margin, 29, "화랑 WORKSPACE")

    pdf.showPage()
    pdf.save()
    buffer.seek(0)
    return buffer.getvalue()


def render_styles() -> None:
    st.markdown(
        """
        <style>
        :root {
            --hw-navy: #16324f;
            --hw-blue: #2f6fa3;
            --hw-blue-soft: #eaf2f8;
            --hw-gold: #c9963d;
            --hw-gold-deep: #a87422;
            --hw-gold-soft: #fbf5e9;
            --hw-text: #203247;
            --hw-muted: #6e7e90;
            --hw-line: #dce4ec;
            --hw-surface: #ffffff;
        }

        h1 a, h2 a, h3 a { display: none !important; }

        div[data-testid="stForm"] {
            padding: 24px 24px 20px;
            border: 1px solid rgba(47, 111, 163, 0.18);
            border-radius: 18px;
            background:
                radial-gradient(circle at 100% 0%, rgba(201,150,61,.10), transparent 32%),
                linear-gradient(145deg, rgba(255,255,255,.98), rgba(244,248,252,.98));
            box-shadow: 0 12px 30px rgba(22, 50, 79, 0.07);
        }

        div[data-testid="stForm"] label p {
            color: var(--hw-text);
            font-weight: 650;
        }

        div[data-testid="stForm"] div[data-baseweb="input"] > div,
        div[data-testid="stForm"] div[data-baseweb="select"] > div {
            border-color: rgba(47, 111, 163, 0.20);
            background: rgba(255,255,255,.94);
        }

        div[data-testid="stFormSubmitButton"] button {
            min-height: 48px;
            border: 0;
            border-radius: 12px;
            color: white;
            font-weight: 750;
            background: linear-gradient(135deg, var(--hw-navy), var(--hw-blue));
            box-shadow: 0 8px 18px rgba(22, 50, 79, 0.18);
            transition: transform .16s ease, box-shadow .16s ease;
        }

        div[data-testid="stFormSubmitButton"] button:hover {
            transform: translateY(-1px);
            box-shadow: 0 11px 24px rgba(22, 50, 79, 0.24);
        }

        div[data-testid="stDownloadButton"] button {
            min-height: 47px;
            border: 1px solid rgba(201,150,61,.52);
            border-radius: 12px;
            color: #704b16;
            font-weight: 750;
            background: linear-gradient(135deg, #fffaf0, var(--hw-gold-soft));
            box-shadow: 0 7px 18px rgba(122,83,23,.09);
        }

        div[data-testid="stDownloadButton"] button:hover {
            border-color: var(--hw-gold);
            color: #5f3f12;
            background: #fff8e8;
        }

        .hw-input-heading {
            display: flex;
            align-items: center;
            gap: 10px;
            margin: 0 0 14px;
            color: var(--hw-navy);
            font-size: 17px;
            font-weight: 750;
        }

        .hw-input-heading::before {
            content: "";
            width: 5px;
            height: 19px;
            border-radius: 999px;
            background: linear-gradient(var(--hw-gold), var(--hw-gold-deep));
        }

        .hw-result-hero {
            margin: 22px 0 4px;
            padding: 25px 18px 22px;
            text-align: center;
            border: 1px solid rgba(201,150,61,.26);
            border-radius: 18px;
            background: linear-gradient(135deg, var(--hw-gold-soft), #ffffff 68%);
        }

        .hw-result-context { color: var(--hw-muted); font-size: 14px; }
        .hw-result-title { margin-top: 6px; color: var(--hw-navy); font-size: 28px; font-weight: 800; }
        .hw-result-title strong { color: var(--hw-gold-deep); }
        .hw-result-basis { margin-top: 7px; color: var(--hw-muted); font-size: 13px; }

        .hw-chart {
            position: relative;
            min-height: 430px;
            display: flex;
            justify-content: center;
            align-items: flex-end;
            gap: clamp(64px, 13vw, 145px);
            padding: 54px 28px 18px;
            margin: 4px 0 12px;
            border-bottom: 1px solid var(--hw-line);
        }

        .hw-chart-grid {
            position: absolute;
            inset: 54px 0 64px;
            z-index: 0;
            background: repeating-linear-gradient(
                to bottom,
                rgba(110,126,144,.12) 0,
                rgba(110,126,144,.12) 1px,
                transparent 1px,
                transparent 74px
            );
        }

        .hw-bar-group { position: relative; z-index: 1; width: min(176px, 31vw); text-align: center; }
        .hw-bar-value { margin-bottom: 8px; color: var(--hw-navy); font-size: 19px; font-weight: 800; }
        .hw-bar {
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 56px;
            border-radius: 9px 9px 2px 2px;
            box-shadow: 0 8px 18px rgba(22,50,79,.10);
        }
        .hw-bar span { font-size: 13px; line-height: 1.38; font-weight: 750; }
        .hw-deposit-bar { color: white; background: linear-gradient(180deg, #5e91bb, var(--hw-blue)); }
        .hw-shortpay-bar { color: white; background: linear-gradient(180deg, #ddb766, var(--hw-gold-deep)); }
        .hw-bar-name { margin-top: 11px; color: var(--hw-navy); font-size: 17px; font-weight: 800; }
        .hw-bar-detail { margin-top: 3px; color: var(--hw-muted); font-size: 12px; }

        .hw-chart-badge {
            position: absolute;
            left: calc(100% + 13px);
            top: 35px;
            padding: 7px 11px;
            white-space: nowrap;
            border: 1px solid rgba(201,150,61,.55);
            border-radius: 999px;
            color: #7a5317;
            background: var(--hw-gold-soft);
            font-size: 12px;
            font-weight: 750;
        }

        .hw-timeline {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            margin: 24px 0 26px;
        }

        .hw-phase {
            position: relative;
            padding: 15px 7px 0;
            text-align: center;
            border-top: 2px solid var(--hw-line);
        }

        .hw-phase::before {
            content: "";
            position: absolute;
            top: -6px;
            left: calc(50% - 5px);
            width: 10px;
            height: 10px;
            border-radius: 50%;
            background: var(--hw-line);
        }

        .hw-phase-main { color: var(--hw-navy); font-size: 13px; font-weight: 750; }
        .hw-phase-sub { margin-top: 3px; color: var(--hw-muted); font-size: 11px; line-height: 1.35; }
        .hw-phase-point { border-top-color: var(--hw-gold); }
        .hw-phase-point::before { background: var(--hw-gold); box-shadow: 0 0 0 4px rgba(201,150,61,.14); }

        .hw-calc-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 18px;
            margin-top: 8px;
        }

        .hw-calc-card {
            padding: 18px 19px 15px;
            border: 1px solid var(--hw-line);
            border-radius: 14px;
            background: var(--hw-surface);
            box-shadow: 0 7px 19px rgba(22,50,79,.05);
        }

        .hw-calc-title { margin-bottom: 9px; color: var(--hw-navy); font-size: 15px; font-weight: 800; }
        .hw-calc-row { display: flex; justify-content: space-between; gap: 16px; padding: 7px 0; border-bottom: 1px solid rgba(220,228,236,.72); color: var(--hw-muted); font-size: 13px; }
        .hw-calc-row:last-child { border-bottom: 0; }
        .hw-calc-row span:last-child { color: var(--hw-text); font-weight: 700; text-align: right; }
        .hw-calc-result span { color: var(--hw-navy) !important; font-weight: 800 !important; }

        .hw-rate-box {
            margin-top: 18px;
            padding: 18px;
            text-align: center;
            border: 1px solid rgba(47,111,163,.17);
            border-radius: 13px;
            color: var(--hw-text);
            background: var(--hw-blue-soft);
        }
        .hw-rate-box strong { color: var(--hw-navy); font-size: 21px; }
        .hw-rate-sub { margin-top: 5px; color: var(--hw-muted); font-size: 12px; }

        .hw-note { margin-top: 14px; color: var(--hw-muted); font-size: 11px; line-height: 1.55; }

        @media (max-width: 680px) {
            .hw-result-title { font-size: 23px; }
            .hw-chart { gap: 36px; padding-left: 8px; padding-right: 8px; }
            .hw-bar-group { width: 132px; }
            .hw-chart-badge { left: auto; right: -14px; top: 10px; font-size: 10px; }
            .hw-timeline { grid-template-columns: 1fr 1fr; gap: 22px 0; }
            .hw-calc-grid { grid-template-columns: 1fr; }
        }

        @media print {
            header, footer, [data-testid="stSidebar"], [data-testid="stForm"], [data-testid="stExpander"] {
                display: none !important;
            }
            .block-container { padding: .35rem 1rem 0 !important; }
            .hw-chart { min-height: 380px; }
            .hw-calc-card { box-shadow: none; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_bar_chart(
    monthly: float,
    pay_years: int,
    deposit_interest: float,
    refund_gain: float,
    advantage: float,
) -> None:
    chart_max = max(deposit_interest, refund_gain, 1)
    deposit_height = max(56, min(300, deposit_interest / chart_max * 300))
    shortpay_height = max(56, min(300, refund_gain / chart_max * 300))

    if advantage >= 0:
        badge = f'<div class="hw-chart-badge">적금 대비 +{format_currency(advantage)}</div>'
    else:
        badge = '<div class="hw-chart-badge">현재 조건은 적금 우위</div>'

    st.markdown(
        f"""
        <div class="hw-chart" role="img" aria-label="적금 10년 누적 세후이자와 단기납 10년 예상 환급차익 비교">
            <div class="hw-chart-grid"></div>
            <div class="hw-bar-group">
                <div class="hw-bar-value">{format_currency(deposit_interest)}</div>
                <div class="hw-bar hw-deposit-bar" style="height:{deposit_height:.1f}px">
                    <span>10년 누적<br>세후이자</span>
                </div>
                <div class="hw-bar-name">적금</div>
                <div class="hw-bar-detail">월 {format_currency(monthly)} · 1년 적금 10회</div>
            </div>
            <div class="hw-bar-group">
                {badge}
                <div class="hw-bar-value">{format_currency(refund_gain)}</div>
                <div class="hw-bar hw-shortpay-bar" style="height:{shortpay_height:.1f}px">
                    <span>10년 시점<br>예상 환급차익</span>
                </div>
                <div class="hw-bar-name">단기납</div>
                <div class="hw-bar-detail">월 {format_currency(monthly)} · {pay_years}년납 후 유지</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_timeline(pay_years: int) -> None:
    holding_years = 10 - pay_years
    st.markdown(
        f"""
        <div class="hw-timeline">
            <div class="hw-phase">
                <div class="hw-phase-main">1~{pay_years}년 납입</div>
                <div class="hw-phase-sub">보험료 납입</div>
            </div>
            <div class="hw-phase">
                <div class="hw-phase-main">{pay_years + 1}~9년 유지</div>
                <div class="hw-phase-sub">추가납입 없이 약 {holding_years}년 유지</div>
            </div>
            <div class="hw-phase hw-phase-point">
                <div class="hw-phase-main">10년 주요 시점</div>
                <div class="hw-phase-sub">환급률·비과세 요건 확인</div>
            </div>
            <div class="hw-phase">
                <div class="hw-phase-main">해지 또는 계속 유지</div>
                <div class="hw-phase-sub">설계서에 따라 환급금 추가 증가 가능</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_calculation_details(deposit: dict, shortpay: dict, pay_years: int, refund_rate: float) -> None:
    st.markdown(
        f"""
        <div class="hw-calc-grid">
            <div class="hw-calc-card">
                <div class="hw-calc-title">적금 계산 내역</div>
                <div class="hw-calc-row"><span>1년 납입원금</span><span>{format_currency(deposit['one_year_principal'])}</span></div>
                <div class="hw-calc-row"><span>1년 세전이자</span><span>{format_currency(deposit['pretax_interest'])}</span></div>
                <div class="hw-calc-row"><span>이자소득세 15.4%</span><span>{format_currency(deposit['tax'])}</span></div>
                <div class="hw-calc-row"><span>1년 세후이자</span><span>{format_currency(deposit['aftertax_interest'])}</span></div>
                <div class="hw-calc-row hw-calc-result"><span>10년 누적 세후이자</span><span>{format_currency(deposit['ten_year_interest'])}</span></div>
            </div>
            <div class="hw-calc-card">
                <div class="hw-calc-title">단기납 계산 내역</div>
                <div class="hw-calc-row"><span>납입기간</span><span>{pay_years}년</span></div>
                <div class="hw-calc-row"><span>총납입보험료</span><span>{format_currency(shortpay['total_premium'])}</span></div>
                <div class="hw-calc-row"><span>10년 예상 환급률</span><span>{refund_rate:,.1f}%</span></div>
                <div class="hw-calc-row"><span>10년 예상 해지환급금</span><span>{format_currency(shortpay['refund_amount'])}</span></div>
                <div class="hw-calc-row hw-calc-result"><span>예상 환급차익</span><span>{format_currency(shortpay['refund_gain'])}</span></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def run():
    render_styles()

    with st.expander("인쇄 방법 및 계산 기준"):
        st.markdown(
            """
            **계산 기준**

            - 적금은 1년 만기 상품을 동일한 월납입액과 금리로 10회 반복한 단리 계산입니다.
            - 매년 만기 원금의 재예치와 복리 효과는 반영하지 않습니다.
            - 적금 이자에는 이자소득세 15.4%를 반영합니다.
            - 단기납은 입력한 10년 시점 예상 해지환급률을 기준으로 계산합니다.

            **인쇄 안내**

            오른쪽 위 메뉴에서 Print를 선택한 뒤 머리글과 바닥글을 해제하고 배경 그래픽을 켜주세요.
            """
        )
        st.markdown(
            """
            <div style="margin-top:14px; color:#6e7e90; font-size:12px;">
                제작자: 비전본부 드림지점 박병선 팀장 · 버전 v2.0.0 · 2026-08-06
            </div>
            """,
            unsafe_allow_html=True,
        )

    page_header(
        "고객 상담",
        "적금 vs 단기납",
        "같은 월납입금액으로 10년 예상 이익을 간편하게 비교합니다.",
        "DS",
    )

    with st.form("hwarang_deposit_shortpay_form"):
        st.markdown('<div class="hw-input-heading">상담 조건 입력</div>', unsafe_allow_html=True)
        left, right = st.columns(2, gap="large")

        with left:
            monthly = st.number_input(
                "공통 월납입금액 (만원)",
                min_value=1,
                step=10,
                value=100,
                format="%d",
                help="적금과 단기납에 동일하게 적용되는 월납입금액입니다.",
            )
            annual_rate = st.number_input(
                "적금 연이율 (%)",
                min_value=0.1,
                max_value=30.0,
                step=0.1,
                value=3.0,
                format="%.1f",
                help="1년 만기 적금의 세전 연이율을 입력하세요.",
            )

        with right:
            pay_years = st.selectbox(
                "단기납 납입기간",
                [5, 7],
                index=0,
                format_func=lambda value: f"{value}년납",
            )
            refund_rate = st.number_input(
                "10년 시점 예상 해지환급률 (%)",
                min_value=100.0,
                max_value=300.0,
                step=0.1,
                value=123.0,
                format="%.1f",
                help="해당 상품의 가입설계서에 기재된 10년 시점 환급률을 입력하세요.",
            )

        submitted = st.form_submit_button("10년 예상 이익 비교", use_container_width=True)

    if submitted:
        st.session_state["hwarang_ds_result"] = {
            "monthly": float(monthly),
            "annual_rate": float(annual_rate),
            "pay_years": int(pay_years),
            "refund_rate": float(refund_rate),
        }

    values = st.session_state.get("hwarang_ds_result")
    if not values:
        st.info("네 가지 상담 조건을 확인한 뒤 ‘10년 예상 이익 비교’를 눌러주세요.")
        return

    with st.spinner("예상 결과를 계산하고 있습니다..."):
        time.sleep(0.25)

    monthly = values["monthly"]
    annual_rate = values["annual_rate"]
    pay_years = values["pay_years"]
    refund_rate = values["refund_rate"]

    deposit = calculate_deposit(monthly, annual_rate)
    shortpay = calculate_shortpay(monthly, pay_years, refund_rate)
    advantage = shortpay["refund_gain"] - deposit["ten_year_interest"]

    if advantage >= 0:
        headline = f"단기납의 예상 환급차익이 <strong>{format_currency(advantage)} 더 큽니다</strong>"
    else:
        headline = f"현재 조건에서는 적금 누적 세후이자가 <strong>{format_currency(abs(advantage))} 더 큽니다</strong>"

    st.markdown(
        f"""
        <div class="hw-result-hero">
            <div class="hw-result-context">같은 월 {format_currency(monthly)}을 활용했을 때</div>
            <div class="hw-result-title">{headline}</div>
            <div class="hw-result-basis">적금 1년 만기 10회 반복 · 단기납 {pay_years}년납 후 10년 시점</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    render_bar_chart(
        monthly,
        pay_years,
        deposit["ten_year_interest"],
        shortpay["refund_gain"],
        advantage,
    )
    render_timeline(pay_years)
    render_calculation_details(deposit, shortpay, pay_years, refund_rate)

    required_rate = calculate_required_deposit_rate(
        monthly,
        shortpay["refund_gain"],
        deposit["interest_weight"],
    )
    required_monthly = calculate_required_monthly_payment(
        annual_rate,
        shortpay["refund_gain"],
        deposit["interest_weight"],
    )

    st.markdown(
        f"""
        <div class="hw-rate-box">
            단기납과 같은 예상 이익을 내려면 적금금리가 연 <strong>{required_rate:,.2f}%</strong> 필요합니다.
            <div class="hw-rate-sub">현재 금리 연 {annual_rate:,.1f}%를 유지한다면 월납입액은 약 {format_currency(required_monthly)}이 필요합니다.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="hw-note">
            적금은 1년 만기 상품을 동일 조건으로 10회 반복한 단리 계산이며 원금 재예치에 따른 복리는 반영하지 않습니다.
            단기납은 10년에 반드시 해지해야 하는 상품이 아니며, 계속 유지하는 경우 상품의 해지환급금 예시표에 따라 환급금이 추가로 증가할 수 있습니다.
            실제 해지환급금과 비과세 적용 여부는 해당 상품의 설계서, 계약조건 및 관련 요건에 따라 달라질 수 있습니다.
        </div>
        """,
        unsafe_allow_html=True,
    )

    try:
        pdf_bytes = create_result_pdf(
            monthly,
            annual_rate,
            pay_years,
            refund_rate,
            deposit,
            shortpay,
            advantage,
            required_rate,
        )
        download_name = f"적금_단기납_비교결과_{datetime.now():%Y%m%d}.pdf"
        st.download_button(
            "A4 상담 결과 PDF 다운로드",
            data=pdf_bytes,
            file_name=download_name,
            mime="application/pdf",
            use_container_width=True,
        )
    except ImportError:
        st.warning(
            "PDF 다운로드 기능을 사용하려면 requirements.txt에 reportlab을 추가해 주세요."
        )
