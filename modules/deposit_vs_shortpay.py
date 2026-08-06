import time

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
            padding: 38px 28px 18px;
            margin: 0 0 12px;
            border-bottom: 1px solid var(--hw-line);
        }

        .hw-chart-grid {
            position: absolute;
            inset: 38px 0 64px;
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
        .hw-bar-value-deposit { color: #58718a; }
        .hw-bar-value-shortpay {
            color: var(--hw-gold-deep);
            font-size: 23px;
            text-shadow: 0 2px 10px rgba(168,116,34,.16);
        }
        .hw-bar {
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 56px;
            border-radius: 9px 9px 2px 2px;
            box-shadow: 0 8px 18px rgba(22,50,79,.10);
        }
        .hw-bar span { font-size: 13px; line-height: 1.38; font-weight: 750; }
        .hw-deposit-bar {
            color: white;
            background: linear-gradient(180deg, #8eacc5, #5f83a2);
            opacity: .84;
        }
        .hw-shortpay-bar {
            color: white;
            background: linear-gradient(180deg, #e3bc69, var(--hw-gold-deep));
            box-shadow: 0 13px 28px rgba(168,116,34,.25), 0 0 0 3px rgba(201,150,61,.09);
        }
        .hw-bar-name { margin-top: 11px; color: var(--hw-navy); font-size: 17px; font-weight: 800; }
        .hw-bar-detail { margin-top: 3px; color: var(--hw-muted); font-size: 12px; }

        .hw-profit-badge {
            position: absolute;
            left: 38%;
            z-index: 4;
            width: max-content;
            max-width: 230px;
            padding: 12px 18px 11px;
            transform: translateX(-50%);
            overflow: hidden;
            white-space: nowrap;
            border: 1px solid rgba(47,111,163,.30);
            border-radius: 15px;
            color: var(--hw-navy);
            background: linear-gradient(145deg, rgba(255,253,248,.98), rgba(244,248,252,.98));
            box-shadow:
                0 12px 28px rgba(22,50,79,.12),
                inset 0 1px 0 rgba(255,255,255,.96);
            text-align: center;
        }

        .hw-profit-badge::before {
            content: "";
            position: absolute;
            inset: 0 0 auto;
            height: 3px;
            background: linear-gradient(90deg, var(--hw-blue), var(--hw-gold), var(--hw-gold-deep));
        }

        .hw-profit-badge span {
            display: block;
            color: #61758a;
            font-size: 12px;
            font-weight: 750;
            letter-spacing: -.15px;
        }

        .hw-profit-badge strong {
            display: block;
            margin-top: 2px;
            color: #c43f3f;
            font-size: 23px;
            font-weight: 900;
            letter-spacing: -.6px;
            text-shadow: 0 3px 12px rgba(196,63,63,.12);
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
        .hw-calc-card-deposit { border-top: 3px solid rgba(47,111,163,.58); }
        .hw-calc-card-shortpay { border-top: 3px solid rgba(201,150,61,.72); }
        .hw-calc-card-deposit .hw-calc-result span:last-child { color: var(--hw-blue) !important; }
        .hw-calc-card-shortpay .hw-calc-result span:last-child {
            color: var(--hw-gold-deep) !important;
            font-size: 16px;
        }

        .hw-rate-box {
            margin-top: 18px;
            display: grid;
            grid-template-columns: 1fr 1fr;
            padding: 0;
            border: 1px solid rgba(47,111,163,.22);
            border-radius: 16px;
            color: var(--hw-text);
            background: linear-gradient(135deg, #f4f9fd, var(--hw-blue-soft));
            box-shadow: 0 9px 22px rgba(22,50,79,.07);
            overflow: hidden;
        }
        .hw-rate-panel {
            min-width: 0;
            padding: 21px 24px 19px;
            text-align: center;
        }
        .hw-rate-panel + .hw-rate-panel {
            border-left: 1px solid rgba(47,111,163,.18);
            background: rgba(255,255,255,.34);
        }
        .hw-rate-label { color: var(--hw-muted); font-size: 14px; font-weight: 650; }
        .hw-rate-main {
            display: flex;
            align-items: center;
            justify-content: center;
            flex-wrap: wrap;
            gap: 2px 4px;
            min-height: 44px;
            margin-top: 5px;
            color: var(--hw-navy);
            font-size: 18px;
            font-weight: 750;
        }
        .hw-rate-percent,
        .hw-required-amount {
            display: inline-block;
            margin: 0 4px;
            font-size: 34px;
            line-height: 1.1;
            font-weight: 900;
            letter-spacing: -.5px;
        }
        .hw-rate-percent {
            color: #d83b3b;
            text-shadow: 0 3px 13px rgba(216,59,59,.13);
        }
        .hw-required-amount {
            color: #176fa7;
            text-shadow: 0 3px 13px rgba(23,111,167,.12);
        }
        .hw-current-rate { color: var(--hw-navy); font-weight: 850; }

        .hw-note { margin-top: 14px; color: var(--hw-muted); font-size: 11px; line-height: 1.55; }

        @media (max-width: 680px) {
            .hw-result-title { font-size: 23px; }
            .hw-chart { gap: 36px; padding-left: 8px; padding-right: 8px; }
            .hw-bar-group { width: 132px; }
            .hw-profit-badge {
                left: 34%;
                max-width: 174px;
                padding: 9px 11px 8px;
            }
            .hw-profit-badge span { font-size: 10px; }
            .hw-profit-badge strong { font-size: 18px; }
            .hw-timeline { grid-template-columns: 1fr 1fr; gap: 22px 0; }
            .hw-calc-grid { grid-template-columns: 1fr; }
            .hw-rate-box { grid-template-columns: 1fr; }
            .hw-rate-panel + .hw-rate-panel {
                border-top: 1px solid rgba(47,111,163,.18);
                border-left: 0;
            }
            .hw-rate-percent,
            .hw-required-amount { font-size: 29px; }
        }

        @page { size: A4 portrait; margin: 9mm; }

        @media print {
            header, footer, [data-testid="stSidebar"], [data-testid="stForm"], [data-testid="stExpander"] {
                display: none !important;
            }

            html, body {
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }

            .block-container {
                max-width: none !important;
                padding: 0 !important;
            }

            .hw-result-hero {
                margin-top: 0;
                padding: 18px 16px 16px;
                break-inside: avoid;
            }
            .hw-result-context { font-size: 11pt; }
            .hw-result-title { font-size: 23pt; }
            .hw-result-basis { font-size: 10.5pt; }

            .hw-chart {
                min-height: 365px;
                padding-top: 28px;
                break-inside: avoid;
            }
            .hw-bar-value { font-size: 15pt; }
            .hw-bar-value-shortpay { font-size: 18pt; }
            .hw-bar span { font-size: 10.5pt; }
            .hw-bar-name { font-size: 14pt; }
            .hw-bar-detail { font-size: 10pt; }
            .hw-profit-badge span { font-size: 10pt; }
            .hw-profit-badge strong { font-size: 18pt; }

            .hw-timeline {
                margin: 18px 0 20px;
                break-inside: avoid;
            }
            .hw-phase-main { font-size: 10.5pt; }
            .hw-phase-sub { font-size: 9.5pt; }

            .hw-calc-grid,
            .hw-rate-box { break-inside: avoid; }
            .hw-calc-card {
                padding: 14px 16px 12px;
                box-shadow: none;
            }
            .hw-calc-title { font-size: 12pt; }
            .hw-calc-row {
                padding: 5px 0;
                font-size: 10.5pt;
            }
            .hw-calc-card-shortpay .hw-calc-result span:last-child { font-size: 12.5pt; }

            .hw-rate-panel { padding: 16px 18px 14px; }
            .hw-rate-label { font-size: 11pt; }
            .hw-rate-main { font-size: 14pt; }
            .hw-rate-percent,
            .hw-required-amount { font-size: 25pt; }
            .hw-note { font-size: 9.5pt; line-height: 1.5; }
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
    gain_multiple = refund_gain / deposit_interest if deposit_interest > 0 else 0

    # 기존 프리미엄 배지 위치를 막대 높이에 맞춰 유지합니다.
    bar_baseline_y = 358
    deposit_top_y = bar_baseline_y - deposit_height
    shortpay_top_y = bar_baseline_y - shortpay_height
    badge_top = max(64, min(184, (deposit_top_y + shortpay_top_y) / 2 - 62))

    if gain_multiple >= 1:
        badge_eyebrow = "단기납 예상 이익"
        badge_value = f"적금의 약 {gain_multiple:,.1f}배"
    else:
        badge_eyebrow = "현재 조건의 예상 이익"
        badge_value = "적금이 더 큽니다"

    st.markdown(
        f"""
        <div class="hw-chart" role="img" aria-label="적금 10년 누적 세후이자와 단기납 10년 예상 환급차익 비교">
            <div class="hw-chart-grid"></div>
            <div class="hw-profit-badge" style="top:{badge_top:.1f}px">
                <span>{badge_eyebrow}</span>
                <strong>{badge_value}</strong>
            </div>
            <div class="hw-bar-group">
                <div class="hw-bar-value hw-bar-value-deposit">{format_currency(deposit_interest)}</div>
                <div class="hw-bar hw-deposit-bar" style="height:{deposit_height:.1f}px">
                    <span>10년 누적<br>세후이자</span>
                </div>
                <div class="hw-bar-name">적금</div>
                <div class="hw-bar-detail">월 {format_currency(monthly)} · 1년 적금 10회</div>
            </div>
            <div class="hw-bar-group">
                <div class="hw-bar-value hw-bar-value-shortpay">{format_currency(refund_gain)}</div>
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
    if pay_years >= 10:
        holding_title = "별도 거치 없음"
        holding_subtitle = "10년까지 보험료 납입"
    else:
        holding_title = f"{pay_years + 1}~9년 유지"
        holding_subtitle = f"추가납입 없이 약 {holding_years}년 유지"
    st.markdown(
        f"""
        <div class="hw-timeline">
            <div class="hw-phase">
                <div class="hw-phase-main">1~{pay_years}년 납입</div>
                <div class="hw-phase-sub">보험료 납입</div>
            </div>
            <div class="hw-phase">
                <div class="hw-phase-main">{holding_title}</div>
                <div class="hw-phase-sub">{holding_subtitle}</div>
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
            <div class="hw-calc-card hw-calc-card-deposit">
                <div class="hw-calc-title">적금 계산 내역</div>
                <div class="hw-calc-row"><span>1년 납입원금</span><span>{format_currency(deposit['one_year_principal'])}</span></div>
                <div class="hw-calc-row"><span>1년 세전이자</span><span>{format_currency(deposit['pretax_interest'])}</span></div>
                <div class="hw-calc-row"><span>이자소득세 15.4%</span><span>{format_currency(deposit['tax'])}</span></div>
                <div class="hw-calc-row"><span>1년 세후이자</span><span>{format_currency(deposit['aftertax_interest'])}</span></div>
                <div class="hw-calc-row hw-calc-result"><span>10년 누적 세후이자</span><span>{format_currency(deposit['ten_year_interest'])}</span></div>
            </div>
            <div class="hw-calc-card hw-calc-card-shortpay">
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

            - Ctrl + P -> 설정 더보기를 누릅니다. 
            - 배율을 맞춤설정으로 변경 후 78로 조정.
            - 머리글과 바닥글, 배경그래픽 체크 해제 후 인쇄.
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
                max_value=100.0,
                step=0.1,
                value=3.0,
                format="%.1f",
                help="1년 만기 적금의 세전 연이율을 입력하세요.",
            )

        with right:
            pay_years = st.selectbox(
                "단기납 납입기간",
                [5, 7, 10],
                index=0,
                format_func=lambda value: f"{value}년납",
            )
            refund_rate = st.number_input(
                "10년 시점 예상 해지환급률 (%)",
                min_value=100.0,
                max_value=300.0,
                step=0.1,
                value=120.0,
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
            <div class="hw-rate-panel">
                <div class="hw-rate-label">단기납과 같은 예상 이익을 내려면</div>
                <div class="hw-rate-main">적금금리가 연 <span class="hw-rate-percent">{required_rate:,.2f}%</span> 필요합니다.</div>
            </div>
            <div class="hw-rate-panel">
                <div class="hw-rate-label">현재 금리 연 <span class="hw-current-rate">{annual_rate:,.1f}%</span>를 유지한다면</div>
                <div class="hw-rate-main">월납입액은 약 <span class="hw-required-amount">{format_currency(required_monthly)}</span>이 필요합니다.</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="hw-note">
            적금은 1년 만기 상품을 동일 조건으로 10회 반복한 단리 계산입니다.
            단기납은 10년에 반드시 해지해야 하는 상품이 아니며, 계속 유지하는 경우 상품의 해지환급금 예시표에 따라 환급금이 추가로 증가할 수 있습니다.
            실제 해지환급금과 비과세 적용 여부는 해당 상품의 설계서, 계약조건 및 관련 요건에 따라 달라질 수 있습니다.
        </div>
        """,
        unsafe_allow_html=True,
    )
