import html
from dataclasses import dataclass
from textwrap import dedent

import pandas as pd
import streamlit as st

from .ui_components import page_header


REFERENCE_MULTIPLIERS = {
    5: [1.22, 1.18, 1.15, 1.12, 1.10, 1.08],
    10: [1.50, 1.30, 1.20, 1.12, 1.08],
    15: [1.80, 1.38, 1.20, 1.10],
    20: [2.10, 1.40, 1.18, 1.08],
    30: [2.60, 1.35, 1.12],
}

SCENARIOS = {
    "완만한 상승": {"weight": 0.70, "cap": 2.80},
    "기준 상승": {"weight": 1.00, "cap": 3.50},
    "높은 상승": {"weight": 1.30, "cap": 4.20},
}


@dataclass
class RenewalPeriod:
    start_age: int
    end_age: int
    monthly_premium: float
    cumulative_multiple: float
    applied_multiple: float | None = None

    @property
    def years(self) -> int:
        return self.end_age - self.start_age

    @property
    def total(self) -> float:
        return self.monthly_premium * self.years * 12


def _won(value: float) -> str:
    return f"{int(round(value)):,}원"


def _compact_won(value: float) -> str:
    if value >= 100_000_000:
        result = f"{value / 100_000_000:.1f}".rstrip("0").rstrip(".")
        return f"{result}억원"
    if value >= 10_000:
        return f"{value / 10_000:,.0f}만원"
    return _won(value)


def _age_weight(age: int) -> float:
    if age < 40:
        return 0.90
    if age < 50:
        return 0.95
    if age < 60:
        return 1.00
    if age < 70:
        return 1.05
    return 1.10


def _scenario_multiple(
    scenario: str,
    cycle: int,
    renewal_index: int,
    renewal_age: int,
) -> float:
    reference = REFERENCE_MULTIPLIERS[cycle]
    base = reference[min(renewal_index, len(reference) - 1)]
    weight = SCENARIOS[scenario]["weight"]
    return 1 + (base - 1) * weight * _age_weight(renewal_age)


def _renewal_ages(current_age: int, end_age: int, cycle: int, next_years: int) -> list[int]:
    ages: list[int] = []
    age = current_age + next_years
    while age < end_age:
        ages.append(age)
        age += cycle
    return ages


def _calculate_periods(
    current_age: int,
    end_age: int,
    cycle: int,
    next_years: int,
    current_premium: float,
    method: str,
    scenario: str,
    custom_values: list[float],
) -> list[RenewalPeriod]:
    periods: list[RenewalPeriod] = []
    premium = current_premium
    age = current_age
    first_period = True
    renewal_index = 0

    while age < end_age:
        duration = min(next_years if first_period else cycle, end_age - age)
        period_end = age + duration
        periods.append(
            RenewalPeriod(
                start_age=age,
                end_age=period_end,
                monthly_premium=premium,
                cumulative_multiple=(premium / current_premium) if current_premium else 1,
            )
        )
        age = period_end

        if age < end_age:
            applied_multiple: float | None = None
            if method == "가입제안서 직접 입력":
                if renewal_index < len(custom_values) and custom_values[renewal_index] > 0:
                    premium = custom_values[renewal_index]
            elif method == "갱신배수 직접 설정":
                applied_multiple = (
                    custom_values[renewal_index]
                    if renewal_index < len(custom_values) and custom_values[renewal_index] > 0
                    else 1
                )
                premium *= applied_multiple
            else:
                applied_multiple = _scenario_multiple(
                    scenario, cycle, renewal_index, age
                )
                premium *= applied_multiple
                premium = min(
                    premium,
                    current_premium * SCENARIOS[scenario]["cap"],
                )

            if periods:
                periods[-1].applied_multiple = applied_multiple
            renewal_index += 1

        first_period = False

    return periods


def _inject_style() -> None:
    st.markdown(
        dedent(
            """
        <style>
        .block-container {max-width: 1180px; padding-top: 1.3rem; padding-bottom: 3rem;}
        .rn-section-title {font-size: 1.04rem; font-weight: 750; color: #173451; margin: .25rem 0 .8rem;}
        .rn-card-label {display:flex; align-items:center; gap:.55rem; margin-bottom:.85rem; color:#173451; font-weight:750;}
        .rn-card-number {display:inline-flex; align-items:center; justify-content:center; width:1.7rem; height:1.7rem; border-radius:.55rem; color:#fff; background:#346f9f; font-size:.8rem;}
        div[data-testid="stVerticalBlockBorderWrapper"] {border-color:#dce7f0; border-radius:18px; box-shadow:0 9px 28px rgba(21,52,81,.055);}
        .rn-result-hero {margin:1.1rem 0; padding:1.3rem 1.4rem; text-align:center; border:1px solid #d8e5ef; border-radius:18px; background:linear-gradient(135deg,#f8fbfe 0%,#edf5fb 100%); color:#173451;}
        .rn-result-kicker {font-size:.86rem; color:#657f96; margin-bottom:.35rem;}
        .rn-result-copy {font-size:1.18rem; font-weight:700;}
        .rn-result-value {display:inline-block; margin-left:.25rem; padding:.04rem .42rem .08rem; color:#c83f45; background:rgba(200,63,69,.065); border-bottom:2px solid rgba(200,63,69,.42); border-radius:.42rem; font-size:1.72rem; font-weight:850;}
        .rn-result-value.negative {color:#a66516; background:rgba(166,101,22,.07); border-bottom-color:rgba(166,101,22,.38);}
        .rn-metric-grid {display:grid; grid-template-columns:1fr 1fr; gap:.85rem; margin:0 0 1.15rem;}
        .rn-metric {padding:1rem 1.1rem; border:1px solid #dce7f0; border-radius:16px; background:#fff; color:#173451;}
        .rn-metric-label {font-size:.86rem; color:#657f96;}
        .rn-metric-value {margin:.28rem 0 .1rem; font-size:1.42rem; font-weight:850;}
        .rn-metric-sub {font-size:.82rem; color:#71889c;}
        .rn-chart-grid {display:grid; grid-template-columns:1fr 1fr; gap:1.2rem; margin-top:.5rem;}
        .rn-panel {min-width:0;}
        .rn-panel-title {display:flex; align-items:center; gap:.5rem; margin-bottom:.65rem; color:#173451; font-weight:800;}
        .rn-dot {width:.68rem; height:.68rem; border-radius:50%; background:#4d86af;}
        .rn-dot.fixed {background:#c79a49;}
        .rn-chart-zone {height:238px; border-bottom:2px solid #cad9e4; background:linear-gradient(180deg,#fbfdff 0%,#f4f8fb 100%); box-shadow:inset 0 -18px 24px rgba(44,80,106,.025);}
        .rn-bars {display:flex; align-items:stretch; gap:7px; height:100%;}
        .rn-column {position:relative; display:flex; flex:1 1 0; flex-direction:column; justify-content:flex-end; min-width:0; padding-bottom:29px; text-align:center;}
        .rn-bar-value {align-self:center; padding:.12rem .45rem; color:#365c78; background:rgba(255,255,255,.82); border:1px solid rgba(92,127,153,.16); border-radius:999px; box-shadow:0 3px 9px rgba(35,70,96,.065); font-size:.92rem; font-weight:800; white-space:nowrap;}
        .rn-bar {position:relative; min-height:28px; margin-top:5px; overflow:hidden; border-top:1px solid rgba(255,255,255,.7); border-radius:10px 10px 2px 2px; background:linear-gradient(160deg,var(--bar-top) 0%,var(--bar-mid) 46%,var(--bar-bottom) 100%); box-shadow:0 12px 20px rgba(37,76,105,.13),inset 1px 0 rgba(255,255,255,.28),inset -1px 0 rgba(19,51,75,.09);}
        .rn-bar::before {content:"";position:absolute;top:0;bottom:0;left:0;width:18%;background:linear-gradient(90deg,rgba(255,255,255,.30),rgba(255,255,255,0));}
        .rn-bar::after {content:"";position:absolute;top:0;right:8%;left:8%;height:2px;border-radius:999px;background:rgba(255,255,255,.55);}
        .rn-bar-age {position:absolute; right:0; bottom:5px; left:0; color:#70879a; font-size:.76rem; white-space:nowrap;}
        .rn-fixed-chart {position:relative; height:calc(100% - 29px);}
        .rn-fixed-block {position:absolute; left:0; bottom:0; display:flex; align-items:center; justify-content:center; width:var(--w); min-width:24%; height:var(--h); min-height:52px; padding:.45rem; overflow:hidden; color:#173451; border-top:1px solid rgba(255,255,255,.72); border-radius:10px 10px 2px 2px; background:linear-gradient(160deg,#f4dfa9 0%,#d9b460 48%,#b7802e 100%); text-align:center; box-shadow:0 12px 20px rgba(135,94,29,.15),inset 1px 0 rgba(255,255,255,.32),inset -1px 0 rgba(93,58,14,.10);}
        .rn-coverage {position:absolute; top:calc(100% - var(--h)); right:0; left:var(--w); padding:.55rem .3rem 0; color:#4f687c; border-top:2px solid #d9e2e8; font-size:.90rem; font-weight:650; line-height:1.45; text-align:center;}
        .rn-coverage strong {color:#38566e; font-weight:800;}
        .rn-axis {position:relative; height:2rem; margin-top:.28rem; color:#71889c; font-size:.76rem;}
        .rn-axis span {position:absolute; top:0; white-space:nowrap;}
        .rn-axis-start {left:0;}.rn-axis-end {right:0;}.rn-axis-complete {left:var(--w); transform:translateX(-50%);}
        .rn-panel-copy {min-height:3rem; color:#71889c; font-size:.8rem; line-height:1.55;}
        .rn-note {margin-top:.75rem; color:#71889c; font-size:.79rem; line-height:1.55;}
        @media (max-width:760px) {.rn-chart-grid,.rn-metric-grid{grid-template-columns:1fr}.rn-chart-zone{height:220px}}
        @media print {
          .block-container{max-width:none!important;padding:.3cm!important;}
          [data-testid="stHeader"],[data-testid="stToolbar"],[data-testid="stSidebar"],[data-testid="stStatusWidget"],.stButton,.st-key-rn_guide,.st-key-rn_input_area{display:none!important;}
          .rn-result-hero{margin-top:.35rem!important;break-inside:avoid;}
          .rn-metric-grid,.rn-chart-grid,.rn-panel,.rn-metric{break-inside:avoid;}
          [data-testid="stExpander"] details:not([open])>:not(summary){display:block!important;}
          .rn-note,.rn-panel-copy,.rn-axis,.rn-bar-age{font-size:10pt!important;}
          .rn-bar-value{font-size:11pt!important;}
          .rn-coverage{font-size:10.5pt!important;}
          .rn-result-copy{font-size:15pt!important}.rn-result-value{font-size:21pt!important}
        }
        </style>
        """
        ).strip(),
        unsafe_allow_html=True,
    )


def _render_chart(
    periods: list[RenewalPeriod],
    current_age: int,
    end_age: int,
    fixed_premium: float,
    fixed_years: int,
) -> None:
    if len(periods) > 5:
        indexes = sorted({0, 1, len(periods) // 2, len(periods) - 2, len(periods) - 1})
        shown = [periods[index] for index in indexes]
    else:
        shown = periods

    highest = max([fixed_premium, 1] + [period.monthly_premium for period in periods])
    bars = []
    premium_palette = [
        ("#b8d0e1", "#83a9c5", "#5f89a8"),
        ("#a6c5da", "#719dbd", "#497a9f"),
        ("#91b7d1", "#5e8eaf", "#386d94"),
        ("#7ca8c6", "#4f819f", "#2d617f"),
        ("#6899bb", "#407493", "#234f6d"),
    ]
    for index, period in enumerate(shown):
        height = 20 + (period.monthly_premium / highest) * 55
        palette = premium_palette[min(index, len(premium_palette) - 1)]
        bars.append(
            dedent(
                f"""
                <div class="rn-column">
                  <div class="rn-bar-value">{html.escape(_compact_won(period.monthly_premium))}</div>
                  <div class="rn-bar" style="height:{height:.1f}%;--bar-top:{palette[0]};--bar-mid:{palette[1]};--bar-bottom:{palette[2]}"></div>
                  <div class="rn-bar-age">{period.start_age}세</div>
                </div>
                """
            ).strip()
        )

    coverage_years = max(1, end_age - current_age)
    pay_years = min(fixed_years, coverage_years)
    fixed_end_age = current_age + pay_years
    pay_width = min(100, max(24, pay_years / coverage_years * 100))
    fixed_height = 20 + (fixed_premium / highest) * 55
    end_label = "" if fixed_end_age >= end_age else f'<span class="rn-axis-end">보장 종료 {end_age}세</span>'

    st.markdown(
        dedent(
            f"""
            <div class="rn-chart-grid">
              <section class="rn-panel">
                <div class="rn-panel-title"><span class="rn-dot"></span>갱신형 보험료 변화</div>
                <div class="rn-chart-zone"><div class="rn-bars">{''.join(bars)}</div></div>
                <div class="rn-axis"><span class="rn-axis-start">현재 {current_age}세</span><span class="rn-axis-end">{end_age}세까지 납입 예상</span></div>
                <div class="rn-panel-copy">갱신 시점마다 보험료가 변동되며 보장기간 동안 보험료 납입이 계속됩니다.</div>
              </section>
              <section class="rn-panel">
                <div class="rn-panel-title"><span class="rn-dot fixed"></span>비갱신형 보험료와 보장기간</div>
                <div class="rn-chart-zone">
                  <div class="rn-fixed-chart" style="--w:{pay_width:.1f}%;--h:{fixed_height:.1f}%">
                    <div class="rn-fixed-block"><span><strong>월 {_won(fixed_premium)}</strong><br>{fixed_end_age}세 납입 완료</span></div>
                    <div class="rn-coverage"><strong>{fixed_end_age}세 이후</strong><br>{end_age}세까지 보장 유지</div>
                  </div>
                </div>
                <div class="rn-axis" style="--w:{pay_width:.1f}%"><span class="rn-axis-start">가입 {current_age}세</span><span class="rn-axis-complete">납입 완료 {fixed_end_age}세</span>{end_label}</div>
                <div class="rn-panel-copy">보험료가 일정하고 정해진 납입기간이 끝난 뒤에도 약정된 보장기간까지 보장이 유지됩니다.</div>
              </section>
            </div>
            """
        ).strip(),
        unsafe_allow_html=True,
    )


def run() -> None:
    _inject_style()
    page_header(
        "고객 상담",
        "갱신형 vs 비갱신형",
        "현재 부담과 앞으로 예상되는 총보험료를 함께 비교합니다.",
        "RN",
    )

    with st.container(key="rn_guide"):
        with st.expander("📘 사용방법 및 인쇄 안내", expanded=False):
            guide_use, guide_print = st.columns(2, gap="large")
            with guide_use:
                st.markdown("#### 사용방법")
                st.markdown(
                    dedent(
                        """
                        1. **기본 정보**에서 현재 나이, 보장 종료 나이와 은퇴 예상 연령을 입력합니다.
                        2. **갱신형 정보**에서 현재 보험료, 갱신주기와 다음 갱신까지 남은 기간을 입력합니다.
                        3. 간편하게 비교할 때는 **간편 시나리오**를 사용합니다.
                        4. 가입제안서에 갱신보험료가 기재되어 있다면 **가입제안서 직접 입력**을 선택합니다.
                        5. 비교하려는 비갱신형 보험료와 납입기간을 입력하면 결과가 자동으로 계산됩니다.
                        """
                    ).strip()
                )
            with guide_print:
                st.markdown("#### 인쇄방법")
                st.markdown(
                    dedent(
                        """
                        1. 모든 정보를 입력하고 결과를 확인합니다.
                        2. 브라우저에서 **Ctrl + P**를 누릅니다.
                        3. 용지는 **A4**, 방향은 **세로**로 설정합니다.
                        4. 배율은 우선 **페이지에 맞춤**을 선택합니다.
                        5. 그래프 색상을 출력하려면 **배경 그래픽**을 활성화합니다.
                        6. 입력 영역은 제외되고 **현재 시점의 핵심 비교**부터 인쇄됩니다.
                        7. 갱신 시점별 상세 보험료 표는 인쇄 시 항상 펼쳐져 표시됩니다.
                        """
                    ).strip()
                )
            st.caption("브라우저와 프린터에 따라 결과가 달라질 수 있습니다. 내용이 잘리는 경우 배율을 90~95%로 조정해 주세요.")

    with st.container(key="rn_input_area"):
        basic_col, renew_col, fixed_col = st.columns(3, gap="medium")

        with basic_col:
            with st.container(border=True):
                st.markdown('<div class="rn-card-label"><span class="rn-card-number">1</span>기본 정보</div>', unsafe_allow_html=True)
                current_age = int(st.number_input("현재 나이", min_value=18, max_value=90, value=40, step=1))
                end_age = int(st.number_input("보장 종료 나이", min_value=current_age + 1, max_value=110, value=max(100, current_age + 1), step=1))
                retirement_age = int(st.number_input("예상 은퇴 나이", min_value=current_age, max_value=end_age, value=min(max(65, current_age), end_age), step=1))
                past_paid = float(st.number_input("현재까지 납입한 총액 · 선택", min_value=0, value=0, step=1_000_000, help="모르는 경우 0원으로 두세요."))

        with renew_col:
            with st.container(border=True):
                st.markdown('<div class="rn-card-label"><span class="rn-card-number">2</span>갱신형 정보</div>', unsafe_allow_html=True)
                current_premium = float(st.number_input("현재 월보험료", min_value=0, value=80_000, step=10_000))
                cycle = int(st.selectbox("갱신 주기", [5, 10, 15, 20, 30], index=3, format_func=lambda value: f"{value}년"))
                next_years = int(st.number_input("다음 갱신까지 남은 기간", min_value=1, max_value=min(cycle, end_age - current_age), value=min(5, cycle, end_age - current_age), step=1))
                method = st.selectbox("갱신보험료 산정 방식", ["간편 시나리오", "가입제안서 직접 입력", "갱신배수 직접 설정"])
                scenario = st.selectbox("갱신 상승 시나리오", list(SCENARIOS), index=1, disabled=method != "간편 시나리오")

        with fixed_col:
            with st.container(border=True):
                st.markdown('<div class="rn-card-label"><span class="rn-card-number">3</span>비갱신형 정보</div>', unsafe_allow_html=True)
                fixed_premium = float(st.number_input("월보험료", min_value=0, value=135_000, step=10_000, key="rn_fixed_premium"))
                fixed_years = int(st.selectbox("납입기간", [10, 15, 20, 25, 30], index=2, format_func=lambda value: f"{value}년"))

        ages = _renewal_ages(current_age, end_age, cycle, next_years)
        custom_values: list[float] = []
        if method != "간편 시나리오" and ages:
            title = "가입제안서의 갱신 시점별 월보험료" if method == "가입제안서 직접 입력" else "갱신 시점별 적용 배수"
            with st.expander(title, expanded=True):
                input_columns = st.columns(min(4, len(ages)))
                preview_premium = current_premium
                reference = REFERENCE_MULTIPLIERS[cycle]
                for index, age in enumerate(ages):
                    base_multiple = reference[min(index, len(reference) - 1)]
                    preview_premium *= base_multiple
                    with input_columns[index % len(input_columns)]:
                        if method == "가입제안서 직접 입력":
                            result = st.number_input(
                                f"{age}세 월보험료",
                                min_value=0,
                                value=int(round(preview_premium / 1_000) * 1_000),
                                step=1_000,
                                key=f"rn_proposal_{cycle}_{age}",
                            )
                        else:
                            result = st.number_input(
                                f"{age}세 갱신배수",
                                min_value=0.0,
                                value=float(base_multiple),
                                step=0.01,
                                format="%.4f",
                                key=f"rn_multiple_{cycle}_{age}",
                            )
                        custom_values.append(float(result))

    periods = _calculate_periods(
        current_age=current_age,
        end_age=end_age,
        cycle=cycle,
        next_years=next_years,
        current_premium=current_premium,
        method=method,
        scenario=scenario,
        custom_values=custom_values,
    )

    renew_future = sum(period.total for period in periods)
    fixed_pay_years = min(fixed_years, end_age - current_age)
    fixed_total = fixed_premium * fixed_pay_years * 12
    saving = renew_future - fixed_total
    highest = max((period.monthly_premium for period in periods), default=0)
    fixed_end_age = current_age + fixed_pay_years

    if saving >= 0:
        result_text = f'비갱신형 전환 시 예상 보험료를 <span class="rn-result-value">약 {_won(saving)} 절감</span>할 수 있습니다.'
    else:
        result_text = f'현재 조건에서는 비갱신형 전환안이 <span class="rn-result-value negative">약 {_won(abs(saving))} 더 높습니다.</span>'

    source = "상담 예상" if method == "간편 시나리오" else ("가입제안서 입력" if method == "가입제안서 직접 입력" else "사용자 설정")
    st.markdown(
        dedent(
            f"""
            <div class="rn-result-hero">
              <div class="rn-result-kicker">현재 시점의 핵심 비교</div>
              <div class="rn-result-copy">{result_text}</div>
            </div>
            <div class="rn-metric-grid">
              <div class="rn-metric"><div class="rn-metric-label">갱신형 미래 예상보험료 · {source}</div><div class="rn-metric-value">{_won(renew_future)}</div><div class="rn-metric-sub">최고 예상 월보험료 {_won(highest)}</div></div>
              <div class="rn-metric"><div class="rn-metric-label">비갱신형 총보험료 · 제안 조건</div><div class="rn-metric-value">{_won(fixed_total)}</div><div class="rn-metric-sub">{fixed_end_age}세 납입 완료</div></div>
            </div>
            """
        ).strip(),
        unsafe_allow_html=True,
    )

    _render_chart(periods, current_age, end_age, fixed_premium, fixed_years)

    with st.expander("갱신 시점별 상세 보험료 보기", expanded=True):
        detail_rows = []
        for index, period in enumerate(periods):
            detail_rows.append(
                {
                    "구간": "현재" if index == 0 else f"{index}차 갱신 후",
                    "연령": f"{period.start_age}~{period.end_age}세",
                    "월보험료": _won(period.monthly_premium),
                    "현재 대비": f"{period.cumulative_multiple:.2f}배",
                    "구간 총액": _won(period.total),
                }
            )
        st.table(pd.DataFrame(detail_rows).set_index("구간"))

    retire_renew = sum(
        period.monthly_premium
        * max(0, period.end_age - max(period.start_age, retirement_age))
        * 12
        for period in periods
    )
    fixed_retire_years = max(0, fixed_end_age - max(current_age, retirement_age))
    retire_fixed = fixed_premium * fixed_retire_years * 12

    comparison_rows = [
        {"구분": "현재 이후 납입", "갱신형 유지": _won(renew_future), "비갱신형 전환": _won(fixed_total)},
        {"구분": f"{retirement_age}세 은퇴 이후 예상 납입", "갱신형 유지": _won(retire_renew), "비갱신형 전환": _won(retire_fixed)},
    ]
    if past_paid > 0:
        comparison_rows.insert(0, {"구분": "현재까지 납입", "갱신형 유지": _won(past_paid), "비갱신형 전환": _won(past_paid)})
        comparison_rows.append({"구분": "가입부터 예상 총액", "갱신형 유지": _won(past_paid + renew_future), "비갱신형 전환": _won(past_paid + fixed_total)})

    st.markdown('<div class="rn-section-title">예상 보험료 비교</div>', unsafe_allow_html=True)
    st.table(pd.DataFrame(comparison_rows).set_index("구분"))

    if method == "간편 시나리오":
        note = "보험업계 평균의 갱신배수 흐름을 참고해 재설계한 상담용 가정입니다."
    elif method == "가입제안서 직접 입력":
        note = "입력한 가입제안서의 보험료 예시를 기준으로 계산했습니다. 실제 갱신보험료는 갱신 시점의 위험률과 손해율 등에 따라 달라질 수 있습니다."
    else:
        note = "보험업계 평균의 갱신배수 흐름을 참고해 재설계한 상담용 가정입니다."
    st.markdown(f'<div class="rn-note">{html.escape(note)}</div>', unsafe_allow_html=True)
