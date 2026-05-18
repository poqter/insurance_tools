
import streamlit as st
import time

def run():
    # 강조 박스 함수
    def emphasize_box(text, bg="#e6f2ff", color="#003366"):
        return f"""<div style='background-color:{bg}; color:{color}; padding:12px; border-radius:10px;
                    font-size:20px; font-weight:bold; margin-bottom:10px;'>{text}</div>"""

    # 금액 포맷 함수 (만원 이하 삭제용)
    def format_currency_trim(value):
        won = int(value * 10000)
        if won % 10000 == 0:
            return f"{won // 10000:,}만원"
        else:
            return f"{won:,}원"

    # 사이드바 인쇄 안내
    with st.sidebar:
        st.markdown("""
        📄 **인쇄 안내**

        🖨️ 오른쪽 위 ... 버튼 → print를 누르면 인쇄하거나 PDF로 저장할 수 있어요.

        🔧 **설정 더 보기**에서:
        - 머리글과 바닥글 ❌ 체크 해제
        - 배경 그래픽 ✅ 체크

        🔍 배율 조정으로 페이지에 알맞게 설정

        🚫 인쇄 시에는 이 안내 페이지 닫기.             
        """)
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("""
        <div style='margin-bottom:20px;'>👨‍💻 <strong>제작자:</strong> 비전본부 드림지점 박병선 팀장</div>
        <div style='margin-bottom:20px;'>🗓️ <strong>버전:</strong> v1.0.0</div>
        <div style='margin-bottom:20px;'>📅 <strong>최종 업데이트:</strong> 2026-05-17</div>
        """, unsafe_allow_html=True)

    # 제목 링크 아이콘 숨기기
    st.markdown("""
    <style>
    h1 a, h2 a, h3 a { display: none !important; }
    </style>
    """, unsafe_allow_html=True)

    st.title("💰 적금 vs 단기납 비교")

    col1, col2 = st.columns(2)

    with col1:
        st.header("📌 적금")
        deposit_monthly = st.number_input("월 납입액 (만원)", min_value=0, step=1, value=None, format="%d", placeholder="예: 100")
        deposit_rate = st.number_input("연 이자율 (%)", min_value=0.0, step=0.1, value=None, placeholder="예: 2.5")

    with col2:
        st.header("📌 단기납")
        insurance_monthly = st.number_input("월 납입액 (만원)", min_value=0, step=1, value=None, format="%d", placeholder="예: 100", key="ins_monthly")
        pay_years = st.selectbox("납입 기간 (년)", [5, 7], index=0)
        return_rate = st.number_input("10년 시점 해지환환급률 (%)", min_value=0.0, step=0.1, value=None, placeholder="예: 123.0")

    if st.button("결과 보기"):
        if deposit_monthly in (None, 0) or deposit_rate in (None, 0.0) or insurance_monthly in (None, 0) or return_rate in (None, 0.0):
            st.warning("⚠️ 모든 항목에 값을 입력해주세요.")
        else:
            with st.spinner("결과를 계산 중입니다..."):
                time.sleep(0.5)

            st.markdown("---")
            st.subheader("🔍 결과 분석")

            monthly_rate = (deposit_rate / 100) / 12
            total_deposit = deposit_monthly * 12
            interest_sum = sum([deposit_monthly * monthly_rate * (12 - m) for m in range(12)])
            pre_tax_interest = interest_sum
            tax = pre_tax_interest * 0.154
            after_tax_interest = pre_tax_interest - tax
            monthly_avg_interest = after_tax_interest / 12
            total_after_tax_interest_10y = after_tax_interest * 10

            total_insurance = insurance_monthly * 12 * pay_years
            refund = total_insurance * (return_rate / 100)
            bonus = refund - total_insurance
            monthly_bonus = bonus / 120

            sum1, sum2 = st.columns(2)
            with sum1:
                st.markdown("### 📜 적금 계산 요약")
                st.write(f"- 원금 합계 (1년): {format_currency_trim(total_deposit)}")
                st.write(f"- 세전 이자: {format_currency_trim(pre_tax_interest)}")
                st.write(f"- 이자 과세 (15.4%): {format_currency_trim(tax)}")
                st.write(f"- 세후 이자: {format_currency_trim(after_tax_interest)}")

            with sum2:
                st.markdown("### 📜 단기납 계산 요약")
                st.write(f"- 원금 합계 ({pay_years}년): {format_currency_trim(total_insurance)}")
                st.write(f"- 10년 시점 해지환급금: {format_currency_trim(refund)}")
                st.write(f"- 단기납 보너스 금액: {format_currency_trim(bonus)}")
                st.write(f"- 10년 이후 해지 시, **비과세 혜택** 적용 가능")

            st.markdown("### ✅ 핵심 요약")
            colm1, colm2 = st.columns(2)
            with colm1:
                st.metric("세후 이자 총합 (10년 기준)", f"{int(total_after_tax_interest_10y // 1)}만원")
                st.markdown(emphasize_box(f"세후 이자 월 평균: {monthly_avg_interest * 10000:,.0f}원", bg="#e6f2ff", color="#003366"), unsafe_allow_html=True)
            with colm2:
                st.metric("단기납 보너스 총합 (10년 기준)", f"{int(bonus // 1)}만원", delta=f"{bonus - total_after_tax_interest_10y:,.0f}만원")
                st.markdown(emphasize_box(f"단기납 보너스 월 평균: {monthly_bonus * 10000:,.0f}원", bg="#fff3e6", color="#663300"), unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("### 📌 참고 계산")

            if deposit_rate > 0:
                monthly_rate = (deposit_rate / 100) / 12
                factor = sum([(12 - m) * monthly_rate for m in range(12)])
                monthly_required = (bonus / 10) / (factor * (1 - 0.154))
                st.markdown(f"""
                <div style='font-size:18px; margin-top:8px; margin-bottom:6px;'>
                    👉 단기납 보너스 총합과 같으려면, 적금 월 납입액을 <span style='color:red; font-weight:bold;'>{monthly_required:,.0f}만원</span>으로 변경해야 합니다.
                </div>
                """, unsafe_allow_html=True)

            if deposit_monthly > 0:
                r_monthly = (bonus / 10) / (deposit_monthly * 78 * (1 - 0.154))
                deposit_rate_needed = r_monthly * 12 * 100
                st.markdown(f"""
                <div style='font-size:18px; margin-top:4px; margin-bottom:8px;'>
                    👉 현재 적금 월 납입액으로 단기납 보너스 총합과 같아지려면, 연 이자율이 <span style='color:red; font-weight:bold;'>{deposit_rate_needed:,.2f}%</span>여야 합니다.
                </div>
                """, unsafe_allow_html=True)

            st.markdown("""
            <style>
            @media print {
                html, body {
                    margin: 0;
                    padding: 0;
                    height: auto !important;
                    overflow: visible !important;
                }
                .block-container {
                    padding-bottom: 0 !important;
                    margin-bottom: 0 !important;
                }
                main:after {
                    content: none !important;
                }
                .no-print {
                    display: none;
                }
            }
            h1 a, h2 a, h3 a { display: none !important; }
            </style>
            """, unsafe_allow_html=True)
