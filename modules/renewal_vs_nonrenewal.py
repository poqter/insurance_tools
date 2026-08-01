import streamlit as st
import pandas as pd

def run():
    st.title("📊 갱신 vs 비갱신 보험 비교")

    # 👉 입력 영역을 왼쪽과 오른쪽으로 나눔
    col_left, col_right = st.columns(2)

    with col_left:
        st.header("🌀 갱신형 보험 입력")
        start_year = st.number_input("가입 연도", min_value=1900, max_value=2100, value=None, step=1)
        start_age = st.number_input("가입 당시 나이", min_value=0, max_value=100, value=None, step=1)
        renewal_cycle = st.selectbox("갱신 주기", [10, 15, 20, 30], index=0)
        end_age = st.number_input("갱신 종료 나이", min_value=0, max_value=110, value=None, step=1)
        monthly_payment = st.number_input("현재 월 납입금액 (원)", min_value=0, value=None, step=1000)

    with col_right:
        st.header("🌱 비갱신형 보험 입력 (선택)")
        nonrenew_monthly = st.number_input("비갱신형 월 납입금액 (원)", min_value=0, value=None, step=1000)
        nonrenew_years = st.selectbox("납입기간", [10, 15, 20, 25, 30], index=2)

    # 통합 사이드바와 분리된 본문 안내 및 증가율 입력
    with st.expander("사용 가이드 및 갱신 주기별 증가율 설정"):
        st.markdown("### 📘 사용 가이드")
        st.markdown("""
        🔄 **갱신 → 비갱신 전환 비교 도구**

        📌 고객의 기존 보험 정보를 바탕으로  
        갱신형과 제안 비갱신형 상품의 총 납입금 차이를 분석합니다.

        📅 **갱신 종료 나이**는  
        현실적인 유지 가능 나이 또는 보장 만기를 기준으로 입력해주세요.

        ⚙️ 아래에서 **갱신 주기별 증가율**을 직접 조정할 수 있습니다.  
        기본값은 통계 기반으로 자동 설정됩니다.

        🖨️ 인쇄시 배율 조정으로 페이지에 알맞게 설정
        """)
        st.markdown("---")
        st.markdown("### 📈 갱신 주기별 증가율 설정")

        # 주기별 기본 증가율 매핑 (필요하면 숫자만 바꾸면 됨)
        rate_defaults = {
            10: [2.2180, 1.5239, 1.4283, 1.2212, 1.0921, 1.0624, 1.0388],  # 7개
            15: [2.7180, 1.8239, 1.6283, 1.0921, 1.0624, 1.0388],           # 6개
            20: [3.8982, 2.1253, 1.2832],                                   # 3개
            30: [5.1982, 2.1253, 1.2832]
        }

        # 현재 선택한 갱신 주기에 맞는 기본값 선택
        default_weights = rate_defaults.get(renewal_cycle, rate_defaults[10])

        # 기본값 길이에 딱 맞춰 안전하게 UI 생성
        user_weights = []
        for i, v in enumerate(default_weights, start=1):
            user_weights.append(
                st.number_input(
                    f"{i}차 갱신 증가율",
                    value=float(v),
                    step=0.01,
                    format="%.4f",
                    key=f"rate_{renewal_cycle}_{i-1}",
                )
            )

        st.markdown("---")
        st.markdown("""
        👨‍💻 **제작자**: 비전본부 드림지점 박병선 팀장  

        🗓️ **버전**: v1.0.3  

        📅 **최종 업데이트**: 2025-06-18
        """)

    # 📌 갱신형 계산 함수
    def calculate_renewal_payment(age_at_start, monthly_payment, renewal_cycle, end_age, increase_rates):
        current_age = age_at_start
        payments = []
        idx = 0

        while current_age < end_age:
            years = min(renewal_cycle, end_age - current_age)
            months = years * 12
            payment = monthly_payment
            total = payment * months

            payments.append({
                "시작나이": f"{int(current_age)}세",
                "월납입금": f"{int(round(payment)):,}원",
                "기간": f"{years}년",
                "기간 총액": f"{int(round(total)):,}원"
            })

            if idx < len(increase_rates):
                monthly_payment *= increase_rates[idx]
            else:
                monthly_payment *= increase_rates[-1]

            current_age += years
            idx += 1

        return payments

    # 📌 비갱신형 계산 함수
    def calculate_nonrenewal_payment(monthly_payment, pay_years):
        total = monthly_payment * pay_years * 12
        return {
            "납입기간": f"{pay_years}년",
            "월납입금": f"{int(round(monthly_payment)):,}원",
            "총 납입금액": f"{int(round(total)):,}원"
        }

    # ✅ 결과 보기 버튼 클릭 시 계산 수행
    if st.button("📊 결과 보기"):
        if None not in (start_year, start_age, end_age, monthly_payment):
            renewal_results = calculate_renewal_payment(start_age, monthly_payment, renewal_cycle, end_age, user_weights)
            df_renew = pd.DataFrame(renewal_results)
            df_renew.index = df_renew.index + 1

            st.subheader("🔹 갱신형 보험 납입 내역")
            st.dataframe(df_renew, use_container_width=True)

            total_renew = sum([int(r["기간 총액"].replace(",", "").replace("원", "")) for r in renewal_results])
            total_months_renew = sum([int(r["기간"].replace("년", "")) * 12 for r in renewal_results])
            avg_monthly_renew = total_renew // total_months_renew if total_months_renew > 0 else 0

            if nonrenew_monthly is not None and nonrenew_monthly > 0:
                nonrenew_result = calculate_nonrenewal_payment(nonrenew_monthly, nonrenew_years)
                df_nonrenew = pd.DataFrame([nonrenew_result])
                df_nonrenew.index = df_nonrenew.index + 1

                st.subheader("🔹 비갱신형 보험 납입 내역")
                st.dataframe(df_nonrenew, use_container_width=True)

                total_nonrenew = int(nonrenew_result["총 납입금액"].replace(",", "").replace("원", ""))
                avg_monthly_nonrenew = int(nonrenew_monthly)
                diff = total_renew - total_nonrenew

                st.markdown("### 💰 총 납입금 비교")
                col1, col2, col3 = st.columns(3)

                with col1:
                    st.markdown("**갱신형 총액**")
                    st.markdown(f"<div style='font-size:3rem; font-weight:bold; line-height:1.1'>{total_renew:,.0f} 원</div>", unsafe_allow_html=True)

                with col2:
                    st.markdown("**비갱신형 총액**")
                    st.markdown(f"<div style='font-size:3rem; font-weight:bold; line-height:1.1'>{total_nonrenew:,.0f} 원</div>", unsafe_allow_html=True)

                with col3:
                    st.markdown("**차이**")
                    color = "red" if diff > 0 else "black"
                    sign = "-" if diff > 0 else ""
                    st.markdown(f"<div style='color:{color}; font-size:3rem; font-weight:bold; line-height:1.1'>{sign}{abs(diff):,} 원</div>", unsafe_allow_html=True)

                st.markdown("<div style='margin-top: 40px;'></div>", unsafe_allow_html=True)
                st.markdown("### 📌 평균 월 납입금")
                st.markdown(f"- 갱신형 평균: **{avg_monthly_renew:,.0f} 원**")
                st.markdown(f"- 비갱신형 평균: **{avg_monthly_nonrenew:,.0f} 원**")

                if diff > 0:
                    st.markdown("<div style='margin-top: 30px;'></div>", unsafe_allow_html=True)
                    saved_years = diff // (avg_monthly_nonrenew * 12) if avg_monthly_nonrenew > 0 else 0
                    year_text = f"이 차이는 약 {saved_years}년치 보험료에 해당합니다." if saved_years > 0 else ""
                    st.success(f"✅ 추천: 비갱신형 전환 시 총 납입금이 절감되어 장기적으로 유리할 수 있습니다.\n\n{year_text}")
            else:
                st.markdown("### 💰 총 납입금")
                st.markdown(f"<div style='font-size:3rem; font-weight:bold; line-height:1.1'>{total_renew:,.0f} 원</div>", unsafe_allow_html=True)
        else:
            st.warning("❗ 갱신형 보험 입력값을 모두 입력해주세요.")
