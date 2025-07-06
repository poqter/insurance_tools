import streamlit as st
import pandas as pd

def run():

    # --- 페이지 설정 ---
    st.set_page_config(page_title="질병 위험률 분석 도구(테스트)", layout="wide")

    # --- 사이드바 정보 ---
    st.sidebar.markdown("---")
    st.sidebar.markdown("""
    👨‍💻 제작자: 드림지점 박병선 팀장  
    🗓️ 버전: v1.0.0  
    📅 최종 업데이트: 2025-07-06
    """)

    # --- 데이터 불러오기 ---
    df_risk = pd.read_csv("disease_risk.csv")
    df_adjust = pd.read_csv("disease_adjust.csv")
    df_treat = pd.read_csv("disease_treatment.csv")
    df_coverage = pd.read_csv("disease_coverage.csv")

    # --- 메인 타이틀 ---
    st.title("🧬 질병 위험률 분석 도구(테스트)")
    st.markdown("""
    #### 고객님의 연령대, 성별, 건강상태를 바탕으로 위험률을 분석합니다.
    """)

    # --- 입력 카드 UI ---
    with st.container():
        st.subheader("📥 상담 정보 입력")
        col1, col2 = st.columns(2)
        with col1:
            age_group = st.selectbox("연령대", options=sorted(df_risk["연령대"].unique()), key="age")
            gender = st.selectbox("성별", options=sorted(df_risk["성별"].unique()), key="gender")
            smoke = st.selectbox("흡연 여부", options=sorted(df_risk["흡연여부"].unique()), key="smoke")
            drink = st.selectbox("음주 여부", options=sorted(df_risk["음주여부"].unique()), key="drink")
        with col2:
            family = st.selectbox("가족력", options=sorted(df_risk["가족력"].unique()), key="family")
            disease_options = df_risk["기저질환"].str.split("+").explode().unique()
            conditions = st.multiselect("보유 질병", options=sorted(disease_options), key="conditions")
            job = st.selectbox("직업", options=["사무직", "육체노동직", "학생", "자영업", "무직"], key="job")
            exercise = st.selectbox("운동 습관", options=["규칙적으로 운동함", "가끔 운동함", "거의 안 함"], key="exercise")

        if st.button("🔄 입력값 초기화"):
            st.session_state.clear()
            st.experimental_rerun()

        run_analysis = st.button("📊 결과 확인하기")

    # --- 결과 분석 ---
    if run_analysis:
        st.subheader("📊 분석 결과")
        category_map = {
            "암": "암",
            "뇌": "뇌혈관질환",
            "심장": "심장질환"
        }

        base_risk_dict = {"암": 377, "뇌": 24, "심장": 16.9}
        max_risk_dict = {"암": 583, "뇌": 238, "심장": 427}

        factor_inputs = {
            "연령대": [age_group],
            "기저질환": conditions,
            "흡연여부": [smoke],
            "음주여부": [drink],
            "가족력": [family],
            "직업": [job],
            "운동 습관": [exercise]
        }

        for cat, disease_name in category_map.items():
            adjust_factors = []
            for kind, values in factor_inputs.items():
                for value in values:
                    row = df_adjust[(df_adjust["질병군"] == cat) & (df_adjust["항목종류"] == kind) & (df_adjust["항목명"] == value)]
                    if not row.empty:
                        adjust_factors.append(row["조정계수"].values[0])

            weights = [1.0] * len(adjust_factors)
            weighted_sum = sum(a * w for a, w in zip(adjust_factors, weights))
            final_adjust = round(weighted_sum / sum(weights), 2) if weights else 1.0
            base_risk = base_risk_dict[cat]
            final_risk = round(min(base_risk * final_adjust, max_risk_dict[cat]), 1)

            treat_info = df_treat[df_treat["질병"] == disease_name]
            coverage_info = df_coverage[df_coverage["질병"] == disease_name]

            d_rate = coverage_info['진단비보유율(%)'].values[0] if not coverage_info.empty else '-'
            t_rate = coverage_info['치료비보유율(%)'].values[0] if not coverage_info.empty else '-'
            s_cost = treat_info['수술비용(만원)'].values[0] if not treat_info.empty else '-'
            r_days = treat_info['회복기간(일)'].values[0] if not treat_info.empty else '-'
            t_cost = treat_info['평균치료비용(만원)'].values[0] if "평균치료비용(만원)" in treat_info.columns else '-'

            st.markdown(f"""
    🔹 **{cat} 위험 - {disease_name}**
    - 기본 위험률: 1000명 중 **{base_risk}명**
    - 보정 위험률 (개인조건 반영): **{final_risk}명** (보정 계수 평균: {final_adjust})
    - 진단비 보유율: {d_rate}% / 치료비 보유율: {t_rate}%
    - 평균 수술비용: {s_cost}만원 / 평균 회복기간: {r_days}일
    {'- 평균 치료비용: ' + str(t_cost) + '만원' if t_cost != '-' else ''}
    """)
            st.markdown("---")
