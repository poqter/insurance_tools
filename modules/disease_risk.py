import streamlit as st
import pandas as pd
import numpy as np

def run():
    """
    질병 위험률 분석 도구 v1.3 ✨  
    🔸 **차트·CSV·확장 로그 제거 → 설득형 메시지 출력**  
    🔸 입력 초기화 버튼·사이드바 가이드 최소화  
    🔸 위험 수준(낮음·주의·위험·고위험)별 **맞춤 멘트** 자동 생성  
    """

    # -------------------------------------------------
    # 기본 설정
    # -------------------------------------------------
    st.set_page_config(page_title="질병 위험률 분석 도구 v1.3", layout="wide")

    @st.cache_data
    def load_data():
        df_risk     = pd.read_csv("disease_risk.csv")
        df_adjust   = pd.read_csv("disease_adjust.csv")
        df_treat    = pd.read_csv("disease_treatment.csv")
        df_coverage = pd.read_csv("disease_coverage.csv")
        return df_risk, df_adjust, df_treat, df_coverage

    DF_RISK, DF_ADJ, DF_TREAT, DF_COV = load_data()

    # -------------------------------------------------
    # 간단 버전 정보 (사이드바)
    # -------------------------------------------------
    with st.sidebar:
        st.markdown("👨‍💻 **제작**: 박병선 팀장  ")
        st.markdown("🗓️ **버전**: v1.3 (2025‑07‑12)")

    # -------------------------------------------------
    # 입력 UI
    # -------------------------------------------------
    st.title("🧬 3대 질병 맞춤 위험 메시지 리포트")
    with st.container(border=True):
        st.subheader("📥 고객 정보 입력")
        col1, col2 = st.columns(2)
        with col1:
            age_group = st.selectbox("연령대", sorted(DF_RISK["연령대"].unique()))
            gender    = st.selectbox("성별",     sorted(DF_RISK["성별"].unique()))
            smoke     = st.selectbox("흡연 여부", ["비흡연", "흡연"])
            drink     = st.selectbox("음주 습관", ["가벼움/없음", "과음"])
        with col2:
            family     = st.selectbox("가족력", ["없음", "있음"])
            disease_pool = (
                DF_RISK["기저질환"].astype(str).str.split("+").explode().dropna().unique().tolist()
            )
            conditions = st.multiselect("보유 질병", sorted(disease_pool))
            job        = st.selectbox("직업", ["사무직", "육체노동직", "학생", "자영업", "무직"])
            exercise   = st.selectbox("운동 습관", ["규칙적으로 운동", "가끔 운동", "거의 안함"])

    # -------------------------------------------------
    # 분석 로직
    # -------------------------------------------------
    if st.button("🔍 맞춤 메시지 확인", type="primary"):
        category_map = {"암": "암", "뇌": "뇌혈관질환", "심장": "심장질환"}
        factor_inputs = {
            "연령대": [age_group],
            "성별":   [gender],
            "흡연여부": [smoke],
            "음주여부": [drink],
            "가족력": [family],
            "기저질환": conditions if conditions else ["없음"],
            "직업": [job],
            "운동 습관": [exercise],
        }

        # 위험 등급 기준 (명/1000)
        risk_thresholds = {
            "암":   [200, 400, 600],   # 낮음 ≤200 < 주의 ≤400 < 위험 ≤600 < 고위험
            "뇌":   [10, 30, 60],
            "심장": [8, 25, 50],
        }

        # 메시지 템플릿
        templates = {
            "낮음": "현재 조건을 반영한 **{name} 위험률**은 1000명 중 약 **{risk}명** 수준입니다. 평균보다 낮은 편이지만, 혹시 모를 치료비에 대비해 최소 {diag}만원은 준비해 두면 안심할 수 있습니다.",
            "주의": "{name} 위험률이 1000명 중 **{risk}명**으로 평균을 넘어섰습니다. 치료비만 평균 {cost}만원이 필요한 만큼, 진단비 {diag}만원 이상은 꼭 확보해 두시는 걸 권장드립니다.",
            "위험": "현재 {name} 위험률이 **{risk}명/1000**로 **높은 단계**에 해당합니다. 치료·회복기간이 길어 평균 {cost}만원의 비용이 듭니다. 진단비 {diag}만원 조차 준비되지 않았다면, 가족 재정에 큰 부담이 될 수 있습니다.",
            "고위험": "⚠️ **{name} 고위험군**입니다 — 1000명 중 무려 **{risk}명** 수준. 치료비 {cost}만원·회복 {days}일 이상이 예상되며, 지금 진단비 보유율은 고작 {rate}%. 최소 {diag}만원 이상 보장을 서둘러 준비하시길 강력히 권합니다.",
        }

        st.header("🔔 맞춤 설득 메시지")

        for cat, disease_name in category_map.items():
            # 1) 기본 위험률
            base_row = DF_RISK[(DF_RISK["질병군"] == cat) & (DF_RISK["연령대"] == age_group) & (DF_RISK["성별"] == gender)]
            if base_row.empty:
                continue
            base_risk = float(base_row["위험률(명/1000)"]).mean()

            # 2) 보정 계수 누적 곱
            adj_mult = 1.0
            for kind, vals in factor_inputs.items():
                for val in vals:
                    row = DF_ADJ[(DF_ADJ["질병군"] == cat) & (DF_ADJ["항목종류"] == kind) & (DF_ADJ["항목명"] == val)]
                    if not row.empty:
                        coef = float(row["조정계수"].values[0])
                        adj_mult *= coef
            final_risk = round(base_risk * adj_mult, 1)

            # 3) 기준 진단비·치료비
            treat = DF_TREAT[DF_TREAT["질병"] == disease_name]
            median_cost = float(treat["평균치료비용(만원)"].values[0]) if not treat.empty else 0
            recovery_days = int(treat["회복기간(일)"].values[0]) if not treat.empty else "-"
            recommended_diag = int(max(median_cost * 2, 2000))
            recommended_diag = int(round(recommended_diag / 1000) * 1000)

            cov = DF_COV[DF_COV["질병"] == disease_name]
            diag_rate = cov["진단비보유율(%)"].values[0] if not cov.empty else "-"

            # 4) 위험 레벨 판정
            t1, t2, t3 = risk_thresholds[cat]
            if final_risk <= t1:
                level = "낮음"
            elif final_risk <= t2:
                level = "주의"
            elif final_risk <= t3:
                level = "위험"
            else:
                level = "고위험"

            msg = templates[level].format(
                name=disease_name,
                risk=final_risk,
                diag=recommended_diag,
                cost=median_cost,
                days=recovery_days,
                rate=diag_rate,
            )

            st.markdown(f"### ▶️ {disease_name}")
            st.markdown(msg)
            st.markdown("---")