import streamlit as st
import pandas as pd
import numpy as np

"""
질병 위험률 분석 도구 v1.2  
🔹 개선 포인트
1) **계수 테이블 강화**: `disease_adjust.csv`에 `가중치` 열(선택)을 사용해 계수별 영향력 차등 적용  
2) **보험 제안 로직**: 평균 치료비를 기반으로 질병군별 **권장 진단비**를 자동 산출  
3) UI 유지, 결과표·차트·다운로드 모두 업데이트
"""

# -------------------------------------------------
# 기본 설정
# -------------------------------------------------
st.set_page_config(page_title="질병 위험률 분석 도구 v1.2", layout="wide")

@st.cache_data
def load_data():
    """CSV 데이터 불러오기 & 캐싱"""
    df_risk     = pd.read_csv("disease_risk.csv")        # 연령·성별별 기본 위험률
    df_adjust   = pd.read_csv("disease_adjust.csv")      # 위험 계수 + (선택) 가중치
    df_treat    = pd.read_csv("disease_treatment.csv")   # 평균 치료/수술 비용·회복 기간
    df_coverage = pd.read_csv("disease_coverage.csv")    # 진단비·치료비 특약 보유율
    return df_risk, df_adjust, df_treat, df_coverage

# 데이터 로드
DF_RISK, DF_ADJ, DF_TREAT, DF_COV = load_data()

# -------------------------------------------------
# 사이드바 : 정보 & 초기화
# -------------------------------------------------
with st.sidebar:
    st.markdown("### ℹ️ 사용 가이드")
    st.markdown(
        """
        1. 왼쪽 카드에 **고객 정보** 입력  
        2. **[결과 분석하기]** 클릭 → 위험률·권장 진단비 산출  
        3. 하단 **CSV 다운로드**로 리포트 저장·공유  
        """
    )
    if st.button("🔄 모든 입력 초기화", use_container_width=True):
        st.session_state.clear()
        st.experimental_rerun()

    st.markdown("---")
    st.markdown("👨‍💻 **제작자**: 드림지점 박병선 팀장  ")
    st.markdown("🗓️ **버전**: v1.2.0  ")
    st.markdown("📅 **최종 업데이트**: 2025-07-12")

# -------------------------------------------------
# 입력 카드 UI
# -------------------------------------------------
st.title("🧬 3대 질병 위험률 퍼스널 리포트")
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
            DF_RISK["기저질환"].str.split("+").explode().dropna().unique().tolist()
        )
        conditions = st.multiselect("보유 질병", sorted(disease_pool))
        job        = st.selectbox("직업", ["사무직", "육체노동직", "학생", "자영업", "무직"])
        exercise   = st.selectbox("운동 습관", ["규칙적으로 운동", "가끔 운동", "거의 안함"])

# -------------------------------------------------
# 분석 버튼
# -------------------------------------------------
if st.button("📊 결과 분석하기", type="primary"):
    category_map = {"암": "암", "뇌": "뇌혈관질환", "심장": "심장질환"}
    factor_inputs = {
        "연령대": [age_group],
        "성별":   [gender],
        "흡연여부": [smoke],
        "음주여부": [drink],
        "가족력":  [family],
        "기저질환": conditions if conditions else ["없음"],
        "직업":    [job],
        "운동 습관": [exercise],
    }

    results = []
    details_dict = {}

    for cat, disease_name in category_map.items():
        # 1) 기본 위험률(명/1000)
        base_row = DF_RISK[(DF_RISK["질병군"] == cat) &
                           (DF_RISK["연령대"] == age_group) &
                           (DF_RISK["성별"]   == gender)]
        if base_row.empty:
            st.warning(f"[데이터 없음] {cat} / {age_group} / {gender}")
            continue
        base_risk = float(base_row["위험률(명/1000)"].values[0])

        # 2) 보정 계수 곱셈 (가중치 적용)
        adjust_mult = 1.0
        factor_logs = []
        for kind, vals in factor_inputs.items():
            for val in vals:
                cond = (
                    (DF_ADJ["질병군"] == cat) &
                    (DF_ADJ["항목종류"] == kind) &
                    (DF_ADJ["항목명"]   == val)
                )
                row = DF_ADJ[cond]
                if not row.empty:
                    coef   = float(row["조정계수"].values[0])
                    weight = float(row["가중치"].values[0]) if "가중치" in row.columns else 1.0
                    adjust_mult *= coef ** weight
                    factor_logs.append((kind, val, coef, weight))

        final_risk = round(base_risk * adjust_mult, 2)

        # 3) 치료/보장 데이터
        treat = DF_TREAT[DF_TREAT["질병"] == disease_name]
        cov   = DF_COV[DF_COV["질병"] == disease_name]

        avg_treat_cost = float(treat["평균치료비용(만원)"].values[0]) if (not treat.empty and "평균치료비용(만원)" in treat.columns) else np.nan
        surgery_cost   = float(treat["수술비용(만원)"].values[0]) if not treat.empty else np.nan
        median_cost    = np.nanmedian([avg_treat_cost, surgery_cost]) if (not np.isnan(avg_treat_cost) or not np.isnan(surgery_cost)) else 0
        recovery_days  = int(treat["회복기간(일)"].values[0]) if (not treat.empty and "회복기간(일)" in treat.columns) else "-"

        diag_rate = cov["진단비보유율(%)"].values[0] if not cov.empty else "-"
        treat_rate = cov["치료비보유율(%)"].values[0] if not cov.empty else "-"

        # 4) 권장 진단비(만원): 평균 치료비×2, 최소 2,000만원, 1,000만원 단위 반올림
        recommended_diag = int(max(median_cost * 2, 2000))
        recommended_diag = int(round(recommended_diag / 1000) * 1000)

        results.append({
            "질병군": cat,
            "기본 위험률": base_risk,
            "보정 위험률": final_risk,
            "적용 계수": round(adjust_mult, 2),
            "진단비 보유율(%)": diag_rate,
            "치료비 보유율(%)": treat_rate,
            "평균 치료비용(만원)": median_cost,
            "평균 회복기간(일)": recovery_days,
            "권장 진단비(만원)": recommended_diag,
        })
        details_dict[cat] = factor_logs

    # -------------------------------
    # 결과 테이블·차트·다운로드
    # -------------------------------
    if not results:
        st.stop()

    df_result = pd.DataFrame(results).set_index("질병군")

    st.subheader("🔎 위험률 & 권장 진단비 요약")
    st.dataframe(df_result[["기본 위험률", "보정 위험률", "진단비 보유율(%)", "권장 진단비(만원)"]], use_container_width=True)

    st.subheader("📊 위험률 비교 차트")
    st.bar_chart(df_result[["기본 위험률", "보정 위험률"]])

    # CSV 다운로드
    st.download_button(
        label="💾 CSV 다운로드",
        data=df_result.to_csv().encode("utf-8-sig"),
        file_name="risk_report_v1_2.csv",
        mime="text/csv",
        help="고객 맞춤 위험률 & 권장 진단비 리포트"
    )

    # -------------------------------
    # 보정 계수 상세 로그
    # -------------------------------
    with st.expander("📖 보정 계수 상세 보기"):
        for cat, logs in details_dict.items():
            st.markdown(f"#### {cat}")
            for kind, val, coef, weight in logs:
                st.markdown(f"- **{kind}**: {val} → 계수 {coef} (가중치 {weight})")
