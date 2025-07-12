import streamlit as st
import pandas as pd
import numpy as np

"""
질병 위험률 분석 도구 v1.2.1  
📌 **KeyError(질병군) 해결 버전**  
─────  
1. CSV **컬럼 자동 매핑**(질병군↔질병, 위험률↔위험률(명/1000))으로 유연성 강화  
2. **데이터 유효성 검사** 화면 표시 → 누락 컬럼·행 즉시 안내  
3. 핵심 로직·UI 동일 (v1.2)
"""

# -------------------------------------------------
# 기본 설정
# -------------------------------------------------
st.set_page_config(page_title="질병 위험률 분석 도구 v1.2.1", layout="wide")

EXPECTED_COLS_RISK = {
    "질병군": ["질병", "질병구분"],
    "연령대": [],
    "성별": [],
    "위험률(명/1000)": ["위험률", "risk_per_1000"],
    "기저질환": [],
}

EXPECTED_COLS_ADJ = {
    "질병군": ["질병"],
    "항목종류": [],
    "항목명": [],
    "조정계수": ["계수"],
    "가중치": [],  # optional
}

@st.cache_data
def load_and_validate(path: str, exp_map: dict, name: str):
    """CSV 불러오고 컬럼 매핑·검증"""
    df = pd.read_csv(path)
    rename_dict = {}
    for req, alts in exp_map.items():
        if req not in df.columns:
            for alt in alts:
                if alt in df.columns:
                    rename_dict[alt] = req
                    break
    df = df.rename(columns=rename_dict)
    missing = [c for c in exp_map if c not in df.columns]
    if missing:
        st.error(f"❌ {name} CSV에 필요한 컬럼이 없습니다: {', '.join(missing)}")
    return df, missing

@st.cache_data
def load_all():
    df_risk, miss_risk = load_and_validate("disease_risk.csv", EXPECTED_COLS_RISK, "disease_risk")
    df_adj,  miss_adj  = load_and_validate("disease_adjust.csv", EXPECTED_COLS_ADJ,  "disease_adjust")
    df_treat = pd.read_csv("disease_treatment.csv")
    df_cov   = pd.read_csv("disease_coverage.csv")
    return df_risk, df_adj, df_treat, df_cov, miss_risk + miss_adj

DF_RISK, DF_ADJ, DF_TREAT, DF_COV, MISSING_COLS = load_all()

if MISSING_COLS:
    st.stop()  # 필수 컬럼 누락 시 앱 중단

# -------------------------------------------------
# 사이드바 : 정보 & 초기화
# -------------------------------------------------
with st.sidebar:
    st.markdown("### ℹ️ 사용 가이드")
    st.markdown(
        """
        1. 왼쪽 카드에 **고객 정보** 입력 후 **[결과 분석하기]** 클릭  
        2. 위험률·권장 진단비 확인, 하단 CSV 저장  
        """
    )
    if st.button("🔄 모든 입력 초기화", use_container_width=True):
        st.session_state.clear()
        st.experimental_rerun()

    st.markdown("---")
    st.markdown("👨‍💻 **제작자**: 드림지점 박병선 팀장  ")
    st.markdown("🗓️ **버전**: v1.2.1  ")
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
            DF_RISK.get("기저질환", pd.Series()).astype(str).str.split("+").explode().dropna().unique().tolist()
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
        surgery_cost   = float(treat["수술비용(만원)"].values[0]) if (not treat.empty and "수술비용(만원)" in treat.columns) else np.nan
        median_cost    = np.nanmedian([avg_treat_cost, surgery_cost]) if (not np.isnan(avg_treat_cost) or not np.isnan(surgery_cost)) else 0

        diag_rate  = cov.get("진단비보유율(%)", pd.Series(["-"])).values[0]
        treat_rate = cov.get("치료비보유율(%)", pd.Series(["-"])).values[0]

        # 권장 진단비 산출
        recommended_diag = int(max(median_cost * 2, 2000))
        recommended_diag = int(round(recommended_diag / 1000) * 1000)

        results.append({
            "질병군": cat,
            "기본 위험률": base_risk,
            "보정 위험률": final_risk,
            "진단비 보유율(%)": diag_rate,
            "권장 진단비(만원)": recommended_diag,
        })
        details_dict[cat] = factor_logs

    if not results:
        st.stop()

    df_result = pd.DataFrame(results).set_index("질병군")

    st.subheader("🔎 위험률 & 권장 진단비 요약")
    st.dataframe(df_result, use_container_width=True)

    st.subheader("📊 위험률 비교 차트")
    st.bar_chart(df_result[["기본 위험률", "보정 위험률"]])

    st.download_button(
        "💾 CSV 다운로드",
        data=df_result.to_csv().encode("utf-8-sig"),
        file_name="risk_report_v1_2_1.csv",
        mime="text/csv",
    )

    with st.expander("📖 보정 계수 상세 보기"):
        for cat, logs in details_dict.items():
            st.markdown(f"#### {cat}")
            for kind, val, coef, weight in logs:
                st.markdown(f"- **{kind}**: {val} → 계수 {coef} (가중치 {weight})")
