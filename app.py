import streamlit as st

from modules import (
    deposit_vs_shortpay,
    renewal_vs_nonrenewal,
    analyzer,
    remodeling,
    # convention,
)

# 페이지 설정
st.set_page_config(page_title="보험컨설팅 멀티 도우미", layout="wide")

# 👉 사이드바 메뉴로 앱 선택 이동
st.sidebar.title("🧰 보험컨설팅 멀티 도우미")
app_option = st.sidebar.radio("📌 사용할 기능을 선택하세요:", [
    "📑 보장 분석 도우미",
    "💰 적금 vs 단기납 비교",
    #"📊 갱신 vs 비갱신 보험 비교",
    #"🔁 보험 리모델링 전/후 비교",
    #"🧮 컨벤션 계산기"
])

# 🧠 선택된 앱 실행
if app_option == "📑 보장 분석 도우미":
    analyzer.run()
elif app_option == "💰 적금 vs 단기납 비교":
    deposit_vs_shortpay.run()
#elif app_option == "📊 갱신 vs 비갱신 보험 비교":
#    renewal_vs_nonrenewal.run()
#elif app_option == "🔁 보험 리모델링 전/후 비교":
#    remodeling.run()
#elif app_option == "🧮 컨벤션 계산기":
#    convention.run()    