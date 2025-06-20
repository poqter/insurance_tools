import streamlit as st
from modules import deposit_vs_shortpay, renewal_vs_nonrenewal, analyzer

st.set_page_config(page_title="보험 멀티 도우미", layout="wide")
st.title("🔧 보험 컨설팅 멀티 도우미")

menu = st.radio("원하는 기능을 선택하세요 👇", [
    "💰 적금 vs 단기납 비교",
    "📊 갱신 vs 비갱신 보험 비교",
    "📑 보장 분석 도우미"
])

if menu == "💰 적금 vs 단기납 비교":
    deposit_vs_shortpay.run()
elif menu == "📊 갱신 vs 비갱신 보험 비교":
    renewal_vs_nonrenewal.run()
elif menu == "📑 보장 분석 도우미":
    analyzer.run()
