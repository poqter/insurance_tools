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

# -----------------------------
# 🔐 여러 비밀번호 인증
# -----------------------------
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    st.title("🔐 보험컨설팅 멀티 도우미")
    st.caption("접근 권한 확인을 위해 비밀번호를 입력해주세요.")

    password = st.text_input("비밀번호", type="password")

    if st.button("로그인"):
        valid_passwords = list(st.secrets["passwords"].values())

        if password in valid_passwords:
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")

    return False


if not check_password():
    st.stop()

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