import streamlit as st

from modules import (
    deposit_vs_shortpay,
    renewal_vs_nonrenewal,
    analyzer,
    remodeling,
    convention,
)

# 페이지 설정
st.set_page_config(page_title="보험컨설팅 멀티 도우미", layout="wide")


# -----------------------------
# 🔐 여러 비밀번호 인증 + 사용자 구분
# -----------------------------
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if "login_user" not in st.session_state:
        st.session_state["login_user"] = None

    if st.session_state["password_correct"]:
        return True

    st.title("🔐 보험컨설팅 멀티 도우미")
    st.caption("접근 권한 확인을 위해 비밀번호를 입력해주세요.")

    password = st.text_input("비밀번호", type="password")

    if st.button("로그인"):
        passwords = dict(st.secrets["passwords"])

        matched_user = None

        for user_name, saved_password in passwords.items():
            if password == saved_password:
                matched_user = user_name
                break

        if matched_user:
            st.session_state["password_correct"] = True
            st.session_state["login_user"] = matched_user
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")

    return False


# 비밀번호가 틀리면 여기서 앱 실행 중단
if not check_password():
    st.stop()


# -----------------------------
# 👤 현재 로그인 사용자
# -----------------------------
login_user = st.session_state.get("login_user")


# -----------------------------
# 🧩 전체 앱 목록
# -----------------------------
all_apps = {
    "📑 보장 분석 도우미": analyzer.run,
    "💰 적금 vs 단기납 비교": deposit_vs_shortpay.run,
    "📊 갱신 vs 비갱신 보험 비교": renewal_vs_nonrenewal.run,
    "🔁 보험 리모델링 전/후 비교": remodeling.run,
    "🧮 컨벤션 계산기": convention.run,
}


# -----------------------------
# 🔑 사용자별 메뉴 권한 설정
# -----------------------------
user_permissions = {
    # team1용: 전체 기능 사용 가능
    "team1": [
        "📑 보장 분석 도우미",
        "💰 적금 vs 단기납 비교",
        # "📊 갱신 vs 비갱신 보험 비교",
        # "🔁 보험 리모델링 전/후 비교",
        # "🧮 컨벤션 계산기",
    ],

    # team2용: 전체 기능 사용 가능
    "team2": [
        "📑 보장 분석 도우미",
        "💰 적금 vs 단기납 비교",
        # "📊 갱신 vs 비갱신 보험 비교",
        # "🔁 보험 리모델링 전/후 비교",
        # "🧮 컨벤션 계산기",
    ],

    # team3용: 전체 기능 사용 가능
    "team3": [
        "📑 보장 분석 도우미",
        "💰 적금 vs 단기납 비교",
        # "📊 갱신 vs 비갱신 보험 비교",
        # "🔁 보험 리모델링 전/후 비교",
        # "🧮 컨벤션 계산기",
    ],

     # team4용: 전체 기능 사용 가능
    "team4": [
        "📑 보장 분석 도우미",
        "💰 적금 vs 단기납 비교",
        # "📊 갱신 vs 비갱신 보험 비교",
        # "🔁 보험 리모델링 전/후 비교",
        # "🧮 컨벤션 계산기",
    ],

    # team5용: 전체 기능 사용 가능
    "team5": [
        "📑 보장 분석 도우미",
        "💰 적금 vs 단기납 비교",
        # "📊 갱신 vs 비갱신 보험 비교",
        # "🔁 보험 리모델링 전/후 비교",
        # "🧮 컨벤션 계산기",
    ],   
    
}


# -----------------------------
# 📌 현재 사용자에게 허용된 앱만 추출
# -----------------------------
allowed_app_names = user_permissions.get(login_user, [])

available_apps = {
    app_name: all_apps[app_name]
    for app_name in allowed_app_names
    if app_name in all_apps
}


# -----------------------------
# 🚫 허용된 메뉴가 없는 경우 차단
# -----------------------------
if not available_apps:
    st.error("현재 계정으로 접근 가능한 메뉴가 없습니다.")
    st.stop()


# -----------------------------
# 🧰 사이드바 메뉴
# -----------------------------
st.sidebar.title("🧰 보험컨설팅 멀티 도우미")
st.sidebar.caption(f"접속 계정: {login_user}")

app_option = st.sidebar.radio(
    "📌 사용할 기능을 선택하세요:",
    list(available_apps.keys())
)


# -----------------------------
# 🧠 선택된 앱 실행
# -----------------------------
available_apps[app_option]()