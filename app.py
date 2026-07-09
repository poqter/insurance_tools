import streamlit as st

from modules import (
    deposit_vs_shortpay,
    renewal_vs_nonrenewal,
    analyzer,
    remodeling,
    convention,
    summer,
    manager_results,
    pdf_converter,
)

# 페이지 설정
st.set_page_config(page_title="보험컨설팅 멀티 도우미", layout="wide")


# -----------------------------
# 🔔 로그인 후 공지 팝업
# -----------------------------
@st.dialog("📌 보험컨설팅 멀티 도우미 공지사항")
def show_login_notice_popup():
    login_user = st.session_state.get("login_user", "사용자")

    st.markdown(f"""
    ### {login_user}님, 로그인되었습니다.

    보험컨설팅 멀티 도우미를 사용하기 전 아래 내용을 확인해주세요.

    ---

    #### ✅ 사용 전 안내사항

    1. **엑셀 파일은 지정된 양식에 맞춰 업로드해주세요.**  
       열 이름이나 순서가 다르면 일부 계산 결과가 다르게 나올 수 있습니다.

    2. **계산 결과는 최종 제출 전 반드시 한 번 더 확인해주세요.**  
       이 도구는 상담과 계산을 돕기 위한 보조 도구입니다.

    3. **썸머 계산기 사용 시 수금자와 보너스율을 정확히 선택해주세요.**  
       보너스 반영 후 금액을 기준으로 등급이 판정됩니다.

    4. **PDF 표 엑셀 변환기는 원본 PDF 상태에 따라 결과가 달라질 수 있습니다.**  
       변환 후 엑셀 내용을 확인하고 사용해주세요.

    5. **고객 정보가 포함된 파일은 외부에 공유하지 않도록 주의해주세요.**

    ---
    """)

    st.info("확인 버튼을 누르면 메인 화면으로 이동합니다.")

    if st.button("확인했습니다", type="primary", use_container_width=True):
        st.session_state["notice_confirmed"] = True
        st.rerun()


# -----------------------------
# 🔐 여러 비밀번호 인증 + 사용자 구분
# -----------------------------
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if "login_user" not in st.session_state:
        st.session_state["login_user"] = None

    if "notice_confirmed" not in st.session_state:
        st.session_state["notice_confirmed"] = False

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

            # 로그인 성공 후 공지 팝업을 다시 띄우기 위한 초기화
            st.session_state["notice_confirmed"] = False

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
# 🔔 로그인 후 공지 팝업 표시
# -----------------------------
if not st.session_state.get("notice_confirmed", False):
    show_login_notice_popup()


# -----------------------------
# 🧩 전체 앱 목록
# -----------------------------
all_apps = {
    "📑 보장 분석 도우미": analyzer.run,
    "💰 적금 vs 단기납 비교": deposit_vs_shortpay.run,
    "📊 갱신 vs 비갱신 보험 비교": renewal_vs_nonrenewal.run,
    "🔁 보험 리모델링 전/후 비교": remodeling.run,
    "🧮 컨벤션 계산기": convention.run,
    "🌞 썸머 계산기": summer.run,
    "📊 매니저 업적 환산": manager_results.run,
    "📄 PDF 표 엑셀 변환기": pdf_converter.run,
}


# -----------------------------
# 🔑 사용자별 메뉴 권한 설정
# -----------------------------
user_permissions = {
    # team1용: 전체 기능 사용 가능
    "team1": [
        "📑 보장 분석 도우미",
        "💰 적금 vs 단기납 비교",
        "📊 갱신 vs 비갱신 보험 비교",
        "🔁 보험 리모델링 전/후 비교",
        "🧮 컨벤션 계산기",
        "🌞 썸머 계산기",
        "📄 PDF 표 엑셀 변환기",
        "📊 매니저 업적 환산",
    ],

    # team2용: 일부 기능 사용 가능
    "team2": [
        "📑 보장 분석 도우미",
        "💰 적금 vs 단기납 비교",
        "📊 갱신 vs 비갱신 보험 비교",
        # "🔁 보험 리모델링 전/후 비교",
        # "🧮 컨벤션 계산기",
        # "🌞 썸머 계산기",
        # "📄 PDF 표 엑셀 변환기",
        # "📊 매니저 업적 환산",
    ],

    # team3용: 일부 기능 사용 가능
    "team3": [
        "📑 보장 분석 도우미",
        "💰 적금 vs 단기납 비교",
        "📊 갱신 vs 비갱신 보험 비교",
        # "🔁 보험 리모델링 전/후 비교",
        "🧮 컨벤션 계산기",
        "🌞 썸머 계산기",
        # "📄 PDF 표 엑셀 변환기",
        "📊 매니저 업적 환산",
    ],

    # team4용: 일부 기능 사용 가능
    "team4": [
        "📑 보장 분석 도우미",
        "💰 적금 vs 단기납 비교",
        # "📊 갱신 vs 비갱신 보험 비교",
        # "🔁 보험 리모델링 전/후 비교",
        # "🧮 컨벤션 계산기",
        # "🌞 썸머 계산기",
        # "📄 PDF 표 엑셀 변환기",
        # "📊 매니저 업적 환산",
    ],

    # team5용: 일부 기능 사용 가능
    "team5": [
        "📑 보장 분석 도우미",
        "💰 적금 vs 단기납 비교",
        # "📊 갱신 vs 비갱신 보험 비교",
        # "🔁 보험 리모델링 전/후 비교",
        # "🧮 컨벤션 계산기",
        # "🌞 썸머 계산기",
        # "📄 PDF 표 엑셀 변환기",
        # "📊 매니저 업적 환산",
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