import streamlit as st

from modules import (
    analyzer,
    convention,
    deposit_vs_shortpay,
    inheritance_tax,
    manager_results,
    remodeling,
    renewal_vs_nonrenewal,
    summer,
)


st.set_page_config(
    page_title="화랑사업부 업무 도우미",
    page_icon="🧰",
    layout="wide",
    initial_sidebar_state="expanded",
)


# 공지는 이 목록만 수정하면 로그인 화면에 반영됩니다.
NOTICE = {
    "date": "2026.08.01",
    "title": "업무 도우미 화면이 개편.",
    "items": [
        "로그인 후 통합 홈에서 필요한 업무를 선택할 수 있습니다.",
        "고객 상담과 실적 관리 메뉴가 업무 목적별로 구분되었습니다.",
        "썸머·컨벤션·상속세 계산기가 추가되었습니다.",
        "비밀번호 입력 후 Enter 키를 눌러 로그인할 수 있습니다.",
    ],
    "important": "비밀번호가 변경되었습니다. 변경된 비밀번호는 박병선 팀장에게 문의해 주세요.",
}


APP_DEFINITIONS = {
    "analyzer": {
        "name": "보장 분석 도우미",
        "icon": "📑",
        "category": "고객 상담",
        "description": "보험사 보장분석 자료를 고객용 양식으로 변환합니다.",
        "action": "보장 분석 시작",
        "run": analyzer.run,
    },
    "remodeling": {
        "name": "보험 리모델링(수정중)",
        "icon": "🔁",
        "category": "고객 상담",
        "description": "변경안을 비교하고 고객용 엑셀 자료를 만듭니다.",
        "action": "리모델링 시작",
        "run": remodeling.run,
    },
    "deposit_vs_shortpay": {
        "name": "적금 vs 단기납",
        "icon": "💰",
        "category": "고객 상담",
        "description": "10년 기준 적금과 단기납의 예상 결과를 비교합니다.",
        "action": "비교 계산 시작",
        "run": deposit_vs_shortpay.run,
    },
    "renewal_vs_nonrenewal": {
        "name": "갱신 vs 비갱신",
        "icon": "📊",
        "category": "고객 상담",
        "description": "보험료 변동을 반영해 장기 총납입액을 비교합니다.",
        "action": "보험료 비교 시작",
        "run": renewal_vs_nonrenewal.run,
    },
    "inheritance_tax": {
        "name": "상속세 계산기",
        "icon": "🧾",
        "category": "고객 상담",
        "description": "예상 상속세와 부족한 현금성 납부재원을 계산합니다.",
        "action": "상속세 계산 시작",
        "run": inheritance_tax.run,
    },
    "convention": {
        "name": "컨벤션 계산기",
        "icon": "🏆",
        "category": "실적 관리",
        "description": "계약 실적을 환산하고 컨벤션 달성 여부를 확인합니다.",
        "action": "컨벤션 계산 시작",
        "run": convention.run,
    },
    "summer": {
        "name": "썸머 계산기",
        "icon": "🌞",
        "category": "실적 관리",
        "description": "7·8월 업적을 반영해 썸머 업적을 계산합니다.",
        "action": "썸머 실적 계산",
        "run": summer.run,
    },
    "manager_results": {
        "name": "매니저 업적 환산",
        "icon": "📈",
        "category": "실적 관리",
        "description": "지점 실적 환산금액을 집계합니다.",
        "action": "매니저 실적 확인",
        "run": manager_results.run,
    },
}


# --------------------------------------------------
# 계정별 기능 권한
# True  = 사용 가능
# False = 홈에서 잠금 표시, 사이드바에서 숨김
# --------------------------------------------------
USER_PERMISSIONS = {
    "Admin": {
        "analyzer": True,                  # 보장 분석 도우미
        "remodeling": True,                # 보험 리모델링
        "deposit_vs_shortpay": True,       # 적금 vs 단기납
        "renewal_vs_nonrenewal": True,     # 갱신 vs 비갱신
        "inheritance_tax": True,           # 상속세 계산기
        "convention": True,                # 컨벤션 계산기
        "summer": True,                    # 썸머 계산기
        "manager_results": True,           # 매니저 업적 환산
    },
    "Manager1": {
        "analyzer": True,                  # 보장 분석 도우미
        "remodeling": True,                # 보험 리모델링
        "deposit_vs_shortpay": True,       # 적금 vs 단기납
        "renewal_vs_nonrenewal": True,     # 갱신 vs 비갱신
        "inheritance_tax": True,           # 상속세 계산기
        "convention": True,                # 컨벤션 계산기
        "summer": True,                    # 썸머 계산기
        "manager_results": True,           # 매니저 업적 환산
    },
    "Basic": {
        "analyzer": True,                  # 보장 분석 도우미
        "remodeling": False,               # 보험 리모델링
        "deposit_vs_shortpay": False,       # 적금 vs 단기납
        "renewal_vs_nonrenewal": False,     # 갱신 vs 비갱신
        "inheritance_tax": False,          # 상속세 계산기
        "convention": True,                # 컨벤션 계산기
        "summer": True,                    # 썸머 계산기
        "manager_results": False,           # 매니저 업적 환산
    },
    "Crew": {
        "analyzer": True,                  # 보장 분석 도우미
        "remodeling": False,               # 보험 리모델링
        "deposit_vs_shortpay": True,       # 적금 vs 단기납
        "renewal_vs_nonrenewal": True,     # 갱신 vs 비갱신
        "inheritance_tax": False,          # 상속세 계산기
        "convention": True,                # 컨벤션 계산기
        "summer": True,                    # 썸머 계산기
        "manager_results": False,          # 매니저 업적 환산
    },
    "Dream": {
        "analyzer": True,                  # 보장 분석 도우미
        "remodeling": True,                # 보험 리모델링
        "deposit_vs_shortpay": True,       # 적금 vs 단기납
        "renewal_vs_nonrenewal": True,     # 갱신 vs 비갱신
        "inheritance_tax": True,           # 상속세 계산기
        "convention": True,                # 컨벤션 계산기
        "summer": True,                    # 썸머 계산기
        "manager_results": False,          # 매니저 업적 환산
    },
}


def initialize_state() -> None:
    st.session_state.setdefault("password_correct", False)
    st.session_state.setdefault("login_user", None)
    st.session_state.setdefault("active_app", "home")


def render_notice() -> None:
    st.markdown("### 공지사항")
    st.caption(f"최근 업데이트 · {NOTICE['date']}")
    with st.container(border=True):
        st.markdown(f"**{NOTICE['title']}**")
        for item in NOTICE["items"]:
            st.markdown(f"- {item}")
    st.info(NOTICE["important"], icon="ℹ️")


def render_login() -> bool:
    if st.session_state["password_correct"]:
        return True

    st.markdown("# 화랑사업부 업무 도우미")
    st.caption("보험 상담자료 제작과 실적 관리에 필요한 기능을 한곳에서 이용하세요.")
    st.divider()

    login_col, notice_col = st.columns([1, 1.15], gap="large")
    with login_col:
        st.markdown("### 로그인")
        st.write("발급받은 비밀번호를 입력해 주세요.")
        with st.form("login_form", clear_on_submit=False):
            password = st.text_input("비밀번호", type="password", placeholder="비밀번호 입력")
            submitted = st.form_submit_button("로그인", type="primary", use_container_width=True)

        if submitted:
            passwords = dict(st.secrets["passwords"])
            matched_user = next(
                (user_name for user_name, saved_password in passwords.items() if password == saved_password),
                None,
            )
            if matched_user:
                st.session_state["password_correct"] = True
                st.session_state["login_user"] = matched_user
                st.session_state["active_app"] = "home"
                st.rerun()
            else:
                st.error("입력한 비밀번호를 확인해 주세요.")

    with notice_col:
        render_notice()

    return False


def allowed_app_ids() -> list[str]:
    user = st.session_state.get("login_user")
    permissions = USER_PERMISSIONS.get(user, {})
    return [
        app_id
        for app_id in APP_DEFINITIONS
        if permissions.get(app_id, False)
    ]


def navigate(app_id: str) -> None:
    st.session_state["active_app"] = app_id
    st.rerun()


def logout() -> None:
    # 고객 입력값과 프로그램별 임시 상태까지 모두 삭제합니다.
    st.session_state.clear()
    st.rerun()


def render_sidebar(allowed_ids: list[str]) -> None:
    with st.sidebar:
        st.title("🧰 업무 도우미")
        st.caption("필요한 업무를 선택하세요.")

        home_active = st.session_state["active_app"] == "home"
        if st.button(
            "🏠  홈",
            key="nav_home",
            type="primary" if home_active else "secondary",
            use_container_width=True,
        ):
            navigate("home")

        for category in ("고객 상담", "실적 관리"):
            category_apps = [
                app_id
                for app_id in allowed_ids
                if APP_DEFINITIONS[app_id]["category"] == category
            ]
            if not category_apps:
                continue

            st.markdown(f"#### {category}")
            for app_id in category_apps:
                app = APP_DEFINITIONS[app_id]
                active = st.session_state["active_app"] == app_id
                if st.button(
                    f"{app['icon']}  {app['name']}",
                    key=f"nav_{app_id}",
                    type="primary" if active else "secondary",
                    use_container_width=True,
                ):
                    navigate(app_id)

        st.divider()
        st.caption(f"접속 계정 · {st.session_state['login_user']}")
        with st.expander("최근 공지"):
            st.caption(NOTICE["date"])
            st.markdown(f"**{NOTICE['title']}**")
        if st.button("🚪로그아웃", key="logout", use_container_width=True):
            logout()


def render_app_card(app_id: str, is_allowed: bool) -> None:
    app = APP_DEFINITIONS[app_id]
    card_key = f"available_card_{app_id}" if is_allowed else f"locked_card_{app_id}"
    with st.container(border=True, key=card_key):
        st.markdown(f"### {app['icon']} {app['name']}")
        if app.get("status"):
            st.caption(f"🛠️ {app['status']}")
        st.write(app["description"])
        button_label = app["action"] if is_allowed else "🔒 사용 권한 없음"
        if st.button(
            button_label,
            key=f"home_{app_id}",
            disabled=not is_allowed,
            use_container_width=True,
        ):
            navigate(app_id)


def render_home(allowed_ids: list[str]) -> None:
    user = st.session_state["login_user"]
    st.title("화랑사업부 업무 도우미")
    st.caption(f"{user} 계정의 업무 도구입니다. 잠금 표시된 기능은 현재 계정에서 사용할 수 없습니다.")

    # 잠긴 카드에만 연한 회색 배경과 점선 테두리를 적용합니다.
    st.markdown(
        """
        <style>
        [class*="st-key-locked_card_"] {
            background-color: #F5F6F7 !important;
            border: 1px dashed #B8BEC5 !important;
            border-radius: 0.5rem !important;
        }

        [class*="st-key-available_card_"] button {
            background-color: #52758A !important;
            color: #FFFFFF !important;
            border: 1px solid #52758A !important;
        }

        [class*="st-key-available_card_"] button:hover {
            background-color: #466779 !important;
            color: #FFFFFF !important;
            border-color: #466779 !important;
        }

        [class*="st-key-available_card_"] button:active {
            background-color: #3D5B6B !important;
            color: #FFFFFF !important;
            border-color: #3D5B6B !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    with st.container(border=True):
        st.markdown("**상담자료 제작부터 실적 환산까지 필요한 업무를 빠르게 시작하세요.**")
        st.caption("왼쪽 메뉴 또는 아래 업무 카드를 선택하면 해당 프로그램으로 이동합니다.")

    for category in ("고객 상담", "실적 관리"):
        category_apps = [
            app_id
            for app_id in APP_DEFINITIONS
            if APP_DEFINITIONS[app_id]["category"] == category
        ]
        if not category_apps:
            continue

        st.markdown(f"## {category}")
        for start in range(0, len(category_apps), 3):
            row_apps = category_apps[start : start + 3]
            columns = st.columns(3, gap="medium")
            for column, app_id in zip(columns, row_apps):
                with column:
                    render_app_card(app_id, app_id in allowed_ids)

    st.divider()
    st.caption("제작 · 박병선 팀장")


def main() -> None:
    initialize_state()
    if not render_login():
        st.stop()

    allowed_ids = allowed_app_ids()
    active_app = st.session_state.get("active_app", "home")
    if active_app != "home" and active_app not in allowed_ids:
        st.session_state["active_app"] = "home"
        active_app = "home"

    render_sidebar(allowed_ids)

    if active_app == "home":
        render_home(allowed_ids)
    else:
        APP_DEFINITIONS[active_app]["run"]()


if __name__ == "__main__":
    main()
