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
from modules.ui_components import inject_global_styles


st.set_page_config(
    page_title="화랑사업부 업무 도우미",
    page_icon="🧰",
    layout="wide",
    initial_sidebar_state="expanded",
)

inject_global_styles()


# 공지는 이 목록만 수정하면 로그인 화면에 반영됩니다.
NOTICE = {
    "date": "2026.08.01",
    "title": "업무 도우미 화면이 새롭게 개편되었습니다.",
    "items": [
        "로그인 후 통합 홈에서 필요한 업무를 선택할 수 있습니다.",
        "고객 상담과 실적 관리 메뉴가 업무 목적별로 구분되었습니다.",
        "썸머·컨벤션·상속세 계산기가 추가되었습니다.",
        "비밀번호 입력 후 Enter 키를 눌러 로그인할 수 있습니다.",
    ],
    "important": "8월1일부터 비밀번호가 변경되었습니다. 변경된 비밀번호는 박병선 팀장에게 문의해 주세요.",
}


APP_DEFINITIONS = {
    "analyzer": {
        "name": "보장 분석 도우미",
        "icon": "📑",
        "code": "BA",
        "category": "고객 상담",
        "description": "보험사 보장분석 자료를 고객용 양식으로 변환합니다.",
        "action": "보장 분석 시작",
        "run": analyzer.run,
    },
    "remodeling": {
        "name": "보험 리모델링",
        "icon": "🔁",
        "code": "RM",
        "category": "고객 상담",
        "description": "변경안을 비교하고 고객용 엑셀 자료를 만듭니다.",
        "action": "리모델링 시작",
        "status": "개선 중",
        "run": remodeling.run,
    },
    "deposit_vs_shortpay": {
        "name": "적금 vs 단기납",
        "icon": "💰",
        "code": "DS",
        "category": "고객 상담",
        "description": "10년 기준 적금과 단기납의 예상 결과를 비교합니다.",
        "action": "비교 계산 시작",
        "run": deposit_vs_shortpay.run,
    },
    "renewal_vs_nonrenewal": {
        "name": "갱신 vs 비갱신",
        "icon": "📊",
        "code": "RN",
        "category": "고객 상담",
        "description": "보험료 변동을 반영해 장기 총납입액을 비교합니다.",
        "action": "보험료 비교 시작",
        "run": renewal_vs_nonrenewal.run,
    },
    "inheritance_tax": {
        "name": "상속세 계산기",
        "icon": "🧾",
        "code": "IT",
        "category": "고객 상담",
        "description": "예상 상속세와 부족한 현금성 납부재원을 계산합니다.",
        "action": "상속세 계산 시작",
        "run": inheritance_tax.run,
    },
    "convention": {
        "name": "컨벤션 계산기",
        "icon": "🏆",
        "code": "CV",
        "category": "실적 관리",
        "description": "계약 실적을 환산하고 컨벤션 달성 여부를 확인합니다.",
        "action": "컨벤션 계산 시작",
        "run": convention.run,
    },
    "summer": {
        "name": "썸머 계산기",
        "icon": "🌞",
        "code": "SU",
        "category": "실적 관리",
        "description": "7·8월 업적을 반영해 썸머 업적을 계산합니다.",
        "action": "썸머 실적 계산",
        "run": summer.run,
    },
    "manager_results": {
        "name": "매니저 업적 환산",
        "icon": "📈",
        "code": "MR",
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

    st.markdown(
        """
        <div class="hw-login-brand"><span class="hw-logo">H</span><strong>화랑 <b>WORKS</b></strong></div>
        <div class="hw-login-hero">
          <span>HWARANG BUSINESS WORKSPACE</span>
          <h1>보험 업무의 복잡함을,<br><em>더 간단하게.</em></h1>
          <p>상담자료 제작부터 실적 관리까지 필요한 업무를 한곳에서 이용하세요.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

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
        st.markdown('<div class="hw-side-brand"><span>H</span><strong>화랑 WORKS</strong></div>', unsafe_allow_html=True)
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
        status = f'<span class="hw-card-badge">{app["status"]}</span>' if app.get("status") else ""
        clean_name = app["name"]
        lock_text = '<span class="hw-card-lock">권한 제한</span>' if not is_allowed else ""
        st.markdown(
            f"""
            <div class="hw-tool-card-head"><span class="hw-tool-icon">{app.get('code', app['icon'])}</span><span class="hw-tool-category">{app['category']}{status}{lock_text}</span></div>
            <div class="hw-tool-title">{clean_name}</div>
            <div class="hw-tool-desc">{app['description']}</div>
            """,
            unsafe_allow_html=True,
        )
        button_label = "시작하기  →" if is_allowed else "🔒  사용 권한 없음"
        if st.button(
            button_label,
            key=f"home_{app_id}",
            disabled=not is_allowed,
            use_container_width=True,
        ):
            navigate(app_id)


def render_home(allowed_ids: list[str]) -> None:
    user = st.session_state["login_user"]
    permissions = USER_PERMISSIONS.get(user, {})
    available_count = sum(bool(value) for value in permissions.values())
    locked_count = len(APP_DEFINITIONS) - available_count

    st.markdown(
        """
        <style>
        .hw-home-hero { position:relative; overflow:hidden; display:grid; grid-template-columns:1.08fr .92fr; gap:3rem;
            align-items:center; margin:-.5rem 0 1.7rem; padding:3.3rem 3.5rem; border:1px solid #DFE9F1; border-radius:24px;
            background:radial-gradient(circle at 78% 30%,rgba(23,105,220,.13),transparent 27%),radial-gradient(circle at 90% 77%,rgba(17,155,152,.12),transparent 25%),linear-gradient(120deg,#FFFFFF,#F5F9FF); box-shadow:0 20px 55px rgba(30,70,100,.08); }
        .hw-home-kicker { color:#3F7197; font-size:.72rem; font-weight:850; letter-spacing:.13em; }
        .hw-home-copy h1 { margin:.9rem 0 1rem; color:#10283D; font-size:clamp(2.65rem,4.5vw,4.5rem); line-height:1.13; letter-spacing:-.065em; }
        .hw-home-copy h1 em { color:#1769DC; font-style:normal; }
        .hw-home-copy p { margin:0; color:#5E7284; font-size:1.03rem; line-height:1.75; }
        .hw-home-dashboard { background:rgba(255,255,255,.94); border:1px solid #DFE9F1; border-radius:18px; padding:1.35rem; box-shadow:0 18px 42px rgba(35,76,108,.13); }
        .hw-dash-title { color:#6C8091; font-size:.7rem; margin-bottom:.8rem; }
        .hw-dash-title strong { display:block; color:#10283D; font-size:1rem; margin-top:.2rem; }
        .hw-dash-grid { display:grid; grid-template-columns:repeat(3,1fr); gap:.6rem; }
        .hw-dash-stat { padding:.9rem .8rem; border:1px solid #E5EDF3; border-radius:12px; color:#748696; font-size:.67rem; }
        .hw-dash-stat strong { display:block; color:#10283D; font-size:1.35rem; margin-top:.25rem; }
        .hw-dash-stat b { color:#119B98; }
        .hw-announcement { display:flex; align-items:center; gap:.75rem; margin:1rem 0 3rem; padding:1rem 1.15rem; background:#FFFFFF; border:1px solid #D9E6EF; border-radius:13px; color:#10283D; box-shadow:0 5px 18px rgba(36,73,103,.04); }
        .hw-announcement span { display:grid; place-items:center; width:2rem; height:2rem; border-radius:.65rem; background:#EAF3FF; color:#1769DC; font-size:.75rem; }
        .hw-announcement strong { font-size:.88rem; }
        .hw-announcement small { color:#8293A1; }
        .hw-category-head { margin:2.6rem 0 1.15rem; }
        .hw-category-head span { color:#1769DC; font-size:.68rem; font-weight:850; letter-spacing:.12em; }
        .hw-category-head h2 { margin:.35rem 0 .3rem !important; font-size:1.7rem; }
        .hw-category-head p { margin:0; color:#647789; font-size:.86rem; }
        [class*="st-key-locked_card_"] {
            background-color: #F2F5F7 !important; opacity:.72;
            border: 1px dashed #B8C7D2 !important; border-radius: 1rem !important;
        }
        [class*="st-key-available_card_"] { min-height:15.2rem; background:#FFFFFF; border:1px solid #DCE6EE !important; border-radius:1rem !important; box-shadow:0 8px 24px rgba(27,64,93,.045); transition:transform .2s ease,box-shadow .2s ease; }
        [class*="st-key-available_card_"]:hover { transform:translateY(-4px); box-shadow:0 16px 36px rgba(27,64,93,.11); }
        [class*="st-key-available_card_"] button { background:#FFFFFF !important; color:#1769DC !important; border-color:#C8D9E7 !important; }
        [class*="st-key-available_card_"] button:hover { background:#1769DC !important; color:#FFFFFF !important; border-color:#1769DC !important; }
        .hw-tool-card-head { display:flex; align-items:center; justify-content:space-between; margin-bottom:1.15rem; }
        .hw-tool-icon { width:3rem; height:3rem; display:grid; place-items:center; background:linear-gradient(145deg,#EAF3FF,#E9F8F7); color:#1769DC; border-radius:.85rem; font-size:.75rem; font-weight:900; }
        .hw-tool-category { color:#718697; font-size:.66rem; font-weight:750; }
        .hw-card-badge,.hw-card-lock { margin-left:.4rem; padding:.2rem .38rem; border-radius:999px; background:#EAF3FF; color:#1769DC; font-size:.58rem; }
        .hw-card-lock { background:#E5EAEE; color:#697A87; }
        .hw-tool-title { color:#10283D; font-size:1.12rem; font-weight:780; letter-spacing:-.035em; margin-bottom:.45rem; }
        .hw-tool-desc { color:#647789; font-size:.79rem; line-height:1.65; min-height:2.65rem; margin-bottom:.8rem; }
        @media(max-width:850px){.hw-home-hero{grid-template-columns:1fr;padding:2.4rem 2rem}.hw-home-dashboard{display:none}}
        @media(max-width:560px){.hw-home-hero{padding:2rem 1.35rem}.hw-home-copy h1{font-size:2.45rem}.hw-dash-grid{grid-template-columns:1fr}.hw-announcement small{display:none}}
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""
        <div class="hw-home-hero">
          <div class="hw-home-copy"><span class="hw-home-kicker">HWARANG BUSINESS WORKSPACE</span><h1>보험 업무의 복잡함을,<br><em>더 간단하게.</em></h1><p>상담자료 제작부터 실적 관리까지<br>필요한 업무를 한곳에서 이용하세요.</p></div>
          <div class="hw-home-dashboard"><div class="hw-dash-title">오늘의 업무 환경<strong>좋은 하루입니다, {user}님</strong></div><div class="hw-dash-grid"><div class="hw-dash-stat">사용 가능 도구<strong>{available_count}개</strong></div><div class="hw-dash-stat">권한 제한<strong>{locked_count}개</strong></div><div class="hw-dash-stat">시스템 상태<strong><b>정상</b></strong></div></div></div>
        </div>
        <div class="hw-announcement"><span>●</span><strong>{NOTICE['title']}</strong><small>{NOTICE['date']}</small></div>
        """,
        unsafe_allow_html=True,
    )

    for category in ("고객 상담", "실적 관리"):
        category_apps = [
            app_id
            for app_id in APP_DEFINITIONS
            if APP_DEFINITIONS[app_id]["category"] == category
        ]
        if not category_apps:
            continue

        description = "고객 설명과 상담자료 제작에 필요한 도구입니다." if category == "고객 상담" else "개인·조직 실적과 행사 달성 현황을 확인합니다."
        st.markdown(f'<div class="hw-category-head"><span>WORK SOLUTIONS</span><h2>{category} 솔루션</h2><p>{description}</p></div>', unsafe_allow_html=True)
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
