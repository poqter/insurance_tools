"""화랑 WORKSPACE Streamlit 진입점 — 보험사 전산 포털 연결 버전."""

import streamlit as st

from modules import (
    analyzer,
    convention,
    deposit_vs_shortpay,
    inheritance_tax,
    insurer_portal,
    manager_results,
    remodeling,
    renewal_vs_nonrenewal,
    summer,
)
from modules.ui_components import inject_global_styles


st.set_page_config(
    page_title="화랑 WORKSPACE",
    page_icon="🧰",
    layout="wide",
    initial_sidebar_state="expanded",
)

inject_global_styles()


NOTICE = {
    "date": "2026.08.02",
    "title": "보험사 전산 포털이 추가되었습니다.",
    "items": [
        "생명보험 16개사와 손해보험 11개사 전산을 한 화면에서 연결합니다.",
        "삼성생명은 Microsoft Edge 브라우저에서 접속해 주세요.",
        "MG손해보험은 안내사항을 확인한 뒤 전산 페이지를 열 수 있습니다.",
        "기존 상담·실적 관리 프로그램의 계산 방식은 변경되지 않았습니다.",
    ],
    "important": "보험사별 보안 정책에 따라 로그인 또는 보안 프로그램 설치가 필요할 수 있습니다.",
}


APP_DEFINITIONS = {
    "insurer_portal": {
        "name": "보험사 전산 포털", "icon": "🌐", "category": "업무 지원",
        "description": "생명보험·손해보험 원수사 전산을 한 화면에서 연결합니다.",
        "run": insurer_portal.run,
    },
    "analyzer": {
        "name": "보장 분석 도우미", "icon": "📑", "category": "고객 상담",
        "description": "보험사 보장분석 자료를 고객용 양식으로 변환합니다.", "run": analyzer.run,
    },
    "remodeling": {
        "name": "보험 리모델링", "icon": "🔁", "category": "고객 상담",
        "description": "변경안을 비교하고 고객용 엑셀 자료를 만듭니다.", "status": "개선 중", "run": remodeling.run,
    },
    "deposit_vs_shortpay": {
        "name": "적금 vs 단기납", "icon": "💰", "category": "고객 상담",
        "description": "10년 기준 적금과 단기납의 예상 결과를 비교합니다.", "run": deposit_vs_shortpay.run,
    },
    "renewal_vs_nonrenewal": {
        "name": "갱신 vs 비갱신", "icon": "📊", "category": "고객 상담",
        "description": "보험료 변동을 반영해 장기 총납입액을 비교합니다.", "run": renewal_vs_nonrenewal.run,
    },
    "inheritance_tax": {
        "name": "상속세 계산기", "icon": "🧾", "category": "고객 상담",
        "description": "예상 상속세와 부족한 현금성 납부재원을 계산합니다.", "run": inheritance_tax.run,
    },
    "convention": {
        "name": "컨벤션 계산기", "icon": "🏆", "category": "실적 관리",
        "description": "계약 실적을 환산하고 컨벤션 달성 여부를 확인합니다.", "run": convention.run,
    },
    "summer": {
        "name": "썸머 계산기", "icon": "🌞", "category": "실적 관리",
        "description": "7·8월 업적을 반영해 썸머 업적을 계산합니다.", "run": summer.run,
    },
    "manager_results": {
        "name": "매니저 업적 환산", "icon": "📈", "category": "실적 관리",
        "description": "지점 실적 환산금액을 집계합니다.", "run": manager_results.run,
    },
}


USER_PERMISSIONS = {
    "Admin": {app_id: True for app_id in APP_DEFINITIONS},
    "Manager1": {app_id: True for app_id in APP_DEFINITIONS},
    "Basic": {
        "insurer_portal": True, "analyzer": True, "remodeling": False,
        "deposit_vs_shortpay": False, "renewal_vs_nonrenewal": False,
        "inheritance_tax": False, "convention": True, "summer": True,
        "manager_results": False,
    },
    "Crew": {
        "insurer_portal": True, "analyzer": True, "remodeling": False,
        "deposit_vs_shortpay": True, "renewal_vs_nonrenewal": True,
        "inheritance_tax": False, "convention": True, "summer": True,
        "manager_results": False,
    },
    "Dream": {
        "insurer_portal": True, "analyzer": True, "remodeling": True,
        "deposit_vs_shortpay": True, "renewal_vs_nonrenewal": True,
        "inheritance_tax": True, "convention": True, "summer": True,
        "manager_results": False,
    },
}

CATEGORY_ORDER = ("업무 지원", "고객 상담", "실적 관리")
CATEGORY_DESCRIPTIONS = {
    "업무 지원": "자주 사용하는 외부 업무 시스템을 빠르게 연결합니다.",
    "고객 상담": "고객 설명과 상담자료 제작에 필요한 도구입니다.",
    "실적 관리": "개인·조직 실적과 행사 달성 현황을 확인합니다.",
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
        <div class="hw-login-brand"><span class="hw-logo">H</span><strong>화랑 <b>WORKSPACE</b></strong></div>
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
            matched_user = next((name for name, saved in passwords.items() if password == saved), None)
            if matched_user:
                st.session_state["password_correct"] = True
                st.session_state["login_user"] = matched_user
                st.session_state["active_app"] = "home"
                st.rerun()
            st.error("입력한 비밀번호를 확인해 주세요.")

    with notice_col:
        render_notice()
    return False


def allowed_app_ids() -> list[str]:
    permissions = USER_PERMISSIONS.get(st.session_state.get("login_user"), {})
    return [app_id for app_id in APP_DEFINITIONS if permissions.get(app_id, False)]


def navigate(app_id: str) -> None:
    st.query_params.clear()
    st.session_state["active_app"] = app_id
    st.rerun()


def logout() -> None:
    st.query_params.clear()
    st.session_state.clear()
    st.rerun()


def render_sidebar(allowed_ids: list[str]) -> None:
    with st.sidebar:
        st.markdown('<div class="hw-side-brand"><span>H</span><strong>화랑 WORKSPACE</strong></div>', unsafe_allow_html=True)
        st.caption("필요한 업무를 선택하세요.")

        home_active = st.session_state["active_app"] == "home"
        if st.button("🏠  홈", key="nav_home", type="primary" if home_active else "secondary", use_container_width=True):
            navigate("home")

        for category in CATEGORY_ORDER:
            category_apps = [app_id for app_id in allowed_ids if APP_DEFINITIONS[app_id]["category"] == category]
            if not category_apps:
                continue
            st.markdown(f"#### {category}")
            for app_id in category_apps:
                app = APP_DEFINITIONS[app_id]
                active = st.session_state["active_app"] == app_id
                if st.button(
                    f"{app['icon']}  {app['name']}", key=f"nav_{app_id}",
                    type="primary" if active else "secondary", use_container_width=True,
                ):
                    navigate(app_id)

        st.divider()
        st.caption(f"접속 계정 · {st.session_state['login_user']}")
        with st.expander("최근 공지"):
            st.caption(NOTICE["date"])
            st.markdown(f"**{NOTICE['title']}**")
        if st.button("🚪  로그아웃", key="logout", use_container_width=True):
            logout()


def render_app_card(app_id: str, is_allowed: bool) -> None:
    app = APP_DEFINITIONS[app_id]
    card_key = f"available_card_{app_id}" if is_allowed else f"locked_card_{app_id}"
    with st.container(border=True, key=card_key):
        status = f'<span class="hw-card-badge">{app["status"]}</span>' if app.get("status") else ""
        lock_text = '<span class="hw-card-lock">권한 제한</span>' if not is_allowed else ""
        st.markdown(
            f"""
            <div class="hw-tool-card-head"><span class="hw-tool-icon">{app['icon']}</span><span class="hw-tool-category">{app['category']}{status}{lock_text}</span></div>
            <div class="hw-tool-title">{app['name']}</div>
            <div class="hw-tool-desc">{app['description']}</div>
            """,
            unsafe_allow_html=True,
        )
        label = "시작하기  →" if is_allowed else "🔒  사용 권한 없음"
        if st.button(label, key=f"home_{app_id}", disabled=not is_allowed, use_container_width=True):
            navigate(app_id)


def render_home(allowed_ids: list[str]) -> None:
    user = st.session_state["login_user"]
    st.markdown(
        """
        <style>
        .hw-home-intro{display:flex;justify-content:flex-end;align-items:center;margin:0 0 .55rem;padding:0 0 .75rem;border-bottom:1px solid #DFE9F1}
        .hw-home-user{padding:.5rem .82rem;border:1px solid #DCE6EE;border-radius:999px;background:rgba(255,255,255,.78);color:#536D80;font-size:.88rem;font-weight:650}
        .hw-category-head{margin:1.05rem 0 .72rem}.hw-category-head h2{margin:0 0 .28rem!important;color:#10283D!important;font-size:1.55rem!important;font-weight:800!important}
        .hw-category-head p{margin:0;color:#5F7486;font-size:.97rem;line-height:1.5}
        [class*="st-key-available_card_"]{min-height:13rem;background:rgba(255,255,255,.86);border:1px solid rgba(210,224,235,.9)!important;border-radius:1rem!important;box-shadow:0 10px 28px rgba(27,64,93,.055)}
        [class*="st-key-available_card_"]:hover{transform:translateY(-2px);box-shadow:0 15px 32px rgba(27,64,93,.1)}
        [class*="st-key-locked_card_"]{background:#F2F5F7!important;opacity:.72;border:1px dashed #B8C7D2!important;border-radius:1rem!important}
        .hw-tool-card-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:.8rem}.hw-tool-icon{width:2.65rem;height:2.65rem;display:grid;place-items:center;background:linear-gradient(145deg,#EAF3FF,#E9F8F7);border-radius:.8rem;font-size:1.05rem}
        .hw-tool-category{color:#718697;font-size:.63rem;font-weight:750}.hw-card-badge,.hw-card-lock{margin-left:.4rem;padding:.2rem .38rem;border-radius:999px;background:#EAF3FF;color:#1769DC;font-size:.56rem}.hw-card-lock{background:#E5EAEE;color:#697A87}
        .hw-tool-title{color:#10283D;font-size:1.06rem;font-weight:780;letter-spacing:-.035em;margin-bottom:.35rem}.hw-tool-desc{color:#647789;font-size:.77rem;line-height:1.55;min-height:2.4rem;margin-bottom:.55rem}
        [class*="st-key-available_card_insurer_portal"]{background:radial-gradient(circle at 90% 15%,rgba(23,105,220,.11),transparent 30%),linear-gradient(145deg,rgba(255,255,255,.95),rgba(241,249,255,.9))}
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(f'<div class="hw-home-intro"><div class="hw-home-user">{user} 계정으로 접속 중</div></div>', unsafe_allow_html=True)

    for category in CATEGORY_ORDER:
        category_apps = [app_id for app_id in APP_DEFINITIONS if APP_DEFINITIONS[app_id]["category"] == category]
        st.markdown(
            f'<div class="hw-category-head"><h2>{category}</h2><p>{CATEGORY_DESCRIPTIONS[category]}</p></div>',
            unsafe_allow_html=True,
        )
        for start in range(0, len(category_apps), 3):
            row_apps = category_apps[start:start + 3]
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

