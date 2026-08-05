import textwrap

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
    page_title="화랑WORKSPACE",
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
    "contact_url": "https://open.kakao.com/o/sFxdv4Rf",
}


APP_DEFINITIONS = {
    "analyzer": {
        "name": "보장 분석 도우미", "icon": "📑", "code": "BA", "category": "고객 상담",
        "badge": {"text": "BEST", "tone": "best"},
        "description": "보험사 보장분석 자료를 고객용 양식으로 변환합니다.", "action": "보장 분석 시작", "run": analyzer.run,
    },
    "remodeling": {
        "name": "보험 리모델링", "icon": "🔁", "code": "RM", "category": "고객 상담",
        "badge": {"text": "NEW", "tone": "new"},
        "description": "변경안을 비교하고 고객용 엑셀 자료를 만듭니다.", "action": "리모델링 시작", "run": remodeling.run,
    },
    "deposit_vs_shortpay": {
        "name": "적금 vs 단기납", "icon": "💰", "code": "DS", "category": "고객 상담",
        "description": "10년 기준 적금과 단기납의 예상 결과를 비교합니다.", "action": "비교 계산 시작", "run": deposit_vs_shortpay.run,
    },
    "renewal_vs_nonrenewal": {
        "name": "갱신 vs 비갱신", "icon": "📊", "code": "RN", "category": "고객 상담",
        "description": "보험료 변동을 반영해 장기 총납입액을 비교합니다.", "action": "보험료 비교 시작", "run": renewal_vs_nonrenewal.run,
    },
    "inheritance_tax": {
        "name": "상속세 계산기", "icon": "🧾", "code": "IT", "category": "고객 상담",
        "description": "예상 상속세와 부족한 현금성 납부재원을 계산합니다.", "action": "상속세 계산 시작", "run": inheritance_tax.run,
    },
    "insurer_portal": {
        "name": "원수사 전산 포털", "icon": "↗", "code": "IP", "category": "고객 상담",
        "description": "생명·손해보험사 원수사 전산을 한 화면에서 연결합니다.", "action": "전산 포털 열기", "run": insurer_portal.run,
    },
    "convention": {
        "name": "컨벤션 계산기", "icon": "🏆", "code": "CV", "category": "실적 관리",
        "description": "계약 실적을 환산하고 컨벤션 달성 여부를 확인합니다.", "action": "컨벤션 계산 시작", "run": convention.run,
    },
    "summer": {
        "name": "썸머 계산기", "icon": "🌞", "code": "SU", "category": "실적 관리",
        "description": "7·8월 업적을 반영해 썸머 업적을 계산합니다.", "action": "썸머 실적 계산", "run": summer.run,
    },
    "manager_results": {
        "name": "매니저 업적 환산", "icon": "📈", "code": "MR", "category": "실적 관리",
        "description": "지점 실적 환산금액을 집계합니다.", "action": "매니저 실적 확인", "run": manager_results.run,
    },
}


# 홈 카드용 아이콘입니다. 외부 이미지나 추가 패키지 없이 동일한 모양으로 표시됩니다.
HOME_ICONS = {
    "analyzer": '<svg viewBox="0 0 24 24"><path d="M9 11l2 2 4-4"/><path d="M12 3l7 3v5c0 4.6-3 8.1-7 10-4-1.9-7-5.4-7-10V6l7-3z"/></svg>',
    "remodeling": '<svg viewBox="0 0 24 24"><path d="M20 7h-6V1"/><path d="M20 7a9 9 0 10 1 7"/><path d="M4 17h6v6"/></svg>',
    "deposit_vs_shortpay": '<svg viewBox="0 0 24 24"><ellipse cx="12" cy="6" rx="7" ry="3"/><path d="M5 6v5c0 1.7 3.1 3 7 3s7-1.3 7-3V6"/><path d="M5 11v5c0 1.7 3.1 3 7 3 1 0 2-.1 2.8-.3"/><circle cx="18" cy="17" r="3"/><path d="M18 15.5v3M16.8 16.2h2.4"/></svg>',
    "renewal_vs_nonrenewal": '<svg viewBox="0 0 24 24"><path d="M20 7h-5V2"/><path d="M20 7a8 8 0 00-14.5-2"/><path d="M4 17h5v5"/><path d="M4 17a8 8 0 0014.5 2"/></svg>',
    "inheritance_tax": '<svg viewBox="0 0 24 24"><circle cx="9" cy="8" r="3"/><circle cx="16" cy="9" r="2.5"/><path d="M3 20c.3-4 2.3-6 6-6s5.7 2 6 6"/><path d="M14 14c4 0 6 2 6 6"/></svg>',
    "insurer_portal": '<svg viewBox="0 0 24 24"><path d="M5 3h14v18H5z"/><path d="M8 7h2M12 7h2M16 7h1M8 11h2M12 11h2M16 11h1"/><path d="M9 21v-5h6v5"/></svg>',
    "convention": '<svg viewBox="0 0 24 24"><path d="M8 4h8v4a4 4 0 01-8 0V4z"/><path d="M8 6H4c0 4 2 6 5 6M16 6h4c0 4-2 6-5 6"/><path d="M12 12v5M8 21h8M9 17h6v4"/></svg>',
    "summer": '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="4"/><path d="M12 2v3M12 19v3M2 12h3M19 12h3M4.9 4.9L7 7M17 17l2.1 2.1M19.1 4.9L17 7M7 17l-2.1 2.1"/></svg>',
    "manager_results": '<svg viewBox="0 0 24 24"><path d="M4 20V10M10 20V6M16 20V3M22 20H2"/><path d="M4 8l5-4 5 2 6-5"/><path d="M17 1h3v3"/></svg>',
}


USER_PERMISSIONS = {
    "Admin": {app_id: True for app_id in APP_DEFINITIONS},
    "Manager1": {app_id: True for app_id in APP_DEFINITIONS},
    "Basic": {
        "analyzer": True, "remodeling": False, "deposit_vs_shortpay": False,
        "renewal_vs_nonrenewal": False, "inheritance_tax": False,
        "insurer_portal": True,
        "convention": True, "summer": True, "manager_results": False,
    },
    "Crew": {
        "analyzer": True, "remodeling": False, "deposit_vs_shortpay": True,
        "renewal_vs_nonrenewal": True, "inheritance_tax": False,
        "insurer_portal": True,
        "convention": True, "summer": True, "manager_results": False,
    },
    "Dream": {
        "analyzer": True, "remodeling": True, "deposit_vs_shortpay": True,
        "renewal_vs_nonrenewal": True, "inheritance_tax": True,
        "insurer_portal": True,
        "convention": True, "summer": True, "manager_results": False,
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
    st.markdown(
        textwrap.dedent(
            f'''
            <style>
            .hw-login-contact {{ display:flex; align-items:center; justify-content:space-between; gap:.9rem;
                margin-top:.55rem; padding:.78rem .9rem; border:1px solid #C9DCF7; border-radius:.75rem;
                background:linear-gradient(135deg,#F3F8FF 0%,#EDF5FF 100%); }}
            .hw-login-contact-copy {{ display:flex; align-items:center; gap:.55rem; min-width:0;
                color:#3F5870; font-size:.82rem; line-height:1.4; }}
            .hw-login-contact-icon {{ flex:0 0 1.55rem; width:1.55rem; height:1.55rem; display:flex;
                align-items:center; justify-content:center; border-radius:50%; background:#DCEAFF;
                color:#2563D9; font-size:.76rem; font-weight:850; }}
            .hw-login-contact-link {{ flex:0 0 auto; display:inline-flex; align-items:center; gap:.28rem;
                padding:.48rem .72rem; border:1px solid #F0C900; border-radius:.58rem;
                background:#FEE500; color:#332A00 !important; text-decoration:none !important;
                font-size:.76rem; line-height:1; font-weight:800;
                box-shadow:0 4px 10px rgba(145,122,0,.12); transition:all .18s ease; }}
            .hw-login-contact-link:hover {{ transform:translateY(-1px); background:#FFEA35;
                box-shadow:0 6px 14px rgba(145,122,0,.18); }}
            @media(max-width:700px) {{
                .hw-login-contact {{ align-items:stretch; flex-direction:column; }}
                .hw-login-contact-link {{ justify-content:center; }}
            }}
            </style>
            <div class="hw-login-contact">
              <div class="hw-login-contact-copy">
                <span class="hw-login-contact-icon">i</span>
                <span>변경된 비밀번호가 필요하신가요?</span>
              </div>
              <a class="hw-login-contact-link" href="{NOTICE['contact_url']}" target="_blank" rel="noopener noreferrer">
                박병선 팀장에게 문의해 주세요 <span>↗</span>
              </a>
            </div>
            '''
        ),
        unsafe_allow_html=True,
    )


def render_login() -> bool:
    if st.session_state["password_correct"]:
        return True

    st.markdown(
        """
        <div class="hw-login-brand"><span class="hw-logo">H</span><strong>화랑 <b>WORKSPACE</b></strong></div>
        <div class="hw-login-hero">
          <div class="hw-login-copy">
            <span class="hw-login-kicker"><i></i>HWARANG BUSINESS WORKSPACE</span>
            <h1><span class="hw-title-top">보험 업무의 복잡함을,</span><em class="hw-title-accent">더 간단하게.</em></h1>
            <p>상담자료 제작부터 실적 관리까지 필요한 업무를 한곳에서 이용하세요.</p>
          </div>
          <div class="hw-glass-stack" aria-label="화랑 WORKSPACE 핵심 업무 영역">
            <div class="hw-glass-card"><span class="hw-glass-signal"></span><b>CONSULTING</b></div>
            <div class="hw-glass-card"><span class="hw-glass-signal"></span><b>PERFORMANCE</b></div>
            <div class="hw-glass-card"><span class="hw-glass-signal"></span><b>INSURANCE PORTAL</b></div>
          </div>
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
            else:
                st.error("입력한 비밀번호를 확인해 주세요.")

    with notice_col:
        render_notice()
    return False


def allowed_app_ids() -> list[str]:
    permissions = USER_PERMISSIONS.get(st.session_state.get("login_user"), {})
    return [app_id for app_id in APP_DEFINITIONS if permissions.get(app_id, False)]


def navigate(app_id: str) -> None:
    st.session_state["active_app"] = app_id
    st.rerun()


def logout() -> None:
    st.session_state.clear()
    st.rerun()


def render_sidebar(allowed_ids: list[str]) -> None:
    with st.sidebar:
        st.markdown(
            textwrap.dedent('''
            <style>
            .hw-side-brand-signature { display:flex; align-items:center; gap:.72rem; margin:.08rem 0 .4rem;
                padding:.72rem .75rem; border:1px solid rgba(94,142,207,.2); border-radius:.82rem;
                background:linear-gradient(145deg,rgba(255,255,255,.74),rgba(238,246,255,.72));
                box-shadow:0 6px 18px rgba(27,64,93,.055); }
            .hw-side-brand-mark { flex:0 0 2.35rem; width:2.35rem; height:2.35rem; display:flex;
                align-items:center; justify-content:center; border-radius:.68rem;
                background:linear-gradient(145deg,#2F73E0,#205CC3); color:#FFFFFF;
                box-shadow:0 5px 12px rgba(37,99,217,.2); font-size:1rem; font-weight:850; }
            .hw-side-brand-copy { display:flex; flex-direction:column; min-width:0; gap:.16rem; }
            .hw-side-brand-title { color:#17334B; font-size:.91rem; line-height:1.2; font-weight:800;
                letter-spacing:-.025em; white-space:nowrap; }
            .hw-side-brand-title b { color:#2563D9; font-weight:850; }
            .hw-side-brand-credit { margin:0 !important; padding:0 !important; color:#667D91;
                font-size:.61rem !important; line-height:1.3 !important; letter-spacing:0; white-space:nowrap; }
            .hw-side-brand-credit b { color:#294A67; font-weight:800; }
            </style>
            <div class="hw-side-brand-signature">
              <span class="hw-side-brand-mark">H</span>
              <div class="hw-side-brand-copy">
                <span class="hw-side-brand-title">화랑 <b>WORKSPACE</b></span>
                <p class="hw-side-brand-credit">Planned &amp; Built by <b>박병선</b></p>
              </div>
            </div>
            '''),
            unsafe_allow_html=True,
        )
        st.caption("필요한 업무를 선택하세요.")

        home_active = st.session_state["active_app"] == "home"
        if st.button("🏠  홈", key="nav_home", type="primary" if home_active else "secondary", use_container_width=True):
            navigate("home")

        for category in ("고객 상담", "실적 관리"):
            category_apps = [app_id for app_id in allowed_ids if APP_DEFINITIONS[app_id]["category"] == category]
            if not category_apps:
                continue
            st.markdown(f"#### {category}")
            for app_id in category_apps:
                app = APP_DEFINITIONS[app_id]
                active = st.session_state["active_app"] == app_id
                if st.button(f"{app['icon']}  {app['name']}", key=f"nav_{app_id}", type="primary" if active else "secondary", use_container_width=True):
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
        badge = app.get("badge")
        badge_html = ""
        if badge:
            badge_html = (
                f'<span class="hw-corner-badge hw-badge-{badge.get("tone", "default")}">'
                f'{badge["text"]}</span>'
            )
        lock_html = '<span class="hw-card-lock">권한 제한</span>' if not is_allowed else ""
        st.markdown(
            f"""
            {badge_html}{lock_html}
            <div class="hw-tool-heading">
              <span class="hw-tool-icon hw-icon-{app_id}">{HOME_ICONS[app_id]}</span>
              <div class="hw-tool-title">{app['name']}</div>
            </div>
            <div class="hw-tool-desc">{app['description']}</div>
            """,
            unsafe_allow_html=True,
        )
        if st.button(
            "시작하기  →" if is_allowed else "🔒  사용 권한 없음",
            key=f"home_{app_id}",
            disabled=not is_allowed,
            use_container_width=True,
        ):
            navigate(app_id)


def render_home(allowed_ids: list[str]) -> None:
    user = st.session_state["login_user"]
    st.markdown(
        """
        <style>
        [class*="st-key-home_intro"] { margin:0 0 1.15rem !important; padding:1.12rem 1.35rem !important;
            position:relative; overflow:hidden;
            background:
                radial-gradient(circle at 88% -20%,rgba(55,116,230,.15),transparent 38%),
                radial-gradient(circle at 58% 135%,rgba(70,175,201,.08),transparent 34%),
                linear-gradient(135deg,#FFFFFF 0%,#F8FBFF 58%,#F2F7FE 100%);
            border:1px solid #D3E1F0; border-top-color:#B8D2F7; border-radius:1.08rem;
            box-shadow:0 14px 34px rgba(24,55,85,.08),inset 0 1px 0 rgba(255,255,255,.95); }
        [class*="st-key-home_intro"]::before { content:""; position:absolute; z-index:0; top:0; left:2rem;
            width:7.5rem; height:2px; border-radius:0 0 999px 999px;
            background:linear-gradient(90deg,#2563EB,#57B6CC); opacity:.88; }
        [class*="st-key-home_intro"]::after { content:""; position:absolute; z-index:0; right:-2.8rem; top:-3.8rem;
            width:10rem; height:10rem; border:1px solid rgba(86,135,209,.12); border-radius:50%;
            box-shadow:0 0 0 1.7rem rgba(95,145,220,.035); pointer-events:none; }
        [class*="st-key-home_intro"] > div { position:relative; z-index:1; }
        .hw-home-greeting { display:flex; align-items:center; gap:1rem; min-height:3.45rem; }
        .hw-home-avatar { flex:0 0 3.35rem; width:3.35rem; height:3.35rem; display:flex; align-items:center;
            justify-content:center; border:1px solid rgba(80,137,225,.3); border-radius:.92rem;
            background:linear-gradient(145deg,#FFFFFF 0%,#EAF2FF 100%); color:#2563D9;
            box-shadow:0 7px 18px rgba(37,99,217,.11),inset 0 1px 0 #FFFFFF; }
        .hw-home-avatar svg { width:1.8rem; height:1.8rem; fill:none; stroke:currentColor; stroke-width:1.75;
            stroke-linecap:round; filter:drop-shadow(0 2px 3px rgba(37,99,217,.12)); }
        .hw-home-copy { display:flex; flex-direction:column; justify-content:center; gap:.28rem; min-width:0; }
        .hw-home-copy h1 { margin:0 !important; padding:0 !important; color:#10283D !important;
            font-size:1.48rem !important; line-height:1.22 !important; font-weight:800 !important;
            letter-spacing:-.035em !important; }
        .hw-home-copy p { margin:0 !important; padding:0 !important; color:#64798C;
            font-size:.84rem !important; line-height:1.4 !important; letter-spacing:-.012em; }
        .hw-category-head { margin:1rem 0 .62rem !important; padding:0 !important; }
        .hw-category-head h2 { margin:0 0 .18rem !important; padding:0 !important; color:#10283D !important;
            font-size:1.48rem !important; line-height:1.25 !important; font-weight:800 !important;
            letter-spacing:-.035em !important; }
        .hw-category-head p { margin:0 !important; padding:0 !important; color:#5F7486;
            font-size:.86rem !important; line-height:1.45 !important; }
        [class*="st-key-locked_card_"] { background-color:#F2F5F7 !important; opacity:.72;
            border:1px dashed #B8C7D2 !important; border-radius:.95rem !important; position:relative; min-height:10.65rem; }
        [class*="st-key-available_card_"] { min-height:10.65rem; position:relative; overflow:visible;
            background:#FFFFFF; border:1px solid #DCE6EE !important; border-radius:.95rem !important;
            box-shadow:0 7px 22px rgba(27,64,93,.055); transition:transform .18s ease,box-shadow .18s ease; }
        [class*="st-key-available_card_"]:hover { transform:translateY(-2px); box-shadow:0 12px 28px rgba(27,64,93,.1); }
        [class*="st-key-available_card_"] button,
        [class*="st-key-locked_card_"] button { width:100% !important; min-height:2.65rem !important; margin:0 !important;
            padding:.48rem .8rem !important; background:#FFFFFF !important; border:1px solid #C8D9E7 !important;
            border-radius:.62rem !important; box-shadow:none !important; color:#1769DC !important;
            font-size:.79rem !important; font-weight:750 !important;
            transition:background-color .18s ease,color .18s ease,border-color .18s ease,box-shadow .18s ease,transform .18s ease !important; }
        [class*="st-key-available_card_"] button:hover { color:#FFFFFF !important; background:#1769DC !important;
            border-color:#1769DC !important; box-shadow:0 7px 16px rgba(23,105,220,.2) !important; transform:translateY(-1px); }
        [class*="st-key-available_card_"] button:active { transform:translateY(0); box-shadow:0 3px 9px rgba(23,105,220,.18) !important; }
        [class*="st-key-available_card_"] button:focus-visible { outline:3px solid rgba(23,105,220,.2) !important; outline-offset:2px; }
        [class*="st-key-locked_card_"] button:disabled { background:#E9EEF2 !important; border-color:#D5DEE5 !important;
            color:#7B8C99 !important; opacity:1 !important; }
        .hw-tool-heading { display:flex; align-items:center; gap:.72rem; min-height:3rem; padding-right:3.65rem; margin-bottom:.48rem; }
        .hw-tool-icon { flex:0 0 2.65rem; width:2.65rem; height:2.65rem; display:flex; align-items:center;
            justify-content:center; border:1px solid #CDDEFA; border-radius:.72rem; background:#F3F7FF; color:#2F6FDB; }
        .hw-tool-icon svg { width:1.55rem; height:1.55rem; fill:none; stroke:currentColor; stroke-width:1.75;
            stroke-linecap:round; stroke-linejoin:round; }
        .hw-icon-remodeling,.hw-icon-renewal_vs_nonrenewal,.hw-icon-summer { color:#10A6AA; background:#EFFBFA; border-color:#C8ECEA; }
        .hw-icon-deposit_vs_shortpay,.hw-icon-manager_results { color:#D89412; background:#FFF8E9; border-color:#F3DEAA; }
        .hw-icon-inheritance_tax { color:#7856D8; background:#F6F2FF; border-color:#DDD3FA; }
        .hw-tool-title { color:#10283D; font-size:1.02rem; line-height:1.3; font-weight:800; letter-spacing:-.035em; }
        .hw-tool-desc { color:#647789; font-size:.76rem; line-height:1.55; min-height:2.4rem; margin:0 0 .45rem 3.37rem; padding-right:.25rem; }
        .hw-corner-badge { position:absolute; z-index:3; top:.88rem; right:.88rem; display:inline-flex;
            align-items:center; justify-content:center; height:1.48rem; min-width:2.9rem; padding:0 .58rem;
            border-radius:999px; font-size:.62rem; line-height:1; font-weight:850; letter-spacing:.055em; }
        .hw-badge-best { background:#F6C453; color:#4A3100; border:1px solid #E7AE2B; box-shadow:0 4px 10px rgba(231,174,43,.2); }
        .hw-badge-new { background:#0EA5A8; color:#FFFFFF; border:1px solid #079195; box-shadow:0 4px 10px rgba(14,165,168,.22); }
        .hw-badge-default { background:#1769DC; color:#FFFFFF; border:1px solid #0E5BC4; }
        .hw-card-lock { position:absolute; z-index:4; top:.88rem; right:.88rem; padding:.25rem .55rem;
            border-radius:999px; background:#E5EAEE; color:#697A87; font-size:.6rem; font-weight:750; }
        [class*="st-key-locked_card_"] .hw-corner-badge { display:none; }
        .hw-home-footer { display:flex; align-items:center; justify-content:center; gap:.75rem;
            margin:2.15rem 0 .45rem; padding:1.05rem 1.25rem;
            border:1px solid #D9E5F1; border-radius:.95rem;
            background:linear-gradient(135deg,rgba(255,255,255,.96),rgba(244,249,255,.96));
            box-shadow:0 8px 24px rgba(27,64,93,.055); text-align:left; }
        .hw-footer-mark { flex:0 0 2.35rem; width:2.35rem; height:2.35rem; display:flex;
            align-items:center; justify-content:center; border-radius:.7rem;
            background:linear-gradient(145deg,#2F73E0,#205CC3); color:#FFFFFF;
            box-shadow:0 6px 14px rgba(37,99,217,.2); font-size:1rem; font-weight:850; }
        .hw-footer-copy { display:flex; flex-direction:column; gap:.14rem; }
        .hw-footer-brand { color:#17334B; font-size:.82rem; line-height:1.25; font-weight:750;
            letter-spacing:-.015em; }
        .hw-footer-brand b { color:#2563D9; font-weight:850; }
        .hw-footer-credit { margin:0 !important; padding:0 !important; color:#697E91;
            font-size:.72rem !important; line-height:1.35 !important; letter-spacing:.01em; }
        .hw-footer-credit b { color:#2B4861; font-weight:800; }
        @media(max-width:900px){
            .hw-tool-desc{margin-left:0}.hw-tool-heading{padding-right:3.3rem}
        }
        @media(max-width:650px){
            [class*="st-key-home_intro"]{padding:.9rem !important}.hw-home-greeting{margin-bottom:.35rem}
            .hw-tool-desc{min-height:auto}.hw-category-head{margin-top:1.2rem !important}
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    with st.container(key="home_intro"):
        account_col, search_col = st.columns([1.15, 1], gap="large")
        with account_col:
            st.markdown(
                f'''<div class="hw-home-greeting">
                  <span class="hw-home-avatar"><svg viewBox="0 0 24 24"><circle cx="12" cy="8" r="4"/><path d="M5 22v-2a7 7 0 0114 0v2"/></svg></span>
                  <div class="hw-home-copy"><h1>안녕하세요, {user}님</h1><p>오늘 필요한 업무를 빠르게 시작해 보세요.</p></div>
                </div>''',
                unsafe_allow_html=True,
            )
        with search_col:
            insurer_portal.render_home_quick_search()

    for category in ("고객 상담", "실적 관리"):
        category_apps = [app_id for app_id in APP_DEFINITIONS if APP_DEFINITIONS[app_id]["category"] == category]
        description = "고객 설명과 상담자료 제작에 필요한 도구입니다." if category == "고객 상담" else "개인·조직 실적과 행사 달성 현황을 확인합니다."
        st.markdown(f'<div class="hw-category-head"><h2>{category}</h2><p>{description}</p></div>', unsafe_allow_html=True)
        for start in range(0, len(category_apps), 3):
            row_apps = category_apps[start:start + 3]
            columns = st.columns(3, gap="medium")
            for column, app_id in zip(columns, row_apps):
                with column:
                    render_app_card(app_id, app_id in allowed_ids)

    st.markdown(
        '''<div class="hw-home-footer">
          <span class="hw-footer-mark">H</span>
          <div class="hw-footer-copy">
            <span class="hw-footer-brand">화랑 <b>WORKSPACE</b></span>
            <p class="hw-footer-credit">Planned &amp; Built by <b>박병선</b></p>
          </div>
        </div>''',
        unsafe_allow_html=True,
    )


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
