import base64
import html
import textwrap
from datetime import date, datetime
from pathlib import Path

import streamlit as st
import modules.commission_calculator as commission_calculator
import modules.insurance_claim_guide as insurance_claim_guide
import modules.silson_generation_comparison as silson_generation_comparison

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


@st.cache_data(show_spinner=False)
def _pretendard_font_data() -> str:
    font_path = Path(__file__).resolve().parent / "assets" / "fonts" / "PretendardVariable.ttf"
    if not font_path.is_file():
        return ""
    return base64.b64encode(font_path.read_bytes()).decode("ascii")


def inject_pretendard_font() -> None:
    font_data = _pretendard_font_data()
    if not font_data:
        return
    st.markdown(
        f"""
        <style>
        @font-face {{
            font-family: 'Pretendard';
            src: url(data:font/ttf;base64,{font_data}) format('truetype');
            font-weight: 100 900;
            font-style: normal;
            font-display: swap;
        }}
        html, body, [class*="css"], [data-testid="stAppViewContainer"],
        [data-testid="stSidebar"], button, input, textarea, select {{
            font-family: 'Pretendard', 'Noto Sans KR', 'Apple SD Gothic Neo', sans-serif !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


inject_pretendard_font()


# 공지는 이 목록만 수정하면 로그인 화면에 반영됩니다.
NOTICE = {
    "date": "2026.08.01",
    "title": "화랑WORKSPACE 화면이 새롭게 개편되었습니다.",
    "items": [
        "로그인 후 통합 홈에서 필요한 업무를 선택할 수 있습니다.",
        "고객 상담과 실적 관리 메뉴가 업무 목적별로 구분되었습니다.",
        "썸머·컨벤션·상속세 계산기가 추가되었습니다.",
        "비밀번호 입력 후 Enter 키를 눌러 로그인할 수 있습니다.",
    ],
    "important": "8월1일부터 비밀번호가 변경되었습니다. 변경된 비밀번호는 박병선 팀장에게 문의해 주세요.",
    "contact_url": "https://open.kakao.com/o/sFxdv4Rf",
}

# 홈 공지와 업데이트는 이 목록만 수정하면 자동으로 최신순 정렬됩니다.
HOME_NOTICES = [
    {
        "date": "2026.08.15", "title": "화랑 WORKSPACE 홈 화면 개편 안내",
        "content": "통합검색과 업무별 도구 구성을 적용했습니다. 필요한 기능을 더 빠르게 찾아보세요.",
        "important": True, "visible": True,
    },
    {
        "date": "2026.08.01", "title": NOTICE["title"],
        "content": " ".join(NOTICE["items"][:2]), "important": True, "visible": True,
    },
    {
        "date": "2026.07.20", "title": "고객 상담 도구 업데이트",
        "content": "상담자료의 디자인과 다운로드 문서 구성을 통일했습니다.",
        "important": False, "visible": True,
    },
    {
        "date": "2026.07.05", "title": "원수사 전산 포털 안내",
        "content": "보험사별 전산 주소와 고객센터 번호를 한 화면에서 확인할 수 있습니다.",
        "important": False, "visible": True,
    },
    {
        "date": "2026.06.18", "title": "실적 관리 도구 개선",
        "content": "컨벤션·썸머·매니저 업적 환산 화면을 정비했습니다.",
        "important": False, "visible": True,
    },
]

HOME_UPDATES = [
    {"date": "2026.08.17", "app": "실손보험 세대 비교", "content": "보험료 기준과 PDF 그래프를 개선했습니다.", "visible": True},
    {"date": "2026.08.15", "app": "보험금 청구 가이드", "content": "고객 안내자료 구성을 개선했습니다.", "visible": True},
    {"date": "2026.08.12", "app": "원수사 전산 포털", "content": "한화라이프랩과 고객센터 번호를 추가했습니다.", "visible": True},
    {"date": "2026.08.10", "app": "수수료 계산기", "content": "상품 자동 매칭과 화면 구조를 정비했습니다.", "visible": True},
    {"date": "2026.08.05", "app": "보험 리모델링", "content": "상세 작성과 즉시 미리보기를 적용했습니다.", "visible": True},
]


APP_DEFINITIONS = {
    "analyzer": {
        "name": "보장 분석 도우미", "icon": "📑", "code": "BA", "category": "고객 상담",
        "badge": {"text": "BEST", "tone": "best"},
        "description": "보험사 보장분석 자료를 고객용 양식으로 변환합니다.", "action": "보장 분석 시작", "run": analyzer.run,
    },
    "remodeling": {
        "name": "보험 리모델링", "icon": "🔁", "code": "RM", "category": "고객 상담",
        "badge": {"text": "NEW", "tone": "new", "until": "2026-09-16"},
        "description": "변경안을 비교하고 고객용 엑셀 자료를 만듭니다.", "action": "리모델링 시작", "run": remodeling.run,
    },
    "deposit_vs_shortpay": {
        "name": "적금 vs 단기납", "icon": "💰", "code": "DS", "category": "고객 상담",
        "badge": {"text": "UPDATE", "tone": "update", "until": "2026-09-16"},
        "description": "10년 기준 적금과 단기납의 예상 결과를 비교합니다.", "action": "비교 계산 시작", "run": deposit_vs_shortpay.run,
    },
    "renewal_vs_nonrenewal": {
        "name": "갱신 vs 비갱신", "icon": "📊", "code": "RN", "category": "고객 상담",
        "badge": {"text": "UPDATE", "tone": "update", "until": "2026-09-16"},
        "description": "보험료 변동을 반영해 장기 총납입액을 비교합니다.", "action": "보험료 비교 시작", "run": renewal_vs_nonrenewal.run,
    },
    "inheritance_tax": {
        "name": "상속세 계산기", "icon": "🧾", "code": "IT", "category": "고객 상담",
        "badge": {"text": "NEW", "tone": "new", "until": "2026-09-16"},
        "description": "예상 상속세와 부족한 현금성 납부재원을 계산합니다.", "action": "상속세 계산 시작", "run": inheritance_tax.run,
    },
    "insurer_portal": {
        "name": "원수사 전산 포털", "icon": "↗", "code": "IP", "category": "고객 상담",
        "badge": {"text": "NEW", "tone": "new", "until": "2026-09-16"},
        "description": "생명·손해보험사 원수사 전산을 한 화면에서 연결합니다.", "action": "전산 포털 열기", "run": insurer_portal.run,
    },
    "insurance_claim_guide": {
        "name": "보험금 청구 가이드", "icon": "📋", "code": "CG", "category": "고객 상담",
        "badge": {"text": "NEW", "tone": "new", "until": "2026-09-16"},
        "description": "청구 항목별 필요서류를 안내하고 보장분석 PDF에서 관련 담보를 찾습니다.",
        "action": "청구 가이드 시작", "run": insurance_claim_guide.run,
    },
    "silson_generation_comparison": {
        "name": "실손보험 세대 비교", "icon": "🩺", "code": "SC", "category": "고객 상담",
        "badge": {"text": "NEW", "tone": "new", "until": "2026-09-16"},
        "description": "현재 가입 실손과 5세대 실손의 보험료와 입원 보장을 비교합니다.",
        "action": "실손 세대 비교 시작", "run": silson_generation_comparison.run,
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
    "commission_calculator": {
        "name": "수수료 계산기", "icon": "💼", "code": "CC", "category": "실적 관리",
        "badge": {"text": "NEW", "tone": "new", "until": "2026-09-16"},
        "description": "생보·손보 예시표에서 상품별 수수료율을 찾아 예상 수당을 계산합니다.",
        "action": "수수료 계산 시작", "run": commission_calculator.run,
    },
}


# 홈 카드용 아이콘입니다. 외부 이미지나 추가 패키지 없이 동일한 모양으로 표시됩니다.
HOME_ICONS = {
    "analyzer": '<svg viewBox="0 0 24 24"><path d="M9 11l2 2 4-4"/><path d="M12 3l7 3v5c0 4.6-3 8.1-7 10-4-1.9-7-5.4-7-10V6l7-3z"/></svg>',
    "insurance_claim_guide": '<svg viewBox="0 0 24 24"><path d="M7 3h10v3H7z"/><path d="M5 5h14v16H5z"/><path d="M8 11l2 2 4-4M8 17h8"/></svg>',
    "silson_generation_comparison": '<svg viewBox="0 0 24 24"><path d="M6 3v6a6 6 0 0012 0V3"/><path d="M9 3v5a3 3 0 006 0V3M12 15v6"/><circle cx="18" cy="18" r="3"/></svg>',
    "remodeling": '<svg viewBox="0 0 24 24"><path d="M20 7h-6V1"/><path d="M20 7a9 9 0 10 1 7"/><path d="M4 17h6v6"/></svg>',
    "deposit_vs_shortpay": '<svg viewBox="0 0 24 24"><ellipse cx="12" cy="6" rx="7" ry="3"/><path d="M5 6v5c0 1.7 3.1 3 7 3s7-1.3 7-3V6"/><path d="M5 11v5c0 1.7 3.1 3 7 3 1 0 2-.1 2.8-.3"/><circle cx="18" cy="17" r="3"/><path d="M18 15.5v3M16.8 16.2h2.4"/></svg>',
    "renewal_vs_nonrenewal": '<svg viewBox="0 0 24 24"><path d="M20 7h-5V2"/><path d="M20 7a8 8 0 00-14.5-2"/><path d="M4 17h5v5"/><path d="M4 17a8 8 0 0014.5 2"/></svg>',
    "inheritance_tax": '<svg viewBox="0 0 24 24"><circle cx="9" cy="8" r="3"/><circle cx="16" cy="9" r="2.5"/><path d="M3 20c.3-4 2.3-6 6-6s5.7 2 6 6"/><path d="M14 14c4 0 6 2 6 6"/></svg>',
    "insurer_portal": '<svg viewBox="0 0 24 24"><path d="M5 3h14v18H5z"/><path d="M8 7h2M12 7h2M16 7h1M8 11h2M12 11h2M16 11h1"/><path d="M9 21v-5h6v5"/></svg>',
    "convention": '<svg viewBox="0 0 24 24"><path d="M8 4h8v4a4 4 0 01-8 0V4z"/><path d="M8 6H4c0 4 2 6 5 6M16 6h4c0 4-2 6-5 6"/><path d="M12 12v5M8 21h8M9 17h6v4"/></svg>',
    "summer": '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="4"/><path d="M12 2v3M12 19v3M2 12h3M19 12h3M4.9 4.9L7 7M17 17l2.1 2.1M19.1 4.9L17 7M7 17l-2.1 2.1"/></svg>',
    "manager_results": '<svg viewBox="0 0 24 24"><path d="M4 20V10M10 20V6M16 20V3M22 20H2"/><path d="M4 8l5-4 5 2 6-5"/><path d="M17 1h3v3"/></svg>',
    "commission_calculator": '<svg viewBox="0 0 24 24"><path d="M4 7h16v13H4z"/><path d="M8 7V4h8v3M4 11h16"/><path d="M9 15h6M12 13v4"/></svg>',
}


USER_PERMISSIONS = {
    "Admin": {
        "insurance_claim_guide": True,
        "silson_generation_comparison": True,
        "analyzer": True, "remodeling": True, "deposit_vs_shortpay": True,
        "renewal_vs_nonrenewal": True, "inheritance_tax": True,
        "insurer_portal": True,
        "convention": True, "summer": True, "manager_results": True,
        "commission_calculator": True,
    },
    "Manager1": {
        "insurance_claim_guide": True,
        "silson_generation_comparison": True,
        "analyzer": True, "remodeling": True, "deposit_vs_shortpay": True,
        "renewal_vs_nonrenewal": True, "inheritance_tax": True,
        "insurer_portal": True,
        "convention": True, "summer": True, "manager_results": True,
        "commission_calculator": True,
    },
    "Basic": {
        "insurance_claim_guide": True,
        "silson_generation_comparison": True,
        "analyzer": True, "remodeling": False, "deposit_vs_shortpay": False,
        "renewal_vs_nonrenewal": False, "inheritance_tax": False,
        "insurer_portal": True,
        "convention": True, "summer": True, "manager_results": False,
        "commission_calculator": False,
    },
    "Crew": {
        "insurance_claim_guide": True,
        "silson_generation_comparison": True,
        "analyzer": True, "remodeling": False, "deposit_vs_shortpay": True,
        "renewal_vs_nonrenewal": True, "inheritance_tax": False,
        "insurer_portal": True,
        "convention": True, "summer": True, "manager_results": False,
        "commission_calculator": False,
    },
    "Dream": {
        "insurance_claim_guide": True,
        "silson_generation_comparison": True,
        "analyzer": True, "remodeling": True, "deposit_vs_shortpay": True,
        "renewal_vs_nonrenewal": True, "inheritance_tax": True,
        "insurer_portal": True,
        "convention": True, "summer": True, "manager_results": False,
        "commission_calculator": False,
    },
}


def initialize_state() -> None:
    st.session_state.setdefault("password_correct", False)
    st.session_state.setdefault("login_user", None)
    st.session_state.setdefault("active_app", "home")
    st.session_state.setdefault("permission_request_app", None)
    st.session_state.setdefault("show_notices_dialog", False)
    st.session_state.setdefault("show_updates_dialog", False)
    st.session_state.setdefault("home_unified_search", "")


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


def _date_value(value: str) -> datetime:
    return datetime.strptime(value, "%Y.%m.%d")


def _visible_badge(app: dict) -> dict | None:
    badge = app.get("badge")
    if not badge or not badge.get("until"):
        return badge
    return badge if date.today() <= date.fromisoformat(badge["until"]) else None


def _sorted_notices() -> list[dict]:
    return sorted((item for item in HOME_NOTICES if item.get("visible", True)), key=lambda item: _date_value(item["date"]), reverse=True)


def _sorted_updates() -> list[dict]:
    return sorted((item for item in HOME_UPDATES if item.get("visible", True)), key=lambda item: _date_value(item["date"]), reverse=True)


@st.dialog("사용 권한이 필요합니다", width="small")
def render_permission_dialog() -> None:
    app_id = st.session_state.get("permission_request_app")
    app_name = APP_DEFINITIONS.get(app_id, {}).get("name", "선택한 기능")
    st.markdown(f"### 🔒 {app_name}")
    st.write("현재 계정에서는 사용할 수 없는 기능입니다.")
    st.caption("사용 권한이 필요하면 박병선 팀장에게 문의해 주세요.")
    st.link_button("카카오톡 문의", NOTICE["contact_url"], use_container_width=True, type="primary")
    if st.button("닫기", key="close_permission_dialog", use_container_width=True):
        st.session_state["permission_request_app"] = None
        st.rerun()


@st.dialog("공지사항", width="large")
def render_notices_dialog() -> None:
    notices = _sorted_notices()[:5]
    for index, item in enumerate(notices):
        label = f"{'[중요] ' if item.get('important') else ''}{item['title']} · {item['date']}"
        with st.expander(label, expanded=index == 0):
            st.write(item["content"])
    if st.button("닫기", key="close_notices_dialog", use_container_width=True):
        st.session_state["show_notices_dialog"] = False
        st.rerun()


@st.dialog("최근 업데이트", width="large")
def render_updates_dialog() -> None:
    with st.container(height=470, border=False):
        previous_month = None
        for item in _sorted_updates():
            month = item["date"][:7]
            if month != previous_month:
                st.markdown(f"#### {month.replace('.', '년 ')}월")
                previous_month = month
            st.markdown(f"**{item['date']} · {item['app']}**  \n{item['content']}")
            st.divider()
    if st.button("닫기", key="close_updates_dialog", use_container_width=True):
        st.session_state["show_updates_dialog"] = False
        st.rerun()


def render_sidebar(allowed_ids: list[str]) -> None:
    with st.sidebar:
        st.markdown(
            textwrap.dedent('''
            <style>
            [data-testid="stSidebar"] > div:first-child {background:linear-gradient(180deg,#082F50 0%,#06395E 62%,#063152 100%);}
            [data-testid="stSidebar"] *{font-family:'Pretendard','Noto Sans KR',sans-serif;}
            .hw-side-brand-signature {display:flex;align-items:center;gap:.72rem;margin:.08rem 0 .55rem;padding:.72rem .72rem;
                border:1px solid rgba(255,255,255,.14);border-radius:.82rem;background:rgba(255,255,255,.055);}
            .hw-side-brand-mark { flex:0 0 2.35rem; width:2.35rem; height:2.35rem; display:flex;
                align-items:center;justify-content:center;border:1px solid rgba(255,255,255,.72);border-radius:.68rem;
                background:rgba(255,255,255,.08);color:#FFFFFF;font-size:1rem;font-weight:850;}
            .hw-side-brand-copy { display:flex; flex-direction:column; min-width:0; gap:.16rem; }
            .hw-side-brand-title { color:#FFFFFF; font-size:.91rem; line-height:1.2; font-weight:800;
                letter-spacing:-.025em; white-space:nowrap; }
            .hw-side-brand-title b {color:#9EC8F1;font-weight:850;}
            .hw-side-brand-credit {margin:0!important;padding:0!important;color:#B8CCE0;
                font-size:.61rem !important; line-height:1.3 !important; letter-spacing:0; white-space:nowrap; }
            .hw-side-brand-credit b {color:#FFFFFF;font-weight:800;}
            .hw-side-section{margin:.8rem .18rem .28rem;color:#9FB8CE;font-size:.68rem;font-weight:750;letter-spacing:.04em;}
            .hw-sidebar-user{position:sticky;bottom:.35rem;z-index:4;margin-top:.8rem;padding:.72rem .75rem;border:1px solid rgba(255,255,255,.14);
                border-radius:.78rem;background:rgba(5,38,65,.94);color:#FFFFFF;box-shadow:0 -8px 22px rgba(2,28,49,.16);}
            .hw-sidebar-user b{font-size:.84rem}.hw-sidebar-user span{display:block;margin-top:.12rem;color:#AAC1D5;font-size:.64rem;}
            [data-testid="stSidebar"] .stButton button{min-height:2.15rem;padding:.35rem .58rem;border-color:transparent;background:transparent;color:#E7F0F7;text-align:left;justify-content:flex-start;font-size:.76rem;}
            [data-testid="stSidebar"] .stButton button:hover{background:rgba(255,255,255,.09);border-color:rgba(255,255,255,.1);color:#FFFFFF;}
            [data-testid="stSidebar"] .stButton button[kind="primary"]{background:#286BA8;border-color:#367CB9;color:#FFFFFF;}
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

        home_active = st.session_state["active_app"] == "home"
        if st.button("🏠  홈", key="nav_home", type="primary" if home_active else "secondary", use_container_width=True):
            navigate("home")

        for category in ("고객 상담", "실적 관리"):
            category_apps = [app_id for app_id in APP_DEFINITIONS if APP_DEFINITIONS[app_id]["category"] == category]
            st.markdown(f'<div class="hw-side-section">{category}</div>', unsafe_allow_html=True)
            for app_id in category_apps:
                app = APP_DEFINITIONS[app_id]
                active = st.session_state["active_app"] == app_id
                is_allowed = app_id in allowed_ids
                label = f"{app['icon']}  {app['name']}" if is_allowed else f"🔒  {app['name']}"
                if st.button(label, key=f"nav_{app_id}", type="primary" if active else "secondary", use_container_width=True):
                    if is_allowed:
                        navigate(app_id)
                    else:
                        st.session_state["permission_request_app"] = app_id
                        st.rerun()

        st.markdown(f'<div class="hw-sidebar-user"><b>{html.escape(str(st.session_state["login_user"]))}</b><span>화랑 WORKSPACE</span></div>', unsafe_allow_html=True)
        if st.button("🚪로그아웃", key="logout", use_container_width=True):
            logout()


def render_app_card(app_id: str, is_allowed: bool) -> None:
    app = APP_DEFINITIONS[app_id]
    badge = _visible_badge(app)
    badge_html = ""
    if badge:
        badge_html = f'<span class="hw-card-badge hw-card-badge-{badge.get("tone", "default")}">{html.escape(badge["text"])}</span>'
    locked_html = '<span class="hw-card-locked">🔒 권한 필요</span>' if not is_allowed else '<span class="hw-card-arrow">→</span>'
    href = f"?go={app_id}" if is_allowed else f"?locked={app_id}"
    category_class = "consulting" if app["category"] == "고객 상담" else "performance"
    st.markdown(
        f'''<a class="hw-tool-card hw-tool-card-{category_class}{' is-locked' if not is_allowed else ''}" href="{href}" target="_self">
          <span class="hw-tool-card-icon">{HOME_ICONS[app_id]}</span>
          <span class="hw-tool-card-title">{html.escape(app['name'])}</span>
          {badge_html}{locked_html}
        </a>''',
        unsafe_allow_html=True,
    )


def _render_home_legacy(allowed_ids: list[str]) -> None:
    """이전 홈 디자인 보관용입니다. 실제 실행은 아래 개편 홈을 사용합니다."""
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
        .hw-badge-update { background:linear-gradient(135deg,#2F73E0,#205CC3); color:#FFFFFF; border:1px solid #1B55B6; box-shadow:0 4px 11px rgba(37,99,217,.22); }
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


def render_home(allowed_ids: list[str]) -> None:
    """2026 홈 개편안: 검색·공지·업무 도구를 한 화면에 정리합니다."""
    user = html.escape(str(st.session_state["login_user"]))
    st.markdown(
        """
        <style>
        .hw-topline{display:flex;align-items:flex-start;justify-content:space-between;gap:1rem;margin:.1rem 0 1rem;}
        .hw-topline h1{margin:0!important;color:#112C43!important;font-size:1.58rem!important;line-height:1.2!important;font-weight:850!important;letter-spacing:-.04em!important;}
        .hw-topline p{margin:.28rem 0 0!important;color:#687D8E;font-size:.84rem;}
        .hw-top-date{padding-top:.2rem;color:#435D72;font-size:.82rem;font-weight:700;white-space:nowrap;}
        [class*="st-key-home_hero"]{position:relative;overflow:visible;margin-bottom:1.1rem;padding:1.45rem 1.7rem!important;border:1px solid #214E74;border-radius:1.05rem;background:linear-gradient(135deg,#0A3659,#0B426B);box-shadow:0 13px 30px rgba(12,48,76,.14);}
        [class*="st-key-home_hero"] [data-testid="column"]:last-child{position:relative;z-index:20;}
        .hw-hero-title{position:relative;z-index:2;padding:.12rem 0;}.hw-hero-title h2{margin:0 0 .48rem!important;color:white!important;font-size:1.7rem!important;font-weight:850!important;letter-spacing:-.045em!important;}
        .hw-hero-title p{max-width:24rem;margin:0!important;color:#C3D5E3;font-size:.82rem;line-height:1.65;}
        .hw-hero-mark{position:absolute;right:2rem;top:50%;transform:translateY(-50%);z-index:0;width:5.2rem;height:5.8rem;display:flex;align-items:center;justify-content:center;border:2px solid rgba(255,255,255,.12);border-radius:1.3rem;color:rgba(255,255,255,.1);font-family:serif;font-size:3.1rem;font-weight:900;}
        [class*="st-key-home_hero"] [data-testid="stTextInput"] input{height:3.35rem;padding-left:1.05rem;border:1px solid rgba(255,255,255,.65);border-radius:.82rem;background:#FFFFFF;color:#17364E;font-size:.88rem;box-shadow:0 8px 20px rgba(2,27,45,.18);}
        .hw-search-results{position:absolute;z-index:999;top:3.65rem;left:0;right:0;overflow:hidden;border:1px solid #D6E0E8;border-radius:.85rem;background:#FFFFFF;box-shadow:0 18px 44px rgba(17,45,67,.2);}
        .hw-search-section{padding:.55rem .82rem .28rem;color:#8393A0;font-size:.65rem;font-weight:800;letter-spacing:.04em;}
        .hw-search-divider{height:1px;margin:.2rem .8rem;background:#E3E9EE;}
        .hw-search-row{display:grid;grid-template-columns:2rem minmax(9rem,1fr) 6.7rem 7.4rem 3.5rem;align-items:center;gap:.45rem;min-height:3rem;padding:.48rem .75rem;color:#17364E!important;text-decoration:none!important;border-top:1px solid #EFF3F6;transition:background .16s ease;}
        .hw-search-row:hover{background:#F3F7FB;}.hw-search-row.is-locked{background:#F6F7F8;color:#8A98A3!important;}
        .hw-search-icon{display:flex;align-items:center;justify-content:center;color:#2D6EAD;font-size:1rem}.hw-search-name{font-size:.82rem;font-weight:800;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
        .hw-search-category{color:#81909C;font-size:.68rem;white-space:nowrap;}.hw-search-phone{color:#17364E;font-size:.82rem;font-weight:800;font-variant-numeric:tabular-nums;text-align:right;white-space:nowrap;}.hw-search-open{color:#2D6EAD;font-size:.72rem;font-weight:800;text-align:right;white-space:nowrap;}
        .hw-search-empty{padding:1rem;color:#7A8B98;font-size:.76rem;text-align:center;}
        [class*="st-key-home_notice"],[class*="st-key-home_updates"]{height:13.2rem;padding:1.05rem 1.15rem!important;border:1px solid #DCE5EC!important;border-radius:.95rem!important;background:#FFFFFF;box-shadow:0 8px 24px rgba(27,64,93,.055);}
        .hw-info-head{display:flex;align-items:center;justify-content:space-between;gap:.6rem;margin-bottom:.8rem}.hw-info-head h3{margin:0!important;color:#17364E!important;font-size:1.05rem!important;font-weight:850!important;}.hw-important{padding:.22rem .48rem;border:1px solid #E8B750;border-radius:999px;background:#FFF8E7;color:#A56600;font-size:.58rem;font-weight:850;}
        .hw-notice-title{color:#17364E;font-size:.9rem;font-weight:800}.hw-notice-date{float:right;color:#778A99;font-size:.68rem}.hw-notice-copy{margin-top:.55rem;color:#64798A;font-size:.73rem;line-height:1.55}.hw-update-row{display:grid;grid-template-columns:1fr auto;gap:.7rem;padding:.48rem 0;border-bottom:1px solid #EDF1F4;color:#36566D;font-size:.72rem}.hw-update-row time{color:#81919E;font-variant-numeric:tabular-nums;white-space:nowrap;}
        [class*="st-key-home_notice"] button,[class*="st-key-home_updates"] button{min-height:2rem!important;margin-top:.4rem;border:0!important;background:transparent!important;color:#2D6EAD!important;justify-content:flex-start!important;padding:.15rem 0!important;box-shadow:none!important;font-size:.72rem!important;font-weight:800!important;}
        .hw-section-title{margin:1.15rem 0 .62rem;color:#17364E;font-size:1.13rem;font-weight:850;letter-spacing:-.03em;}
        .hw-tool-card{position:relative;display:flex;align-items:center;gap:.7rem;min-height:5.1rem;padding:.8rem .9rem;border:1px solid #DCE5EC;border-radius:.83rem;background:#FFFFFF;color:#17364E!important;text-decoration:none!important;box-shadow:0 5px 18px rgba(27,64,93,.045);transition:transform .18s ease,box-shadow .18s ease,border-color .18s ease;}
        .hw-tool-card:hover{transform:translateY(-3px);box-shadow:0 12px 27px rgba(27,64,93,.12);border-color:#72A3D0;}.hw-tool-card-performance:hover{border-color:#63AFAA;}
        .hw-tool-card-icon{flex:0 0 2.35rem;width:2.35rem;height:2.35rem;display:flex;align-items:center;justify-content:center;color:#2D6EAD}.hw-tool-card-performance .hw-tool-card-icon{color:#2A918C}.hw-tool-card-icon svg{width:1.55rem;height:1.55rem;fill:none;stroke:currentColor;stroke-width:1.75;stroke-linecap:round;stroke-linejoin:round}.hw-tool-card-title{min-width:0;color:#17364E;font-size:.82rem;font-weight:800;line-height:1.3}.hw-card-arrow{margin-left:auto;color:#2D6EAD;font-size:1rem}.hw-card-badge{margin-left:auto;padding:.2rem .4rem;border-radius:999px;font-size:.54rem;font-weight:850}.hw-card-badge-best{background:#F4C85D;color:#523900}.hw-card-badge-new{background:#2A918C;color:#FFF}.hw-card-badge-update,.hw-card-badge-default{background:#2D6EAD;color:#FFF}.hw-card-locked{margin-left:auto;color:#84929C;font-size:.62rem;font-weight:750}.hw-tool-card.is-locked{background:#F4F6F7;border-style:dashed;opacity:.74;}
        .hw-home-signature{margin:1.4rem 0 .3rem;text-align:center;color:#728493;font-size:.7rem;}
        @media(max-width:900px){.hw-topline{align-items:flex-start}.hw-search-row{grid-template-columns:1.8rem 1fr 4rem}.hw-search-category,.hw-search-phone{display:none}[class*="st-key-home_notice"],[class*="st-key-home_updates"]{height:auto}.hw-hero-mark{display:none}}
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        f'<div class="hw-topline"><div><h1>안녕하세요, {user}님</h1><p>오늘 필요한 업무를 빠르게 시작해 보세요.</p></div>'
        f'<span class="hw-top-date">{date.today():%Y년 %m월 %d일}</span></div>',
        unsafe_allow_html=True,
    )

    with st.container(key="home_hero"):
        st.markdown('<span class="hw-hero-mark">H</span>', unsafe_allow_html=True)
        title_col, search_col = st.columns([0.32, 0.68], gap="large")
        with title_col:
            st.markdown('<div class="hw-hero-title"><h2>스마트한 업무의 시작</h2><p>고객 상담부터 실적 관리까지 필요한 도구를 한곳에서 확인하세요.</p></div>', unsafe_allow_html=True)
        with search_col:
            query = st.text_input("통합 업무 검색", key="home_unified_search", placeholder="기능명, 보험사 또는 업무를 검색하세요", label_visibility="collapsed")
            normalized = "".join(query.lower().split())
            if normalized:
                function_matches = []
                for app_id, app in APP_DEFINITIONS.items():
                    haystack = "".join(f"{app['name']} {app['description']} {app['category']}".lower().split())
                    if normalized in haystack:
                        function_matches.append((0 if "".join(app["name"].lower().split()).startswith(normalized) else 1, app_id, app))
                function_matches.sort(key=lambda item: (item[0], item[2]["name"]))

                insurer_items = []
                insurer_sources = []
                for attr in ("HANWHA_LIFELAB_PORTAL",):
                    value = getattr(insurer_portal, attr, None)
                    if value:
                        insurer_sources.append(value)
                insurer_sources.extend(getattr(insurer_portal, "LIFE_INSURERS", []))
                insurer_sources.extend(getattr(insurer_portal, "NON_LIFE_INSURERS", []))
                phone_map = getattr(insurer_portal, "CUSTOMER_CENTER_NUMBERS", {})
                aliases = getattr(insurer_portal, "SEARCH_ALIASES", {})
                for insurer in insurer_sources:
                    name = str(insurer.get("name", ""))
                    haystack = "".join((name + " " + " ".join(aliases.get(name, ()))).lower().split())
                    if normalized in haystack:
                        insurer_items.append((0 if "".join(name.lower().split()).startswith(normalized) else 1, name, phone_map.get(name, "")))
                insurer_items.sort(key=lambda item: (item[0], item[1]))

                rows = []
                remaining = 5
                if function_matches:
                    rows.append('<div class="hw-search-section">기능</div>')
                    for _, app_id, app in function_matches[:remaining]:
                        allowed = app_id in allowed_ids
                        href = f"?go={app_id}" if allowed else f"?locked={app_id}"
                        rows.append(
                            f'<a class="hw-search-row{"" if allowed else " is-locked"}" href="{href}" target="_self">'
                            f'<span class="hw-search-icon">⌕</span><span class="hw-search-name">{html.escape(app["name"])}</span>'
                            f'<span class="hw-search-category">{html.escape(app["category"])}</span><span class="hw-search-phone"></span>'
                            f'<span class="hw-search-open">{"열기 →" if allowed else "🔒 권한 필요"}</span></a>'
                        )
                    remaining -= min(len(function_matches), remaining)
                if remaining and insurer_items:
                    if function_matches:
                        rows.append('<div class="hw-search-divider"></div>')
                    rows.append('<div class="hw-search-section">보험사</div>')
                    for _, name, phone in insurer_items[:remaining]:
                        rows.append(
                            f'<a class="hw-search-row" href="?go=insurer_portal" target="_self"><span class="hw-search-icon">▥</span>'
                            f'<span class="hw-search-name">{html.escape(name)}</span><span class="hw-search-category">원수사 포털</span>'
                            f'<span class="hw-search-phone">{html.escape(phone)}</span><span class="hw-search-open">열기 →</span></a>'
                        )
                if not rows:
                    rows.append('<div class="hw-search-empty">일치하는 기능이나 보험사를 찾지 못했습니다.</div>')
                st.markdown('<div class="hw-search-results">' + "".join(rows) + '</div>', unsafe_allow_html=True)

    notice_col, update_col = st.columns([0.55, 0.45], gap="medium")
    notices = _sorted_notices()
    latest_notice = next((item for item in notices if item.get("important")), notices[0] if notices else None)
    with notice_col:
        with st.container(key="home_notice"):
            if latest_notice:
                st.markdown(
                    f'<div class="hw-info-head"><h3>공지사항</h3>{"<span class=\"hw-important\">IMPORTANT</span>" if latest_notice.get("important") else ""}</div>'
                    f'<div><span class="hw-notice-title">{html.escape(latest_notice["title"])}</span><span class="hw-notice-date">{latest_notice["date"]}</span></div>'
                    f'<div class="hw-notice-copy">{html.escape(latest_notice["content"])}</div>', unsafe_allow_html=True)
            if st.button("전체 공지 보기  →", key="open_notices_dialog"):
                st.session_state["show_notices_dialog"] = True
                st.rerun()
    with update_col:
        with st.container(key="home_updates"):
            st.markdown('<div class="hw-info-head"><h3>최근 업데이트</h3></div>', unsafe_allow_html=True)
            for item in _sorted_updates()[:3]:
                st.markdown(f'<div class="hw-update-row"><span>{html.escape(item["app"])} · {html.escape(item["content"])}</span><time>{item["date"]}</time></div>', unsafe_allow_html=True)
            if st.button("전체 업데이트 보기  →", key="open_updates_dialog"):
                st.session_state["show_updates_dialog"] = True
                st.rerun()

    for category in ("고객 상담", "실적 관리"):
        st.markdown(f'<div class="hw-section-title">{category} 도구</div>', unsafe_allow_html=True)
        category_apps = [app_id for app_id in APP_DEFINITIONS if APP_DEFINITIONS[app_id]["category"] == category]
        columns = st.columns(4, gap="medium")
        for index, app_id in enumerate(category_apps):
            if index and index % 4 == 0:
                columns = st.columns(4, gap="medium")
            with columns[index % 4]:
                render_app_card(app_id, app_id in allowed_ids)

    st.markdown('<div class="hw-home-signature">Planned &amp; Built by 박병선 팀장</div>', unsafe_allow_html=True)


def main() -> None:
    initialize_state()
    if not render_login():
        st.stop()

    allowed_ids = allowed_app_ids()
    requested_app = st.query_params.get("go")
    requested_locked = st.query_params.get("locked")
    if requested_app:
        st.query_params.clear()
        st.session_state["home_unified_search"] = ""
        if requested_app in allowed_ids:
            st.session_state["active_app"] = requested_app
        elif requested_app in APP_DEFINITIONS:
            st.session_state["permission_request_app"] = requested_app
        st.rerun()
    if requested_locked:
        st.query_params.clear()
        st.session_state["permission_request_app"] = requested_locked if requested_locked in APP_DEFINITIONS else None
        st.rerun()

    active_app = st.session_state.get("active_app", "home")
    if active_app != "home" and active_app not in allowed_ids:
        st.session_state["active_app"] = "home"
        active_app = "home"

    render_sidebar(allowed_ids)
    if st.session_state.get("permission_request_app"):
        render_permission_dialog()
    if st.session_state.get("show_notices_dialog"):
        render_notices_dialog()
    if st.session_state.get("show_updates_dialog"):
        render_updates_dialog()
    if active_app == "home":
        render_home(allowed_ids)
    else:
        APP_DEFINITIONS[active_app]["run"]()


if __name__ == "__main__":
    main()
