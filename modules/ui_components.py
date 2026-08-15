"""화랑 WORKSPACE 공통 화면 구성요소와 디자인 시스템."""

from __future__ import annotations

import html
import streamlit as st


def inject_global_styles() -> None:
    st.markdown(
        """
        <style>
        :root {
            --hw-ink: #10283D;
            --hw-muted: #647789;
            --hw-blue: #1769DC;
            --hw-teal: #119B98;
            --hw-line: #DCE6EE;
            --hw-bg: #F7FAFC;
            --hw-soft-blue: #EAF3FF;
            --hw-soft-teal: #E9F8F7;
            --hw-surface: rgba(255, 255, 255, 0.92);
            --hw-premium-line: rgba(191, 211, 226, 0.78);
            --hw-premium-shadow: 0 14px 38px rgba(37, 72, 98, 0.075), 0 2px 8px rgba(37, 72, 98, 0.035);
        }
        html, body, [class*="css"] { font-family: Pretendard, "Noto Sans KR", "Apple SD Gothic Neo", sans-serif; }
        /* 로그인 후 홈·내부 프로그램: 중앙은 밝고 가장자리에만 은은한 색감을 둡니다. */
        .stApp {
            color: var(--hw-ink);
            background:
                radial-gradient(circle at 88% 8%, rgba(23,105,220,.075) 0%, rgba(23,105,220,.028) 24%, transparent 43%),
                radial-gradient(circle at 8% 84%, rgba(17,155,152,.060) 0%, rgba(17,155,152,.022) 25%, transparent 44%),
                linear-gradient(145deg, #F7FAFD 0%, #FFFFFF 45%, #F3F8FB 100%) !important;
            background-attachment: fixed !important;
        }
        /* 로그인 화면: 브랜드 첫인상을 위해 배경의 블루·민트 농도를 한 단계 높입니다. */
        .stApp:has(.hw-login-hero) {
            background:
                radial-gradient(circle at 84% 10%, rgba(23,105,220,.13) 0%, rgba(23,105,220,.050) 26%, transparent 47%),
                radial-gradient(circle at 12% 88%, rgba(17,155,152,.10) 0%, rgba(17,155,152,.035) 27%, transparent 48%),
                linear-gradient(140deg, #F5F9FD 0%, #FFFFFF 43%, #EFF7F8 100%) !important;
        }
        /*
         * Streamlit Community Cloud의 고정 상단바는 약 64~66px입니다.
         * 88px(5.5rem)의 상단 여백으로 상단바와 20px 이상의 안전 간격을 확보합니다.
         */
        header[data-testid="stHeader"] {
            background: rgba(248, 251, 253, 0.78);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
        }
        #MainMenu { visibility: hidden; }
        [data-testid="stDecoration"] { display: none; }
        .block-container { max-width: 1280px; padding-bottom: 5rem; }
        /*
         * Streamlit 버전별 본문 컨테이너 이름을 모두 지원합니다.
         * 배포 화면에서는 기존 .block-container 선택자만으로는 상단 여백이 적용되지 않았습니다.
         */
        [data-testid="stAppViewContainer"] .main .block-container,
        [data-testid="stAppViewBlockContainer"],
        [data-testid="stMainBlockContainer"],
        .stMainBlockContainer,
        main .block-container {
            max-width: 1280px;
            padding-top: 3rem !important;
            padding-bottom: 5rem !important;
        }
        /* app.py의 홈 히어로 음수 상단 여백도 안전하게 무효화합니다. */
        .hw-home-hero { margin-top: 0 !important; }
        h1, h2, h3, h4 { color: var(--hw-ink); letter-spacing: -0.035em; }
        h1 a, h2 a, h3 a, h4 a { display: none !important; }
        h2 { margin-top: 2.5rem !important; }
        h3 { margin-top: 1.6rem !important; }
        [data-testid="stCaptionContainer"] { color: var(--hw-muted); }
        [data-testid="stSidebar"] {
            background: rgba(255,255,255,.91) !important;
            border-right: 1px solid rgba(199,216,228,.82);
            box-shadow: 10px 0 34px rgba(40,72,96,.035);
            backdrop-filter: blur(15px);
            -webkit-backdrop-filter: blur(15px);
        }
        [data-testid="stSidebar"] .block-container { padding-top: 1.4rem !important; }
        [data-testid="stSidebar"] hr { border-color: #E8EEF3; }
        .stButton > button, .stDownloadButton > button, [data-testid="stFormSubmitButton"] > button {
            min-height: 2.75rem; border-radius: 10px; font-weight: 700;
            border-color: #C9D7E2; transition: transform .16s ease, box-shadow .16s ease, border-color .16s ease;
        }
        .stButton > button:hover, .stDownloadButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover {
            transform: translateY(-1px); border-color: var(--hw-blue); box-shadow: 0 7px 18px rgba(23,105,220,.10);
        }
        button[kind="primary"], .stDownloadButton > button[kind="primary"] {
            background: var(--hw-blue) !important; border-color: var(--hw-blue) !important; color: white !important;
        }
        [data-testid="stFileUploaderDropzone"] {
            background: rgba(255,255,255,.88); border: 1px dashed #A9C1D3; border-radius: 16px; padding: 1rem;
            box-shadow: 0 9px 26px rgba(37,72,98,.045);
        }
        /* 모든 도구의 입력 구역을 같은 카드 문법으로 정리합니다. */
        [data-testid="stForm"] {
            background: rgba(255,255,255,.94);
            border: 1px solid var(--hw-premium-line) !important;
            border-radius: 16px !important;
            padding: 1.25rem 1.35rem 1.35rem !important;
            box-shadow: var(--hw-premium-shadow);
        }
        [data-testid="stForm"] [data-testid="stFormSubmitButton"] {
            margin-top: .45rem;
        }
        [data-testid="stFileUploaderDropzone"]:hover { border-color: var(--hw-blue); background: #FAFCFF; }
        [data-testid="stMetric"] {
            background: var(--hw-surface); border: 1px solid var(--hw-premium-line); border-radius: 15px;
            padding: 1.15rem 1.25rem; box-shadow: var(--hw-premium-shadow);
            backdrop-filter: blur(10px); -webkit-backdrop-filter: blur(10px);
        }
        [data-testid="stMetricLabel"] { color: var(--hw-muted); }
        [data-testid="stMetricValue"] { color: var(--hw-ink); letter-spacing: -0.035em; }
        [data-testid="stExpander"] {
            background: var(--hw-surface); border: 1px solid var(--hw-premium-line) !important; border-radius: 14px !important;
            box-shadow: 0 10px 30px rgba(37,72,98,.055);
            backdrop-filter: blur(10px); -webkit-backdrop-filter: blur(10px);
        }
        [data-testid="stAlert"] { border-radius: 13px; }
        [data-testid="stDataFrame"], [data-testid="stDataEditor"] {
            background: rgba(255,255,255,.94); border: 1px solid var(--hw-premium-line); border-radius: 13px;
            overflow: hidden; box-shadow: 0 10px 28px rgba(37,72,98,.05);
        }
        /* Streamlit border container와 홈 프로그램 카드에 동일한 표면 질감을 적용합니다. */
        [data-testid="stVerticalBlockBorderWrapper"] {
            background: rgba(255,255,255,.88);
            border-radius: 16px;
            box-shadow: 0 11px 32px rgba(37,72,98,.055);
            backdrop-filter: blur(9px);
            -webkit-backdrop-filter: blur(9px);
        }
        [class*="st-key-available_card_"] {
            background: rgba(255,255,255,.93) !important;
            border-color: rgba(190,211,226,.86) !important;
            box-shadow: var(--hw-premium-shadow) !important;
        }
        [class*="st-key-available_card_"]:hover {
            border-color: rgba(133,174,207,.92) !important;
            box-shadow: 0 18px 44px rgba(37,72,98,.115), 0 3px 10px rgba(37,72,98,.05) !important;
        }
        [data-baseweb="tab-list"] { gap: .35rem; background: #EDF3F7; padding: .3rem; border-radius: 12px; }
        [data-baseweb="tab"] { height: 2.7rem; border-radius: 9px; padding: 0 1rem; }
        [aria-selected="true"][data-baseweb="tab"] { background: #FFFFFF; box-shadow: 0 2px 8px rgba(20,55,80,.09); }
        [data-baseweb="tab-highlight"], [data-baseweb="tab-border"] { display: none; }
        [data-baseweb="input"], [data-baseweb="select"] > div, textarea {
            border-radius: 10px !important; border-color: #CAD8E3 !important; background: #FFFFFF !important;
        }
        hr { border-color: #E3EBF1; }
        .hw-page-head { display:flex; align-items:flex-start; gap:1rem; margin: .1rem 0 1.65rem; }
        .hw-page-icon { flex:0 0 auto; width:3.3rem; height:3.3rem; display:grid; place-items:center;
            border-radius:1rem; background:linear-gradient(145deg,var(--hw-soft-blue),var(--hw-soft-teal));
            color:var(--hw-blue); font-size:1.45rem; border:1px solid rgba(190,215,233,.88);
            box-shadow:0 9px 24px rgba(37,72,98,.06); }
        .hw-page-copy { min-width:0; }
        .hw-breadcrumb { color:#52758A; font-size:.73rem; font-weight:800; letter-spacing:.08em; text-transform:uppercase; margin-bottom:.25rem; }
        .hw-page-title { margin:0; font-size:2rem; line-height:1.25; font-weight:780; letter-spacing:-.055em; color:var(--hw-ink); }
        .hw-page-desc { margin:.35rem 0 0; color:var(--hw-muted); font-size:.93rem; line-height:1.55; }
        .hw-section-label { color:var(--hw-blue); font-size:.72rem; font-weight:800; letter-spacing:.1em; }
        .hw-section-head { margin:2.15rem 0 .85rem; }
        .hw-section-head .hw-section-label { margin-bottom:.24rem; }
        .hw-section-title { margin:0; color:var(--hw-ink); font-size:1.42rem; line-height:1.35;
            font-weight:780; letter-spacing:-.045em; }
        .hw-section-desc { margin:.32rem 0 0; color:var(--hw-muted); font-size:.88rem; line-height:1.55; }
        .hw-side-brand { display:flex; align-items:center; gap:.7rem; margin:.1rem 0 1rem; color:#10283D; }
        .hw-side-brand span,.hw-login-brand .hw-logo { width:2.25rem; height:2.25rem; display:grid; place-items:center; border-radius:.7rem;
            background:linear-gradient(145deg,#1769DC,#119B98); color:white; font-weight:900; box-shadow:0 8px 18px rgba(23,105,220,.18); }
        .hw-side-brand strong { font-size:1.05rem; letter-spacing:-.035em; }
        .hw-login-brand { display:flex; align-items:center; gap:.75rem; margin:0 0 1.25rem !important; }
        .hw-login-brand strong { color:#10283D; font-size:1.2rem; letter-spacing:-.04em; }
        .hw-login-brand b { font-weight:800; }
        .hw-login-hero { position:relative; overflow:hidden; display:grid; grid-template-columns:minmax(0,1.08fr) minmax(250px,.92fr);
            align-items:center; gap:2rem; min-height:0 !important; height:auto !important; margin-bottom:1.5rem !important;
            padding:1.9rem 2.4rem !important; border:1px solid rgba(184,208,226,.82); border-radius:18px;
            background:radial-gradient(circle at 88% 22%,rgba(23,105,220,.17),transparent 31%),radial-gradient(circle at 72% 86%,rgba(17,155,152,.13),transparent 30%),linear-gradient(120deg,rgba(255,255,255,.97),rgba(243,248,255,.94));
            box-shadow:0 22px 55px rgba(34,70,98,.10), inset 0 1px 0 rgba(255,255,255,.88);
            backdrop-filter:blur(12px); -webkit-backdrop-filter:blur(12px); }
        .hw-login-copy { position:relative; z-index:2; min-width:0; }
        .hw-login-kicker { display:flex; align-items:center; gap:.5rem; margin:0 0 .75rem !important; padding:0 !important;
            color:#3F7197; font-size:.65rem !important; font-weight:850; letter-spacing:.13em; }
        .hw-login-kicker i { width:.42rem; height:.42rem; flex:0 0 auto; border-radius:50%; background:#119B98;
            box-shadow:0 0 0 .28rem rgba(17,155,152,.10); }
        .hw-login-hero h1 { margin:0 0 1.15rem !important; padding:0 !important;
            font-size:clamp(2.1rem,3.5vw,3.35rem) !important; line-height:1 !important; letter-spacing:normal; }
        .hw-title-top { display:block; line-height:1.08 !important; letter-spacing:-.055em; }
        .hw-title-accent { display:block; margin-top:.18em; color:#1769DC; font-style:normal;
            line-height:1.04 !important; letter-spacing:-.045em; }
        .hw-login-hero p { margin:0 !important; padding:0 !important; color:#5F7486;
            font-size:.92rem !important; line-height:1.65 !important; }
        .hw-glass-stack { position:relative; z-index:1; min-height:13.5rem; }
        .hw-glass-card { position:absolute; left:8%; right:2%; display:flex; align-items:center; gap:.75rem;
            padding:.9rem 1.05rem; border:1px solid rgba(190,211,226,.70); border-radius:.85rem;
            background:rgba(255,255,255,.66); color:#29495F; box-shadow:0 16px 36px rgba(31,68,96,.09),inset 0 1px 0 rgba(255,255,255,.90);
            backdrop-filter:blur(13px); -webkit-backdrop-filter:blur(13px); animation:hw-glass-float 6.5s ease-in-out infinite; }
        .hw-glass-card::before { content:""; width:1.1rem; height:1px; background:linear-gradient(90deg,#119B98,#8EC9C7); }
        .hw-glass-card:nth-child(1) { top:3%; left:18%; animation-delay:0s; }
        .hw-glass-card:nth-child(2) { top:37%; right:10%; animation-delay:-2.1s; }
        .hw-glass-card:nth-child(3) { top:71%; left:24%; animation-delay:-4.2s; }
        .hw-glass-card b { font-size:.76rem; font-weight:800; letter-spacing:.09em; }
        .hw-glass-signal { width:.42rem; height:.42rem; flex:0 0 auto; border-radius:50%; background:#119B98;
            box-shadow:0 0 0 .26rem rgba(17,155,152,.09); }
        @keyframes hw-glass-float { 0%,100% { transform:translateY(0); } 50% { transform:translateY(-3px); } }
        @media (prefers-reduced-motion: reduce) { .hw-glass-card { animation:none !important; } }
        [data-testid="stSidebar"] .stButton>button[kind="primary"] { background:#EAF3FF !important; color:#1769DC !important; border-color:#CFE1F4 !important; }
        @media (max-width: 880px) {
            .hw-login-hero { grid-template-columns:1fr; }
            .hw-glass-stack { display:none; }
        }
        @media (max-width: 768px) {
            /* 모바일 상단바 아래에도 약 16px의 안전 여백을 둡니다. */
            [data-testid="stAppViewContainer"] .main .block-container,
            [data-testid="stAppViewBlockContainer"],
            [data-testid="stMainBlockContainer"],
            .stMainBlockContainer,
            main .block-container {
                padding: 5.5rem 1rem 4rem !important;
            }
            .hw-page-title { font-size:1.65rem; }
            .hw-page-icon { width:2.9rem; height:2.9rem; }
            .hw-login-hero { padding:1.55rem 1.4rem !important; }
            [data-testid="stHorizontalBlock"] { gap:.8rem; }
        }
        @media print {
            [data-testid="stSidebar"], [data-testid="stHeader"], .stButton, .stDownloadButton { display:none !important; }
            .stApp, .block-container { background:white !important; padding-top:0 !important; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def page_header(category: str, title: str, description: str, icon: str) -> None:
    st.markdown(
        f"""
        <div class="hw-page-head">
          <div class="hw-page-icon">{html.escape(icon)}</div>
          <div class="hw-page-copy">
            <div class="hw-breadcrumb">화랑 WORKSPACE&nbsp;&nbsp;/&nbsp;&nbsp;{html.escape(category)}</div>
            <div class="hw-page-title">{html.escape(title)}</div>
            <div class="hw-page-desc">{html.escape(description)}</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_intro(label: str, title: str, description: str = "") -> None:
    desc = f'<div class="hw-section-desc">{html.escape(description)}</div>' if description else ""
    st.markdown(
        f"""
        <div class="hw-section-head">
          <div class="hw-section-label">{html.escape(label)}</div>
          <div class="hw-section-title">{html.escape(title)}</div>
          {desc}
        </div>
        """,
        unsafe_allow_html=True,
    )
