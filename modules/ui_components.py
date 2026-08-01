"""화랑 WORKS 공통 화면 구성요소와 디자인 시스템."""

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
        }
        html, body, [class*="css"] { font-family: Pretendard, "Noto Sans KR", "Apple SD Gothic Neo", sans-serif; }
        .stApp { background: var(--hw-bg); color: var(--hw-ink); }
        /*
         * Streamlit Community Cloud의 고정 상단바는 약 64~66px입니다.
         * 88px(5.5rem)의 상단 여백으로 상단바와 20px 이상의 안전 간격을 확보합니다.
         */
        header[data-testid="stHeader"] {
            background: rgba(247, 250, 252, 0.90);
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
        [data-testid="stSidebar"] { background: #FFFFFF; border-right: 1px solid var(--hw-line); }
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
            background: #FFFFFF; border: 1px dashed #AFC4D4; border-radius: 16px; padding: 1rem;
        }
        [data-testid="stFileUploaderDropzone"]:hover { border-color: var(--hw-blue); background: #FAFCFF; }
        [data-testid="stMetric"] {
            background: #FFFFFF; border: 1px solid var(--hw-line); border-radius: 15px;
            padding: 1.15rem 1.25rem; box-shadow: 0 7px 22px rgba(30,70,100,.045);
        }
        [data-testid="stMetricLabel"] { color: var(--hw-muted); }
        [data-testid="stMetricValue"] { color: var(--hw-ink); letter-spacing: -0.035em; }
        [data-testid="stExpander"] {
            background: #FFFFFF; border: 1px solid var(--hw-line) !important; border-radius: 14px !important;
            box-shadow: 0 5px 18px rgba(30,70,100,.035);
        }
        [data-testid="stAlert"] { border-radius: 13px; }
        [data-testid="stDataFrame"], [data-testid="stDataEditor"] {
            background: #FFFFFF; border: 1px solid var(--hw-line); border-radius: 13px; overflow: hidden;
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
            color:var(--hw-blue); font-size:1.45rem; border:1px solid #D8E7F4; }
        .hw-page-copy { min-width:0; }
        .hw-breadcrumb { color:#52758A; font-size:.73rem; font-weight:800; letter-spacing:.08em; text-transform:uppercase; margin-bottom:.25rem; }
        .hw-page-title { margin:0; font-size:2rem; line-height:1.25; font-weight:780; letter-spacing:-.055em; color:var(--hw-ink); }
        .hw-page-desc { margin:.35rem 0 0; color:var(--hw-muted); font-size:.93rem; line-height:1.55; }
        .hw-section-label { color:var(--hw-blue); font-size:.72rem; font-weight:800; letter-spacing:.1em; }
        .hw-side-brand { display:flex; align-items:center; gap:.7rem; margin:.1rem 0 1rem; color:#10283D; }
        .hw-side-brand span,.hw-login-brand .hw-logo { width:2.25rem; height:2.25rem; display:grid; place-items:center; border-radius:.7rem;
            background:linear-gradient(145deg,#1769DC,#119B98); color:white; font-weight:900; box-shadow:0 8px 18px rgba(23,105,220,.18); }
        .hw-side-brand strong { font-size:1.05rem; letter-spacing:-.035em; }
        .hw-login-brand { display:flex; align-items:center; gap:.75rem; margin:.2rem 0 2.1rem; }
        .hw-login-brand strong { color:#10283D; font-size:1.2rem; letter-spacing:-.04em; }
        .hw-login-brand b { font-weight:800; }
        .hw-login-hero { margin-bottom:2rem; padding:2.6rem 3rem; border:1px solid #DFE9F1; border-radius:22px;
            background:radial-gradient(circle at 88% 25%,rgba(23,105,220,.13),transparent 28%),radial-gradient(circle at 75% 80%,rgba(17,155,152,.12),transparent 26%),linear-gradient(120deg,#FFFFFF,#F5F9FF);
            box-shadow:0 18px 45px rgba(30,70,100,.07); }
        .hw-login-hero>span { color:#3F7197; font-size:.68rem; font-weight:850; letter-spacing:.13em; }
        .hw-login-hero h1 { margin:.8rem 0 1rem; font-size:clamp(2.4rem,4.2vw,4rem); line-height:1.13; letter-spacing:-.06em; }
        .hw-login-hero h1 em { color:#1769DC; font-style:normal; }
        .hw-login-hero p { margin:0; color:#5F7486; font-size:.96rem; line-height:1.65; }
        [data-testid="stSidebar"] .stButton>button[kind="primary"] { background:#EAF3FF !important; color:#1769DC !important; border-color:#CFE1F4 !important; }
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
            .hw-login-hero { padding:2rem 1.4rem; }
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
            <div class="hw-breadcrumb">화랑 WORKS&nbsp;&nbsp;/&nbsp;&nbsp;{html.escape(category)}</div>
            <div class="hw-page-title">{html.escape(title)}</div>
            <div class="hw-page-desc">{html.escape(description)}</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_intro(label: str, title: str, description: str = "") -> None:
    st.markdown(f'<div class="hw-section-label">{html.escape(label)}</div>', unsafe_allow_html=True)
    st.subheader(title)
    if description:
        st.caption(description)
