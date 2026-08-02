"""화랑 WORKSPACE 보험회사 원수사 전산 포털."""

from __future__ import annotations

import base64
import html
from pathlib import Path
from urllib.parse import urlparse

import streamlit as st

try:
    from modules.ui_components import page_header
except ImportError:  # 모듈 단독 미리보기용
    from ui_components import page_header


PROJECT_ROOT = Path(__file__).resolve().parents[1]
LOGO_DIR = PROJECT_ROOT / "assets" / "insurer_logos"

LIFE_INSURERS = [
    {"name": "한화생명", "slug": "hanwha_life", "url": "https://hmp.hanwhalife.com/online/solutions/websquare/websquare.html?w2xPath=/online/ui/uv/pmn/uvpmn010mvw.xml"},
    {"name": "라이나생명", "slug": "lina_life", "url": "https://ga.lina.co.kr/"},    
    {"name": "미래에셋생명", "slug": "miraeasset_life", "url": "https://www.loveageplan.com/"},   
    {"name": "KB라이프생명", "slug": "kb_life", "url": "https://sfa.kblife.co.kr/"},
    {"name": "신한라이프", "slug": "shinhan_life", "url": "https://ga.shinhanlife.co.kr"},    
    {"name": "삼성생명", "slug": "samsung_life", "url": "https://ga.samsunglife.com/", "badge": "Edge 전용", "edge_only": True},
    {"name": "흥국생명", "slug": "heungkuk_life", "url": "https://sales.heungkuklife.co.kr/"},
    {"name": "IBK연금보험", "slug": "ibk_pension", "url": "https://sf.ibki.co.kr/"},
    {"name": "교보생명", "slug": "kyobo_life", "url": "https://sso.kyobo.com:5443/3rdParty/certLoginFormPage.jsp?"},
    {"name": "동양생명", "slug": "tongyang_life", "url": "https://1004.myangel.co.kr/colgnsf001m.wqv?bizCode=COE0051"},
    {"name": "MetLife", "slug": "metlife", "url": "https://metplus.metlife.co.kr/"},
    {"name": "ABL생명", "slug": "abl_life", "url": "https://ga.abllife.co.kr/"},
    {"name": "DB생명", "slug": "db_life", "url": "https://ga.idblife.com/"},
    {"name": "KDB생명", "slug": "kdb_life", "url": "https://kss.kdblife.co.kr/"},
    {"name": "NH농협생명", "slug": "nh_life", "url": "https://sfa.nhlife.co.kr:8443/"},
    {"name": "BNP파리바 카디프생명", "slug": "bnp_cardif_life", "url": "https://ga.cardif.co.kr/"},
]

NON_LIFE_INSURERS = [
    {"name": "KB손해보험", "slug": "kb_insurance", "url": "https://nsales.kbinsure.co.kr/"},
    {"name": "흥국화재", "slug": "heungkuk_fire", "url": "https://sales.heungkukfire.co.kr/"},
    {"name": "한화손해보험", "slug": "hanwha_general", "url": "https://portal.hwgeneralins.com/"},
    {"name": "DB손해보험", "slug": "db_insurance", "url": "https://www.mdbins.com/"},
    {"name": "롯데손해보험", "slug": "lotte_insurance", "url": "https://lottero.lotteins.co.kr/"},
    {"name": "메리츠화재", "slug": "meritz_fire", "url": "https://nsso.meritzfire.com/LoginServer/loginFormPageMulti.jsp"},
    {"name": "삼성화재", "slug": "samsung_fire", "url": "https://login.samsungfire.com/"},
    {"name": "현대해상", "slug": "hyundai_marine", "url": "https://sp.hi.co.kr/"},
    {"name": "하나손해보험", "slug": "hana_insurance", "url": "https://sfa.saleshana.com/"},
    {"name": "AIG손해보험", "slug": "aig_insurance", "url": "https://ga.aig.co.kr/"},
    {"name": "MG손해보험", "slug": "mg_insurance", "url": "https://mganet.mggeneralins.com/", "badge": "확인 필요", "notice": True},

]


def _safe_external_url(url: str) -> str:
    parsed = urlparse(url)
    if parsed.scheme != "https" or not parsed.netloc:
        raise ValueError(f"허용되지 않은 보험사 주소입니다: {url}")
    return html.escape(url, quote=True)


def _logo_data_uri(slug: str) -> str:
    logo_path = LOGO_DIR / f"{slug}.png"
    if not logo_path.is_file():
        return ""
    encoded = base64.b64encode(logo_path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def _card(insurer: dict[str, object]) -> str:
    name = html.escape(str(insurer["name"]))
    slug = str(insurer["slug"])
    badge = insurer.get("badge")
    badge_html = f'<span class="ip-badge">{html.escape(str(badge))}</span>' if badge else ""
    logo_uri = _logo_data_uri(slug)
    logo_html = (
        f'<img class="ip-logo" src="{logo_uri}" alt="{name} 로고">'
        if logo_uri
        else f'<span class="ip-logo ip-logo-fallback">{name[:1]}</span>'
    )

    if insurer.get("edge_only"):
        safe_url = _safe_external_url(str(insurer["url"]))
        href = f"microsoft-edge:{safe_url}"
        target = "_self"
        rel = ""
        card_class = "ip-card"
        action = "Edge로 전산 열기"

    elif insurer.get("notice"):
        href = _safe_external_url(str(insurer["url"]))
        target = "_blank"
        rel = ' rel="noopener noreferrer"'
        card_class = "ip-card ip-warning-card"
        action = "안내 확인 후 전산 열기"

    else:
        href = _safe_external_url(str(insurer["url"]))
        target = "_blank"
        rel = ' rel="noopener noreferrer"'
        card_class = "ip-card"
        action = "전산 열기"

    return (
        f'<a class="{card_class}" href="{href}" target="{target}"{rel} '
        f'aria-label="{name} 전산 페이지 열기">'
        f'<span class="ip-logo-box">{logo_html}</span>'
        f'<span class="ip-card-copy"><strong>{name}</strong><small>{action}</small></span>'
        f'<span class="ip-card-side">{badge_html}<i aria-hidden="true">↗</i></span>'
        '</a>'
    )


def _section(title: str, count: int, insurers: list[dict[str, object]], section_class: str) -> str:
    cards = "".join(_card(insurer) for insurer in insurers)
    return (
        f'<section class="ip-panel {section_class}">'
        '<div class="ip-panel-head">'
        f'<div><span>INSURANCE NETWORK</span><div class="ip-panel-title">{html.escape(title)}</div></div>'
        f'<b>{count}개사</b>'
        '</div>'
        f'<div class="ip-card-grid">{cards}</div>'
        '</section>'
    )

def run() -> None:
    """보험사 전산 포털 화면을 렌더링합니다."""
    page_header(
        "업무 지원",
        "보험사 전산 포털",
        "생명보험사와 손해보험사 원수사 전산을 한 화면에서 빠르게 연결합니다.",
        "↗",
    )

    _show_mg_notice()

    st.markdown(
        """
        <style>
        .ip-layout { display:grid; grid-template-columns:minmax(0,1.08fr) minmax(0,.92fr); gap:1rem; align-items:start; }
        .ip-panel { padding:1rem; border:1px solid rgba(193,211,225,.78); border-radius:20px;
            background:linear-gradient(145deg,rgba(255,255,255,.9),rgba(248,252,255,.76));
            box-shadow:0 18px 42px rgba(35,72,100,.075); backdrop-filter:blur(16px); }
        .ip-panel-head { display:flex; align-items:flex-end; justify-content:space-between; margin:0 .15rem .85rem; }
        .ip-panel-head span { color:#4B7DA2; font-size:.56rem; font-weight:850; letter-spacing:.13em; }
        .ip-panel-title { margin:.15rem 0 0; color:##4B7DA2; font-size:1.7rem; line-height:1.2; font-weight:500; letter-spacing:-.035em; }
        .ip-panel-head b { color:#718697; font-size:.7rem; font-weight:700; }
        .ip-card-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:.52rem; }
        .ip-card { min-width:0; min-height:4.35rem; display:flex; align-items:center; gap:.68rem; padding:.58rem .62rem;
            border:1px solid rgba(207,220,231,.9); border-radius:13px; background:rgba(255,255,255,.88);
            color:#10283D !important; text-decoration:none !important; box-shadow:0 5px 14px rgba(35,72,100,.035);
            transition:transform .18s ease,border-color .18s ease,box-shadow .18s ease,background .18s ease; }
        .ip-card:hover { transform:translateY(-2px); border-color:#AFCBE2; background:#FFFFFF;
            box-shadow:0 10px 22px rgba(35,72,100,.09); }
        .ip-logo-box { flex:0 0 2.55rem; width:2.55rem; height:2.55rem; display:grid; place-items:center;
            border:1px solid #E3ECF3; border-radius:11px; background:linear-gradient(145deg,#FFF,#F5F9FC); overflow:hidden; }
        .ip-logo { display:block; width:2rem; height:2rem; object-fit:contain; }
        .ip-logo-fallback { color:#1769DC; font-weight:850; }
        .ip-card-copy { min-width:0; display:flex; flex:1; flex-direction:column; gap:.08rem; }
        .ip-card-copy strong { overflow:hidden; color:#18334A; font-size:.79rem; font-weight:780; letter-spacing:-.025em;
            text-overflow:ellipsis; white-space:nowrap; }
        .ip-card-copy small { color:#8495A3; font-size:.61rem; }
        .ip-card-side { flex:0 0 auto; display:flex; align-items:center; gap:.28rem; }
        .ip-card-side i { color:#7290A7; font-size:.8rem; font-style:normal; }
        .ip-badge { padding:.16rem .32rem; border-radius:999px; background:#EEF5FB; color:#356D96; font-size:.5rem; font-weight:800; }
        .ip-warning-card { border-color:#E7D6B1; background:linear-gradient(145deg,#FFFDF8,#FFF9EC); }
        .ip-warning-card .ip-badge { background:#FFF0CE; color:#98641D; }
        @media(max-width:1050px){.ip-layout{grid-template-columns:1fr}.ip-card-grid{grid-template-columns:repeat(3,minmax(0,1fr))}}
        @media(max-width:760px){.ip-card-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
        @media(max-width:480px){.ip-card-grid{grid-template-columns:1fr}.ip-panel{padding:.8rem}.ip-card{min-height:4rem}}
        </style>
        """,
        unsafe_allow_html=True,
    )

    life_section = _section("생명보험", len(LIFE_INSURERS), LIFE_INSURERS, "ip-life")
    non_life_section = _section("손해보험", len(NON_LIFE_INSURERS), NON_LIFE_INSURERS, "ip-non-life")
    st.markdown(f'<div class="ip-layout">{life_section}{non_life_section}</div>', unsafe_allow_html=True)
    st.caption("각 보험사의 접속 정책과 보안 프로그램에 따라 로그인 방식이 달라질 수 있습니다.")
