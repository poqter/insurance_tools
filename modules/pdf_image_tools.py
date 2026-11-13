"""PDF/image editing in the existing insurance_tools authenticated application."""
from pathlib import Path
import hashlib
import streamlit as st
import streamlit.components.v1 as components

ALLOWED_ROLES = frozenset({"Admin", "Manager1", "Basic", "Crew", "Dream"})
_component = components.declare_component(
    "insurance_pdf_tools", path=str(Path(__file__).with_name("pdf_tools"))
)


def run() -> None:
    role = st.session_state.get("login_user")
    if not st.session_state.get("password_correct") or role not in ALLOWED_ROLES:
        st.error("로그인과 도구 이용 권한을 확인해 주세요.")
        return
    st.title("PDF·이미지 도구")
    st.caption("파일 추가 → 페이지 편집 → 저장 전 확인 → 다운로드")
    # All document data stays inside the component; no server/session file payloads.
    # Role changes recreate the component, and logout clears the application state.
    scope = hashlib.sha256(str(role).encode()).hexdigest()[:16]
    _component(key="insurance_pdf_tools_" + scope, default=None)
