# modules/kakao_sender.py

import os
import time
import subprocess
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st
import pyautogui
import pyperclip


REQUIRED_COLUMNS = ["고객명", "카톡검색명", "보낼메시지", "발송상태", "발송일시"]
OPTIONAL_COLUMNS = ["고객구분", "메시지유형", "실패사유", "메모"]

DEFAULT_KAKAO_PATHS = [
    r"C:\Program Files (x86)\Kakao\KakaoTalk\KakaoTalk.exe",
    r"C:\Program Files\Kakao\KakaoTalk\KakaoTalk.exe",
]


def find_default_kakao_path() -> str:
    """일반적인 카카오톡 설치 경로를 탐색합니다."""
    for path in DEFAULT_KAKAO_PATHS:
        if os.path.exists(path):
            return path
    return ""


def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """엑셀 데이터를 앱에서 사용하기 좋게 정리합니다."""
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]

    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    for col in OPTIONAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df.fillna("")
    df["발송상태"] = df["발송상태"].replace("", "대기")

    return df


def validate_dataframe(df: pd.DataFrame) -> list[str]:
    """데이터 검증 후 경고 메시지 목록을 반환합니다."""
    warnings = []

    empty_name = df["고객명"].astype(str).str.strip().eq("").sum()
    empty_kakao = df["카톡검색명"].astype(str).str.strip().eq("").sum()
    empty_message = df["보낼메시지"].astype(str).str.strip().eq("").sum()

    if empty_name > 0:
        warnings.append(f"고객명이 비어 있는 행이 {empty_name}개 있습니다.")

    if empty_kakao > 0:
        warnings.append(f"카톡검색명이 비어 있는 행이 {empty_kakao}개 있습니다.")

    if empty_message > 0:
        warnings.append(f"보낼메시지가 비어 있는 행이 {empty_message}개 있습니다.")

    duplicated_kakao = (
        df[df["카톡검색명"].astype(str).str.strip() != ""]["카톡검색명"]
        .duplicated()
        .sum()
    )

    if duplicated_kakao > 0:
        warnings.append(
            f"카톡검색명이 중복된 행이 {duplicated_kakao}개 있습니다. 오발송 방지를 위해 확인하세요."
        )

    return warnings


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """DataFrame을 엑셀 파일 bytes로 변환합니다."""
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="고객목록")

    return output.getvalue()


def open_kakao(kakao_path: str) -> bool:
    """카카오톡 PC버전을 실행합니다."""
    try:
        if kakao_path and os.path.exists(kakao_path):
            subprocess.Popen(kakao_path)
        else:
            subprocess.Popen("KakaoTalk.exe")

        time.sleep(2)
        return True

    except Exception:
        return False


def activate_kakao_window() -> bool:
    """카카오톡 창을 찾아 활성화합니다."""
    try:
        possible_titles = ["KakaoTalk", "카카오톡"]

        for title in possible_titles:
            windows = pyautogui.getWindowsWithTitle(title)

            if windows:
                win = windows[0]

                if win.isMinimized:
                    win.restore()

                win.activate()
                time.sleep(0.5)
                return True

    except Exception:
        pass

    return False


def paste_message_to_kakao(
    kakao_search_name: str,
    message: str,
    kakao_path: str,
    delay: float = 0.6,
    search_hotkey: str = "ctrl+f",
) -> tuple[bool, str]:
    """
    카카오톡 PC버전에서 고객명을 검색하고 메시지를 붙여넣습니다.
    안전을 위해 Enter 전송은 하지 않습니다.
    """
    kakao_search_name = str(kakao_search_name).strip()
    message = str(message).strip()

    if not kakao_search_name:
        return False, "카톡검색명이 비어 있습니다."

    if not message:
        return False, "보낼메시지가 비어 있습니다."

    if not activate_kakao_window():
        opened = open_kakao(kakao_path)

        if not opened:
            return False, "카카오톡 실행에 실패했습니다. 경로를 확인하세요."

        if not activate_kakao_window():
            return False, "카카오톡 창을 찾지 못했습니다. 카카오톡을 직접 열고 다시 시도하세요."

    try:
        # 검색창 열기
        if search_hotkey == "ctrl+k":
            pyautogui.hotkey("ctrl", "k")
        else:
            pyautogui.hotkey("ctrl", "f")

        time.sleep(delay)

        # 기존 검색어 제거 후 카톡검색명 입력
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.1)

        pyperclip.copy(kakao_search_name)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(delay)

        # 검색 결과 첫 번째 항목 진입
        pyautogui.press("enter")
        time.sleep(delay + 0.5)

        # 메시지 붙여넣기
        pyperclip.copy(message)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(delay)

        return True, "카카오톡 창에 메시지를 붙여넣었습니다. 최종 전송은 직접 Enter로 확인하세요."

    except Exception as e:
        return False, f"자동화 중 오류가 발생했습니다: {e}"


def update_send_status(
    df: pd.DataFrame,
    row_index: int,
    status: str,
    reason: str = "",
) -> pd.DataFrame:
    """선택 행의 발송상태, 발송일시, 실패사유를 갱신합니다."""
    df = df.copy()

    df.loc[row_index, "발송상태"] = status

    if status == "완료":
        df.loc[row_index, "발송일시"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        df.loc[row_index, "실패사유"] = ""

    elif status == "실패":
        df.loc[row_index, "실패사유"] = reason or "사용자 실패 처리"

    elif status == "보류":
        df.loc[row_index, "실패사유"] = reason or "사용자 보류 처리"

    return df


def make_sample_excel() -> bytes:
    """샘플 엑셀 파일을 생성합니다."""
    sample_df = pd.DataFrame(
        [
            {
                "고객명": "김민지",
                "카톡검색명": "김민지 고객님",
                "고객구분": "태아보험",
                "메시지유형": "상담안내",
                "보낼메시지": (
                    "김민지 고객님 안녕하세요. 박병선입니다 😊\n"
                    "태아보험 관련해서 안내드릴 내용이 있어 연락드립니다."
                ),
                "발송상태": "대기",
                "발송일시": "",
                "실패사유": "",
                "메모": "",
            }
        ]
    )

    return dataframe_to_excel_bytes(sample_df)


def run():
    st.title("💬 카카오톡 발송 도우미")
    st.caption(
        "엑셀 고객목록을 불러와 카카오톡 PC버전에 메시지를 붙여넣는 반자동 발송 도구입니다."
    )

    # 다른 모듈과 session_state 충돌 방지를 위해 prefix 사용
    df_key = "kakao_sender_df"
    selected_key = "kakao_sender_selected_row_index"

    if df_key not in st.session_state:
        st.session_state[df_key] = None

    if selected_key not in st.session_state:
        st.session_state[selected_key] = None

    with st.sidebar:
        st.divider()
        st.subheader("💬 카톡 발송 설정")

        default_path = find_default_kakao_path()

        kakao_path = st.text_input(
            "카카오톡 실행 파일 경로",
            value=default_path,
            help="비워도 실행을 시도하지만, 실행이 안 되면 KakaoTalk.exe 경로를 직접 입력하세요.",
            key="kakao_sender_path",
        )

        search_hotkey = st.selectbox(
            "카카오톡 검색 단축키",
            ["ctrl+f", "ctrl+k"],
            index=0,
            help="PC 환경에 따라 검색창 단축키가 다를 수 있습니다.",
            key="kakao_sender_search_hotkey",
        )

        delay = st.slider(
            "자동화 동작 간격(초)",
            min_value=0.2,
            max_value=2.0,
            value=0.6,
            step=0.1,
            help="PC가 느리거나 카카오톡 반응이 늦으면 0.8~1.2초로 올리세요.",
            key="kakao_sender_delay",
        )

    uploaded_file = st.file_uploader(
        "customers.xlsx 파일을 업로드하세요",
        type=["xlsx"],
        key="kakao_sender_uploader",
    )

    if uploaded_file is not None:
        try:
            raw_df = pd.read_excel(uploaded_file)
            st.session_state[df_key] = normalize_dataframe(raw_df)
            st.session_state[selected_key] = None
            st.success("엑셀 파일을 불러왔습니다.")

        except Exception as e:
            st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")

    if st.session_state[df_key] is None:
        st.info("먼저 고객목록 엑셀 파일을 업로드하세요.")

        st.download_button(
            "샘플 엑셀 다운로드",
            data=make_sample_excel(),
            file_name="customers_sample.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="kakao_sender_sample_download",
        )

        st.markdown("### 필수 엑셀 열")
        st.code("고객명, 카톡검색명, 보낼메시지, 발송상태, 발송일시")

        return

    df = st.session_state[df_key]

    warnings = validate_dataframe(df)

    if warnings:
        with st.expander("데이터 확인 필요", expanded=True):
            for warning in warnings:
                st.warning(warning)

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("전체 고객", len(df))
    col2.metric("대기", int((df["발송상태"] == "대기").sum()))
    col3.metric("완료", int((df["발송상태"] == "완료").sum()))
    col4.metric("실패/보류", int(df["발송상태"].isin(["실패", "보류"]).sum()))

    st.divider()

    st.subheader("1. 고객 선택")

    status_options = ["전체"] + sorted(
        df["발송상태"].astype(str).replace("", "대기").unique().tolist()
    )

    default_index = status_options.index("대기") if "대기" in status_options else 0

    selected_status = st.selectbox(
        "발송상태 필터",
        status_options,
        index=default_index,
        key="kakao_sender_status_filter",
    )

    search_keyword = st.text_input(
        "고객명 검색",
        placeholder="예: 김민지",
        key="kakao_sender_search_keyword",
    )

    filtered_df = df.copy()

    if selected_status != "전체":
        filtered_df = filtered_df[filtered_df["발송상태"] == selected_status]

    if search_keyword.strip():
        keyword = search_keyword.strip()

        filtered_df = filtered_df[
            filtered_df["고객명"].astype(str).str.contains(keyword, case=False, na=False)
            | filtered_df["카톡검색명"].astype(str).str.contains(
                keyword, case=False, na=False
            )
        ]

    if filtered_df.empty:
        st.info("조건에 맞는 고객이 없습니다.")
    else:
        display_columns = [
            "고객명",
            "카톡검색명",
            "고객구분",
            "메시지유형",
            "발송상태",
            "발송일시",
            "메모",
        ]

        existing_display_columns = [
            col for col in display_columns if col in filtered_df.columns
        ]

        st.dataframe(
            filtered_df[existing_display_columns],
            use_container_width=True,
            height=280,
        )

        row_labels = [
            f"{idx} | {row['고객명']} | {row['카톡검색명']} | {row['발송상태']}"
            for idx, row in filtered_df.iterrows()
        ]

        selected_label = st.selectbox(
            "발송할 고객을 선택하세요",
            row_labels,
            key="kakao_sender_selected_label",
        )

        selected_row_index = int(selected_label.split(" | ")[0])
        st.session_state[selected_key] = selected_row_index

    st.divider()

    st.subheader("2. 메시지 미리보기")

    selected_row_index = st.session_state[selected_key]

    if selected_row_index is None:
        st.info("고객을 먼저 선택하세요.")
    else:
        row = df.loc[selected_row_index]

        left, right = st.columns([1, 1])

        with left:
            st.markdown("#### 고객 정보")
            st.write(f"**고객명:** {row['고객명']}")
            st.write(f"**카톡검색명:** {row['카톡검색명']}")
            st.write(f"**고객구분:** {row.get('고객구분', '')}")
            st.write(f"**메시지유형:** {row.get('메시지유형', '')}")
            st.write(f"**현재 상태:** {row['발송상태']}")

        with right:
            st.markdown("#### 메시지")
            edited_message = st.text_area(
                "보낼 메시지",
                value=str(row["보낼메시지"]),
                height=220,
                key=f"kakao_sender_message_{selected_row_index}",
            )

            if edited_message != str(row["보낼메시지"]):
                if st.button(
                    "수정한 메시지를 현재 데이터에 반영",
                    key="kakao_sender_apply_message",
                ):
                    st.session_state[df_key].loc[selected_row_index, "보낼메시지"] = (
                        edited_message
                    )
                    st.success("메시지를 반영했습니다.")
                    st.rerun()

        st.divider()

        st.subheader("3. 카카오톡에 붙여넣기")
        st.warning(
            "안전상 이 앱은 메시지를 붙여넣기까지만 합니다. 최종 전송 Enter는 직접 눌러 확인하세요."
        )

        col_a, col_b, col_c = st.columns(3)

        with col_a:
            if st.button(
                "카카오톡 열고 메시지 붙여넣기",
                type="primary",
                use_container_width=True,
                key="kakao_sender_paste",
            ):
                success, msg = paste_message_to_kakao(
                    kakao_search_name=row["카톡검색명"],
                    message=edited_message,
                    kakao_path=kakao_path,
                    delay=delay,
                    search_hotkey=search_hotkey,
                )

                if success:
                    st.success(msg)
                else:
                    st.error(msg)

        with col_b:
            if st.button(
                "발송 완료 처리",
                use_container_width=True,
                key="kakao_sender_complete",
            ):
                st.session_state[df_key] = update_send_status(
                    st.session_state[df_key],
                    selected_row_index,
                    "완료",
                )
                st.success("발송 완료로 기록했습니다.")
                st.rerun()

        with col_c:
            if st.button(
                "보류 처리",
                use_container_width=True,
                key="kakao_sender_hold",
            ):
                st.session_state[df_key] = update_send_status(
                    st.session_state[df_key],
                    selected_row_index,
                    "보류",
                    "사용자 보류 처리",
                )
                st.info("보류로 기록했습니다.")
                st.rerun()

        fail_reason = st.text_input(
            "실패 처리 사유",
            placeholder="예: 카톡검색명 불일치, 대화방 미확인 등",
            key="kakao_sender_fail_reason",
        )

        if st.button("실패 처리", key="kakao_sender_fail"):
            st.session_state[df_key] = update_send_status(
                st.session_state[df_key],
                selected_row_index,
                "실패",
                fail_reason,
            )
            st.error("실패로 기록했습니다.")
            st.rerun()

    st.divider()

    st.subheader("4. 결과 엑셀 다운로드")

    result_bytes = dataframe_to_excel_bytes(st.session_state[df_key])

    st.download_button(
        "현재 상태를 엑셀로 다운로드",
        data=result_bytes,
        file_name=f"customers_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="kakao_sender_result_download",
    )

    st.caption(
        "주의: Streamlit Cloud나 서버 배포 환경에서는 사용자의 PC 카카오톡을 제어할 수 없습니다. "
        "이 기능은 카카오톡 PC버전이 설치된 로컬 Windows PC에서 실행해야 합니다."
    )