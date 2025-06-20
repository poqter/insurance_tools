import streamlit as st
import openpyxl
import re
from io import BytesIO
from datetime import datetime

def run():
    st.header("📑 보장 분석 도우미")

# ✅ 기본 템플릿 파일 로드 (다운로드 버튼용)
with open("print.xlsx", "rb") as f:
    default_template_data = f.read()

# ✅ 사이드바 안내문 + 다운로드 버튼 + 제작자 정보
st.sidebar.markdown("### 📘 사용 방법 안내")
st.sidebar.markdown("""
1. 컨설팅보장분석.xlsx 업로드  
2. (선택) 개인용 보장분석 폼.xlsx 업로드  
3. 분석 후 **결과 파일 다운로드**

📌 참고:  
- print.xlsx 미첨부 시, **기본 폼 자동 사용**  
- 지원 파일: .xlsx (엑셀 전용)
""")
st.sidebar.markdown("📝 **기본 폼을 수정하려면 아래 파일을 받아 수정 후 업로드하세요.**")

st.sidebar.download_button(
    label="📥 기본 폼(print.xlsx) 다운로드",
    data=default_template_data,
    file_name="print.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.sidebar.markdown("---")
st.sidebar.markdown("👨‍💻 **제작자:** 비전본부 드림지점 박병선 팀장")  
st.sidebar.markdown("🗓️ **버전:** v1.0.0")  
st.sidebar.markdown("📅 **최종 업데이트:** 2025-06-13")

# ✅ 제목 및 설명
st.title("📊 보장 분석 도우미")
st.write("컨설팅보장분석.xlsx 파일을 업로드하면 자동으로 결과 분석 파일이 생성됩니다.")

# ✅ 엑셀 파일 업로드
uploaded_main = st.file_uploader("⬆️ 컨설팅보장분석.xlsx 파일을 업로드하세요", type=["xlsx"])
uploaded_print = st.file_uploader("🖨️ (선택) 개인용 보장분석 폼.xlsx 파일을 업로드하세요", type=["xlsx"])

# ✅ print.xlsx 로드
try:
    if uploaded_print:
        print_wb = openpyxl.load_workbook(uploaded_print)
        st.info("✅ 업로드한 print.xlsx를 사용합니다.")
    else:
        print_wb = openpyxl.load_workbook("print.xlsx")
        st.info("📌 기본 내장된 print.xlsx를 사용합니다.")
    print_ws = print_wb.active
except Exception as e:
    st.error(f"❌ print.xlsx 파일을 열 수 없습니다: {e}")
    st.stop()

# ✅ main.xlsx 처리
if uploaded_main:
    try:
        main_wb = openpyxl.load_workbook(uploaded_main, data_only=True)
        main_ws1 = main_wb["계약사항"]
        main_ws2 = main_wb["보장사항"]

        # 계약사항 시트: J, K, L열 → D10~AD10, D8~AD9
        for idx in range(27):
            print_ws.cell(row=10, column=4 + idx).value = main_ws1[f"J{9+idx}"].value
        for row_offset, col in enumerate(['K', 'L']):
            for idx in range(27):
                print_ws.cell(row=8 + row_offset, column=4 + idx).value = main_ws1[f"{col}{9+idx}"].value

        # 보장사항 시트: F7~AD7 숫자만 추출하여 D7~Z7에 붙여넣기
        for col in range(6, 30):  # F(6) to AD(30)
            raw_value = main_ws2.cell(row=7, column=col).value
            if raw_value is not None:
                number = re.sub(r"[^\d]", "", str(raw_value))
                print_ws.cell(row=7, column=col - 2).value = int(number) if number else ""

        # 보장사항 시트: 2~6행, 9~45행은 원래대로 복사
        for row in range(2, 7):  # 2~6행
            for col in range(6, 30):
                print_ws.cell(row=row, column=col - 2).value = main_ws2.cell(row=row, column=col).value

        for row in range(9, 46):  # 9~45행
            for col in range(6, 30):
                print_ws.cell(row=row + 3, column=col - 2).value = main_ws2.cell(row=row, column=col).value

        # 제목 생성
        name_prefix = (main_ws1["B2"].value or "고객")[:3]
        detail_text = main_ws1["D2"].value or ""
        print_ws["A1"] = f"{name_prefix}님의 기존 보험 보장 분석 {detail_text}"

        # 결과 저장 및 다운로드
        today_str = datetime.today().strftime("%Y%m%d")
        filename = f"{name_prefix}님의_보장분석_{today_str}.xlsx"
        output_excel = BytesIO()
        print_wb.save(output_excel)
        output_excel.seek(0)

        st.success("✅ 분석이 완료되었습니다.")
        st.download_button(
            label="📥 결과 파일 다운로드",
            data=output_excel,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"⚠️ 오류가 발생했습니다: {str(e)}")
")

   

