# 화랑 WORKS 모던 UI 버전

기존 계산 및 엑셀 생성 로직을 유지하면서 홈, 로그인, 사이드바와 8개 업무 프로그램에 공통 SaaS 디자인을 적용한 전체 파일입니다.

## 파일 구성

- `app.py`: 로그인, 계정 권한, 통합 홈과 프로그램 이동
- `modules/ui_components.py`: 공통 색상, 글꼴, 버튼, 카드, 표, 탭과 프로그램 헤더
- `modules/*.py`: 8개 업무 프로그램
- `requirements.txt`: 배포에 필요한 Python 패키지

## 기존 배포에 적용

1. 기존 프로젝트를 별도로 백업합니다.
2. 이 폴더의 `app.py`와 `modules` 폴더를 프로젝트에 복사합니다.
3. 기존 프로젝트에서 사용하는 `print.xlsx`와 `.streamlit/secrets.toml`은 그대로 유지합니다.
4. `streamlit run app.py`로 실행합니다.

계정 비밀번호는 기존과 동일하게 `st.secrets["passwords"]`에서 불러옵니다.
