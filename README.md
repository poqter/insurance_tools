# 화랑 WORKSPACE 통합 디자인 버전

기존 계산·판정·파일 처리 로직을 유지하면서 로그인, 통합 홈, 사이드바와 전체 업무 기능에 화랑 WORKSPACE 공통 디자인을 적용한 배포용 프로젝트입니다.

## 파일 구성

- `app.py`: 로그인, 계정 권한, 통합 홈과 프로그램 이동
- `modules/ui_components.py`: 공통 색상, 글꼴, 단색 아이콘, 안내 카드, 기능 헤더와 하단 정보
- `modules/*.py`: 고객 상담·실적 관리·업무 지원 기능
- `assets/`: Pretendard 글꼴과 보험회사 로고
- `requirements.txt`: 배포에 필요한 Python 패키지

## 기존 배포에 적용

1. 기존 프로젝트를 별도로 백업합니다.
2. 이 폴더의 `app.py`와 `modules` 폴더를 프로젝트에 복사합니다.
3. 기존 프로젝트에서 사용하는 `.streamlit/secrets.toml`은 별도로 유지합니다.
4. `streamlit run app.py`로 실행합니다.

계정 비밀번호는 기존과 동일하게 `st.secrets["passwords"]`에서 불러옵니다.

## 브랜드 적용 기준

- 프로그램 화면: `Planned & Built by 박병선`
- 고객용 PDF·엑셀: `H │ 화랑 WORKSPACE`
- 기능별 버전과 업데이트 날짜는 각 기능에서 개별 관리합니다.
