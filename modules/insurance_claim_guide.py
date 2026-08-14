from __future__ import annotations

import hashlib
import html
import re
from dataclasses import asdict, dataclass
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Iterable

import pandas as pd
import pdfplumber
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

try:
    from .ui_components import page_header, section_intro
except ImportError:  # 단독 파일 점검용
    from ui_components import page_header, section_intro


GUIDE_VERSION = "1.1.0"
GUIDE_STANDARD_DATE = "2026.08"
STATE_PREFIX = "cg_"


@dataclass(frozen=True)
class DocumentRule:
    name: str
    required_info: str
    group: str = "병원 발급"
    level: str = "기본 준비"
    default_selected: bool = True


@dataclass(frozen=True)
class CoverageRow:
    company: str
    product: str
    category: str
    coverage: str
    amount: str
    contract_date: str
    expiry_date: str
    source_page: int
    extraction_status: str = "정상 추출"


CLAIM_GROUPS = {
    "의료비·기본 치료": [
        "실손 통원", "실손 입원", "약제비", "입원일당", "수술",
        "응급실", "간병인", "간호간병통합", "기타 치료비",
    ],
    "주요 진단": [
        "암", "뇌질환", "심장질환", "골절", "화상", "치매", "장기요양", "기타 진단",
    ],
    "암 치료": [
        "항암약물", "표적항암", "항암방사선", "양성자·세기조절", "CAR-T", "중입자치료",
    ],
    "사고·특수 청구": [
        "일반 상해", "교통사고", "운전자비용", "후유장해", "사망", "태아·출산", "치아", "배상책임",
    ],
}


COMMON_DOCUMENTS = [
    DocumentRule("보험금 청구서", "보험회사별 양식, 청구인 서명", "직접 준비", default_selected=False),
    DocumentRule("개인·신용정보 처리동의서", "청구인 자필서명", "직접 준비", default_selected=False),
    DocumentRule("신분증 사본", "청구인 기준", "직접 준비"),
    DocumentRule("보험금 수령 계좌정보", "청구인 또는 보험수익자 계좌", "직접 준비"),
]


DOC_RULES: dict[str, list[DocumentRule]] = {
    "실손 통원": [
        DocumentRule("진료비 계산서·영수증", "환자명, 진료일, 본인부담금"),
        DocumentRule("진료비 세부내역서", "급여·비급여 치료내역"),
        DocumentRule("처방전", "처방받은 경우, 진단명 또는 질병분류코드"),
        DocumentRule("통원확인서", "진단명, 질병분류코드, 통원일", level="진단정보가 부족하거나 보험회사 요청 시", default_selected=False),
    ],
    "실손 입원": [
        DocumentRule("진료비 계산서·영수증", "환자명, 입원기간, 본인부담금"),
        DocumentRule("진료비 세부내역서", "급여·비급여 치료내역"),
        DocumentRule("입퇴원확인서", "진단명, 질병분류코드, 입원일·퇴원일"),
        DocumentRule("진단서", "진단명, 질병분류코드, 진단일", level="대체서류 확인"),
    ],
    "약제비": [
        DocumentRule("약제비 계산서·영수증", "환자명, 처방일, 약제비"),
        DocumentRule("처방전", "질병분류코드, 처방 의료기관"),
    ],
    "입원일당": [
        DocumentRule("입퇴원확인서", "진단명, 질병분류코드, 입원일·퇴원일"),
        DocumentRule("중환자실·병실 확인자료", "병실 종류와 이용기간", level="해당 시"),
    ],
    "수술": [
        DocumentRule("수술확인서", "진단명, 질병분류코드, 정확한 수술명, 수술일"),
        DocumentRule("수술기록지", "수술방법과 수술 부위", level="보험사 요청 가능"),
        DocumentRule("조직병리검사 결과지", "최종 병리진단", level="조직검사한 경우"),
    ],
    "응급실": [
        DocumentRule("응급실 진료확인서", "진단명, 질병분류코드, 내원일, 응급환자 여부"),
        DocumentRule("응급실 진료기록", "내원경위와 처치내용", level="보험사 요청 가능"),
    ],
    "간병인": [
        DocumentRule("간병인 사용확인서", "간병기간, 간병인 정보"),
        DocumentRule("간병비 영수증", "지급일, 지급금액, 수령인"),
        DocumentRule("입퇴원확인서", "진단명, 입원일·퇴원일"),
    ],
    "간호간병통합": [
        DocumentRule("간호·간병통합서비스 사용확인서", "사용병동과 이용기간"),
        DocumentRule("입퇴원확인서", "진단명, 입원일·퇴원일"),
    ],
    "기타 치료비": [
        DocumentRule("치료확인서", "진단명, 질병분류코드, 치료명, 치료일"),
        DocumentRule("진료비 세부내역서", "급여·비급여 치료내역"),
        DocumentRule("진료기록지", "치료 목적과 시행내용", level="보험사 요청 가능"),
    ],
    "암": [
        DocumentRule("진단서", "암 진단명, 질병분류코드, 확정진단일"),
        DocumentRule("조직병리검사 결과지", "최종 병리진단과 조직검사 결과"),
        DocumentRule("영상검사 결과지", "조직검사가 불가능한 경우 진단 근거", level="해당 시"),
    ],
    "뇌질환": [
        DocumentRule("진단서", "진단명, 질병분류코드, 확정진단일"),
        DocumentRule("CT·MRI 영상검사 판독지", "진단 근거가 되는 영상검사 결과"),
        DocumentRule("뇌혈관조영술 결과지", "혈관 병변과 검사 결과", level="해당 시"),
    ],
    "심장질환": [
        DocumentRule("진단서", "진단명, 질병분류코드, 확정진단일"),
        DocumentRule("심장검사 결과지", "심전도·심초음파·심장효소 등 진단 근거"),
        DocumentRule("관상동맥조영술 결과지", "혈관 병변과 시술내용", level="해당 시"),
    ],
    "골절": [
        DocumentRule("진단서 또는 통원확인서", "골절 진단명, 질병분류코드, 골절 부위"),
        DocumentRule("영상검사 판독지", "골절 부위와 검사 결과"),
        DocumentRule("깁스·부목 치료확인서", "치료 종류와 시행일", level="해당 시"),
    ],
    "화상": [
        DocumentRule("진단서", "화상 진단명, 질병분류코드, 화상 정도와 부위"),
        DocumentRule("진료기록지", "화상 깊이·범위와 치료내용", level="보험사 요청 가능"),
    ],
    "치매": [
        DocumentRule("진단서", "치매 진단명, 질병분류코드, 진단일"),
        DocumentRule("인지기능검사 결과지", "검사명, 점수, 시행일"),
        DocumentRule("진료기록지", "임상치매척도 등 진단 근거"),
        DocumentRule("대리청구 관계서류", "지정대리청구 여부와 관계 확인", "직접 준비", "해당 시"),
    ],
    "장기요양": [
        DocumentRule("장기요양인정서", "인정등급, 인정일, 유효기간", "직접 준비"),
        DocumentRule("개인별 장기요양이용계획서", "급여 종류와 이용계획", "직접 준비"),
    ],
    "기타 진단": [
        DocumentRule("진단서", "정확한 진단명, 질병분류코드, 진단일"),
        DocumentRule("진료기록지", "진단 및 치료 경과", level="보험사 요청 가능", default_selected=False),
        DocumentRule("사고사실 확인자료", "사고일자와 사고내용", "상황별 추가", "사고로 인한 진단인 경우", False),
    ],
    "항암약물": [
        DocumentRule("항암치료확인서", "암 진단명, 항암제명, 투여일, 치료 목적"),
        DocumentRule("투약기록", "약제명과 투여일", level="보험사 요청 가능"),
        DocumentRule("진료비 세부내역서", "항암제와 치료내역"),
    ],
    "표적항암": [
        DocumentRule("표적항암치료확인서", "암 진단명, 약제명, 투여일, 허가치료 여부"),
        DocumentRule("투약기록", "정확한 약제명과 투여일"),
        DocumentRule("유전자·바이오마커 검사결과", "표적치료 적용 근거", level="해당 시"),
        DocumentRule("진료비 세부내역서", "표적항암제와 치료내역"),
    ],
    "항암방사선": [
        DocumentRule("방사선치료확인서", "암 진단명, 치료 종류, 치료기간, 치료횟수"),
        DocumentRule("진료비 세부내역서", "방사선치료 내역"),
        DocumentRule("방사선 치료기록지", "치료방법, 치료기간, 치료횟수", level="추가 확인이 필요한 경우", default_selected=False),
    ],
    "양성자·세기조절": [
        DocumentRule("방사선치료확인서", "양성자 또는 세기조절 치료기법, 치료기간, 횟수"),
        DocumentRule("진료비 세부내역서", "치료기법과 치료내역"),
    ],
    "CAR-T": [
        DocumentRule("CAR-T 치료확인서", "약제명, 투여일, 허가치료 여부"),
        DocumentRule("투약기록", "약제명과 투여일"),
    ],
    "중입자치료": [
        DocumentRule("중입자방사선치료확인서", "치료기법, 치료기간, 치료횟수"),
        DocumentRule("진료비 세부내역서", "중입자치료 내역"),
    ],
    "일반 상해": [
        DocumentRule("상해사고 경위서", "사고일자, 장소, 사고과정, 다친 부위", "직접 준비"),
        DocumentRule("초진기록지", "최초 내원일과 사고경위", level="보험사 요청 가능"),
    ],
    "교통사고": [
        DocumentRule("자동차보험 지급결의서", "보험회사, 사고번호, 사고일, 부상등급, 지급내역", "상황별 추가"),
        DocumentRule("교통사고 사실확인원", "사고일자, 사고내용, 당사자", "상황별 추가"),
        DocumentRule("진단서", "진단명, 질병분류코드, 진단일"),
    ],
    "운전자비용": [
        DocumentRule("사고사실확인서", "사고일, 사고내용, 운전 여부", "상황별 추가"),
        DocumentRule("형사합의서·지급증빙", "합의내용과 실제 지급금액", "상황별 추가", "교통사고처리지원금"),
        DocumentRule("약식명령문·판결문", "벌금액과 사건내용", "상황별 추가", "벌금 청구"),
        DocumentRule("변호사 선임계약서·영수증", "선임일과 실제 지급금액", "상황별 추가", "변호사선임비용"),
    ],
    "후유장해": [
        DocumentRule("후유장해진단서", "장해 원인·부위·상태, 증상 고정, 평가 근거"),
        DocumentRule("검사결과·진료기록", "장해 판단의 의학적 근거", level="보험사 요청 가능"),
    ],
    "사망": [
        DocumentRule("사망진단서 또는 사체검안서", "사망일, 사망원인"),
        DocumentRule("기본증명서·가족관계증명서", "사망사실과 수익자 관계", "직접 준비"),
        DocumentRule("수익자 신분·계좌서류", "보험수익자 기준", "직접 준비"),
        DocumentRule("사고사실확인서", "사고일과 사고원인", "상황별 추가", "재해사망인 경우"),
    ],
    "태아·출산": [
        DocumentRule("출생증명서", "출생일, 출생체중, 임신주수"),
        DocumentRule("입퇴원확인서", "진단명, 입원일·퇴원일, 신생아중환자실 기간"),
        DocumentRule("진단서", "진단명, 질병분류코드"),
        DocumentRule("가족관계증명서", "출생아와 청구인 관계", "직접 준비"),
        DocumentRule("유산진단서", "유산 진단명, 진단일", level="유산 청구 시", default_selected=False),
        DocumentRule("사산증명서", "사산 사실과 일자", level="사산 청구 시", default_selected=False),
    ],
    "치아": [
        DocumentRule("치과치료확인서", "치아번호, 진단명, 치료명, 진단일·치료일"),
        DocumentRule("치과진료기록 사본", "치료 전후 상태와 치료내용"),
        DocumentRule("치과 방사선 사진", "치료 전후 치아상태", level="보험사 요청 가능"),
    ],
    "배상책임": [
        DocumentRule("사고경위서", "사고일시, 장소, 원인, 피해내용", "직접 준비"),
        DocumentRule("피해사실·손해액 확인자료", "수리견적서·영수증·사진 등", "상황별 추가"),
        DocumentRule("당사자 관계 확인자료", "피보험자와 피해자의 관계", "상황별 추가"),
    ],
}


MATCH_RULES: dict[str, dict[str, list[str]]] = {
    "실손 통원": {"direct": ["통원의료비", "외래의료비", "통원실손", "외래실손"], "related": ["처방조제", "비급여도수", "체외충격파", "비급여주사", "비급여mri"]},
    "실손 입원": {"direct": ["입원의료비", "입원실손"], "related": ["비급여도수", "체외충격파", "비급여주사", "비급여mri"]},
    "약제비": {"direct": ["처방조제", "약제의료비", "약제비"], "related": []},
    "입원일당": {"direct": ["입원일당", "입원급여", "입원생활비"], "related": ["중환자실", "1인실", "2~3인실", "상급종합병원입원", "종합병원입원", "간병인사용입원", "간호간병통합"]},
    "수술": {"direct": ["수술비", "수술급여", "수술특약", "종수술", "재해수술"], "related": ["시술비", "이식수술", "로봇수술"]},
    "응급실": {"direct": ["응급실내원", "응급치료"], "related": []},
    "간병인": {"direct": ["간병인사용", "간병인지원", "간병비"], "related": ["입원일당"]},
    "간호간병통합": {"direct": ["간호간병통합"], "related": ["입원일당"]},
    "기타 치료비": {"direct": ["치료비", "치료지원금"], "related": ["검사비", "재활치료", "투석치료"]},
    "암": {"direct": ["암진단", "암치료자금", "유사암", "소액암", "고액암", "특정암", "재진단암", "경계성종양", "상피내암", "기타피부암", "갑상선암"], "related": ["암수술", "암입원", "암통원", "암주요치료", "항암", "납입면제"]},
    "뇌질환": {"direct": ["뇌혈관질환진단", "뇌졸중진단", "뇌출혈진단", "뇌경색진단", "중증질환뇌혈관"], "related": ["뇌혈관질환수술", "혈전용해", "혈전제거", "납입면제"]},
    "심장질환": {"direct": ["허혈성심장질환진단", "허혈심장질환진단", "급성심근경색", "협심증진단", "심부전진단", "부정맥진단", "중증질환심장"], "related": ["허혈성심장질환수술", "관상동맥", "스텐트", "풍선혈관", "납입면제"]},
    "골절": {"direct": ["골절진단", "골절급여"], "related": ["5대골절", "특정골절", "골절수술", "골절철심", "골절부목", "깁스치료", "상해수술", "상해입원"]},
    "화상": {"direct": ["화상진단", "중대한화상", "중대화상", "화상및부식"], "related": ["화상수술", "상해수술", "상해입원"]},
    "치매": {"direct": ["치매진단", "경증치매", "중등도치매", "중증치매"], "related": ["치매간병", "치매생활자금", "납입면제"]},
    "장기요양": {"direct": ["장기요양", "재가급여", "시설급여"], "related": ["간병생활자금"]},
    "기타 진단": {"direct": ["진단비", "진단급여"], "related": []},
    "항암약물": {"direct": ["항암약물", "항암방사선약물"], "related": ["암주요치료"]},
    "표적항암": {"direct": ["표적항암"], "related": ["항암약물", "암주요치료"]},
    "항암방사선": {"direct": ["항암방사선"], "related": ["암주요치료"]},
    "양성자·세기조절": {"direct": ["양성자", "세기조절방사선"], "related": ["항암방사선"]},
    "CAR-T": {"direct": ["car-t", "cart항암", "카티항암"], "related": ["항암약물"]},
    "중입자치료": {"direct": ["중입자"], "related": ["항암방사선"]},
    "일반 상해": {"direct": ["일반상해", "상해진단", "재해진단"], "related": ["상해수술", "상해입원", "상해후유장해", "상해의료비"]},
    "교통사고": {"direct": ["자동차사고부상", "교통상해", "자동차사고입원"], "related": ["상해수술", "상해입원", "상해후유장해"]},
    "운전자비용": {"direct": ["교통사고처리지원금", "변호사선임", "교통사고벌금", "운전자벌금", "면허정지", "면허취소"], "related": ["자동차사고부상"]},
    "후유장해": {"direct": ["후유장해", "후유장애"], "related": ["납입면제"]},
    "사망": {"direct": ["질병사망", "상해사망", "재해사망", "교통상해사망", "사망보험금"], "related": []},
    "태아·출산": {"direct": ["저체중아", "신생아", "태아", "선천이상", "출산", "산모"], "related": ["신생아중환자실"]},
    "치아": {"direct": ["치아보철", "치아보존", "임플란트", "브리지", "틀니", "크라운", "충전치료", "치수치료", "영구치발치"], "related": ["치아파절"]},
    "배상책임": {"direct": ["배상책임"], "related": ["법률비용"]},
}


LIMIT_TERMS = {
    "유사암제외": "암 종류 확인",
    "치아파절제외": "치아파절 제외 여부 확인",
    "요양병원제외": "의료기관 종류 확인",
    "최초1회": "이전 지급 여부 확인",
    "연간1회": "같은 연도의 과거 청구 확인",
    "급여": "급여 치료 여부 확인",
    "비급여": "비급여 치료 여부 확인",
    "상급종합병원": "의료기관 종별 확인",
    "중환자실": "중환자실 이용 여부 확인",
    "운전자": "사고 당시 운전 여부 확인",
    "갱신": "사고일 당시 계약 유지 여부 확인",
}


INSURER_ALIASES = {
    "DB손보": "DB손해보험", "KB손보": "KB손해보험", "NH손보": "NH농협손해보험",
    "농협손해보험": "NH농협손해보험", "신한생명": "신한라이프", "우체국": "우체국보험",
}


INSURER_PATTERN = re.compile(
    r"DB손해보험|DB손보|KB손해보험|KB손보|현대해상|메리츠화재|한화손해보험|흥국화재|"
    r"삼성화재|롯데손해보험|MG손해보험|NH농협손해보험|농협손해보험|NH손보|캐롯손해보험|"
    r"신한라이프|신한생명|한화생명|교보생명|삼성생명|라이나생명|ABL생명|AIA생명|동양생명|"
    r"흥국생명|NH농협생명|미래에셋생명|KDB생명|하나생명|IBK연금보험|처브라이프|"
    r"우체국보험|우체국|새마을금고|수협"
)


def normalize_text(value: str) -> str:
    value = str(value or "").lower()
    value = value.replace("ㆍ", "").replace("·", "").replace("–", "-")
    return re.sub(r"[\s()\[\]{},._/\\:]+", "", value)


def normalize_company(value: str) -> str:
    value = str(value or "").strip()
    return INSURER_ALIASES.get(value, value)


def parse_amount(value: str) -> str:
    raw = str(value or "").strip()
    if raw in {"", "-", "-원"}:
        return "확인 필요"
    match = re.search(r"[\d,]+(?:\.\d+)?", raw)
    return f"{match.group(0)}만원" if match else raw


def extract_pdf(pdf_bytes: bytes) -> dict:
    pages_text: list[str] = []
    pages_words: list[tuple[float, list[dict]]] = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            pages_text.append(page.extract_text(x_tolerance=2, y_tolerance=3, layout=True) or "")
            pages_words.append((float(page.width), page.extract_words(x_tolerance=1, y_tolerance=2) or []))

    first_text = "\n".join(pages_text[:2])
    customer_match = re.search(r"([가-힣]{2,5})님을\s*위한", first_text)
    report_match = re.search(r"작성일자\s*(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})", first_text)
    customer = customer_match.group(1) if customer_match else "확인 필요"
    report_date = (
        f"{report_match.group(1)}.{int(report_match.group(2)):02d}.{int(report_match.group(3)):02d}"
        if report_match else "확인 필요"
    )

    rows: list[CoverageRow] = []
    for page_no, (page_width, words) in enumerate(pages_words, start=1):
        if not words:
            continue
        scale = 595.28 / page_width if page_width else 1.0
        line_groups: list[list[dict]] = []
        for word in sorted(words, key=lambda item: (float(item["top"]), float(item["x0"]))):
            if not line_groups or abs(float(word["top"]) - float(line_groups[-1][0]["top"])) > 2.2:
                line_groups.append([word])
            else:
                line_groups[-1].append(word)

        current_category = ""
        for line_words in line_groups:
            fields = {"category": [], "company": [], "product": [], "coverage": [], "amount": [], "contract": [], "expiry": []}
            for word in sorted(line_words, key=lambda item: float(item["x0"])):
                x = float(word["x0"]) * scale
                text = str(word["text"]).strip()
                if x < 95:
                    fields["category"].append(text)
                elif x < 168:
                    fields["company"].append(text)
                elif x < 303:
                    fields["product"].append(text)
                elif x < 428:
                    fields["coverage"].append(text)
                elif x < 473:
                    fields["amount"].append(text)
                elif x < 528:
                    fields["contract"].append(text)
                else:
                    fields["expiry"].append(text)

            category_text = " ".join(fields["category"]).strip()
            company_text = " ".join(fields["company"]).strip()
            if category_text and not company_text and len(category_text) <= 35:
                current_category = category_text
            insurer_match = INSURER_PATTERN.search(company_text)
            if not insurer_match:
                continue
            product = " ".join(fields["product"]).strip()
            coverage = " ".join(fields["coverage"]).strip()
            amount_raw = " ".join(fields["amount"]).strip()
            contract_date = " ".join(fields["contract"]).strip()
            expiry_date = " ".join(fields["expiry"]).strip()
            if not product or not coverage or not re.fullmatch(r"\d{4}[-.]\d{1,2}(?:[-.]\d{1,2})?", contract_date):
                continue
            if not re.fullmatch(r"\d{4}[-.]\d{1,2}(?:[-.]\d{1,2})?|종신", expiry_date):
                expiry_date = "확인 필요"
            status = "담보명 잘림 가능성" if coverage.endswith(("(", "제", "갱", "지", "수")) else "정상 추출"
            rows.append(
                CoverageRow(
                    company=normalize_company(insurer_match.group(0)),
                    product=product,
                    category=category_text or current_category,
                    coverage=coverage,
                    amount=parse_amount(amount_raw),
                    contract_date=contract_date,
                    expiry_date=expiry_date,
                    source_page=page_no,
                    extraction_status=status,
                )
            )

    deduped: list[CoverageRow] = []
    seen: set[tuple] = set()
    for row in rows:
        key = (row.company, row.product, row.coverage, row.amount, row.contract_date, row.expiry_date)
        if key not in seen:
            seen.add(key)
            deduped.append(row)

    return {
        "customer": customer,
        "report_date": report_date,
        "coverages": [asdict(row) for row in deduped],
        "page_count": len(pages_text),
    }


def match_coverages(coverages: list[dict], selected_claims: list[str]) -> pd.DataFrame:
    matched: dict[tuple, dict] = {}
    for row in coverages:
        searchable = normalize_text(f"{row.get('category', '')} {row.get('coverage', '')}")
        for claim in selected_claims:
            rules = MATCH_RULES.get(claim, {"direct": [], "related": []})
            direct_hits = [term for term in rules["direct"] if normalize_text(term) in searchable]
            related_hits = [term for term in rules["related"] if normalize_text(term) in searchable]
            if not direct_hits and not related_hits:
                continue

            relation = "직접 관련" if direct_hits else "함께 확인"
            hit = (direct_hits or related_hits)[0]
            checks = []
            for term, note in LIMIT_TERMS.items():
                if normalize_text(term) in searchable:
                    checks.append(note)
            if checks and relation == "직접 관련":
                relation = "조건부 관련"
            note = " · ".join(dict.fromkeys(checks)) or (
                f"{claim} 청구와 직접 관련" if relation == "직접 관련" else f"{claim} 상황에서 함께 확인"
            )

            key = (
                row.get("company"), row.get("product"), row.get("coverage"), row.get("amount"), row.get("contract_date")
            )
            priority = {"직접 관련": 3, "조건부 관련": 2, "함께 확인": 1}
            candidate = {
                "포함": relation != "함께 확인",
                "보험회사": row.get("company", "확인 필요"),
                "상품명": row.get("product", "확인 필요"),
                "관련 담보": row.get("coverage", "확인 필요"),
                "가입금액": row.get("amount", "확인 필요"),
                "분류": relation,
                "확인사항": note,
                "추출상태": row.get("extraction_status", "정상 추출"),
                "계약일": row.get("contract_date", ""),
                "만기일": row.get("expiry_date", ""),
                "원본쪽": row.get("source_page", ""),
                "매칭근거": hit,
            }
            if key not in matched or priority[relation] > priority[matched[key]["분류"]]:
                matched[key] = candidate

    columns = ["포함", "보험회사", "상품명", "관련 담보", "가입금액", "분류", "확인사항", "추출상태", "계약일", "만기일", "원본쪽", "매칭근거"]
    if not matched:
        return pd.DataFrame(columns=columns)
    df = pd.DataFrame(matched.values(), columns=columns)
    relation_order = pd.Categorical(df["분류"], ["직접 관련", "조건부 관련", "함께 확인"], ordered=True)
    return df.assign(_order=relation_order).sort_values(["보험회사", "_order", "상품명"]).drop(columns="_order").reset_index(drop=True)


def refine_matches_with_answers(df: pd.DataFrame, answers: dict) -> pd.DataFrame:
    """추가 질문 답변과 맞지 않는 넓은 매칭은 '함께 확인'으로 내린다."""
    if df.empty:
        return df
    result = df.copy()

    def downgrade(mask: pd.Series) -> None:
        result.loc[mask, "분류"] = "함께 확인"
        result.loc[mask, "포함"] = False

    cause = answers.get("death_cause")
    if cause in {"질병", "재해·상해", "교통사고"}:
        normalized = result["관련 담보"].map(normalize_text)
        if cause == "질병":
            keep = normalized.str.contains("질병사망")
        elif cause == "재해·상해":
            keep = normalized.str.contains("상해사망|재해사망", regex=True) & ~normalized.str.contains("교통")
        else:
            keep = normalized.str.contains("교통.*사망|자동차.*사망", regex=True)
        death_rows = normalized.str.contains("사망")
        downgrade(death_rows & ~keep)

    birth_claims = set(answers.get("birth_claims", []))
    if birth_claims:
        keyword_map = {
            "저체중아": ["저체중", "출생체중"], "신생아 질환·입원": ["신생아", "신생아입원"],
            "신생아중환자실": ["신생아중환자", "nicu"], "선천이상": ["선천"],
            "유산": ["유산"], "사산": ["사산"], "산모 입원·수술": ["산모", "임신", "출산"],
            "기타 출산 관련": ["태아", "출산", "신생아"],
        }
        allowed = [term for choice in birth_claims for term in keyword_map.get(choice, [])]
        if allowed:
            normalized = result["관련 담보"].map(normalize_text)
            birth_rows = normalized.str.contains("저체중|신생아|태아|선천|출산|산모|유산|사산", regex=True)
            keep = normalized.map(lambda value: any(normalize_text(term) in value for term in allowed))
            downgrade(birth_rows & ~keep)

    order = pd.Categorical(result["분류"], ["직접 관련", "조건부 관련", "검색 추가", "직접 추가", "함께 확인"], ordered=True)
    return result.assign(_order=order).sort_values(["보험회사", "_order", "상품명"]).drop(columns="_order").reset_index(drop=True)


def merged_documents(selected_claims: Iterable[str]) -> list[DocumentRule]:
    merged: dict[tuple[str, str], DocumentRule] = {}
    for doc in COMMON_DOCUMENTS:
        merged[(doc.name, doc.group)] = doc
    for claim in selected_claims:
        for doc in DOC_RULES.get(claim, []):
            key = (doc.name, doc.group)
            if key not in merged:
                merged[key] = doc
            else:
                current = merged[key]
                info_parts = [p.strip() for p in (current.required_info + ", " + doc.required_info).split(",") if p.strip()]
                merged[key] = DocumentRule(
                    current.name,
                    ", ".join(dict.fromkeys(info_parts)),
                    current.group,
                    current.level,
                    current.default_selected or doc.default_selected,
                )
    group_order = {"병원 발급": 0, "직접 준비": 1, "상황별 추가": 2}
    return sorted(merged.values(), key=lambda d: (group_order.get(d.group, 9), d.name))


def render_conditional_questions(selected_claims: list[str]) -> tuple[dict, list[str]]:
    """복잡한 청구의 짧은 추가 질문을 표시하고 담보 검색용 세부 항목을 반환합니다."""
    answers: dict[str, object] = {}
    derived: list[str] = []

    if "암" in selected_claims:
        with st.container(border=True):
            st.markdown("#### 암 청구 추가 확인")
            st.session_state.setdefault("cg_cancer_claims", ["암 진단비"])
            answers["cancer_claims"] = st.multiselect(
                "청구할 보험금",
                ["암 진단비", "암 수술비", "항암약물치료", "표적항암치료", "방사선치료", "양성자·세기조절치료", "중입자치료", "CAR-T"],
                key="cg_cancer_claims",
            )
            c1, c2 = st.columns(2)
            answers["biopsy"] = c1.selectbox("조직검사를 받았나요?", ["선택 전", "받음", "받지 못함", "잘 모르겠음"], key="cg_biopsy")
            answers["blood_cancer"] = c2.selectbox("혈액암 또는 골수검사로 진단받았나요?", ["선택 전", "아니요", "예", "잘 모르겠음"], key="cg_blood_cancer")

        cancer_map = {
            "암 수술비": "수술", "항암약물치료": "항암약물", "표적항암치료": "표적항암",
            "방사선치료": "항암방사선", "양성자·세기조절치료": "양성자·세기조절",
            "중입자치료": "중입자치료", "CAR-T": "CAR-T",
        }
        derived.extend(cancer_map[x] for x in answers["cancer_claims"] if x in cancer_map)

    if "사망" in selected_claims:
        with st.container(border=True):
            st.markdown("#### 사망 청구 추가 확인")
            c1, c2 = st.columns(2)
            answers["death_cause"] = c1.selectbox("사망 원인", ["선택 전", "질병", "재해·상해", "교통사고", "확인 중"], key="cg_death_cause")
            answers["beneficiary"] = c2.selectbox("보험수익자", ["선택 전", "지정수익자", "법정상속인", "잘 모르겠음"], key="cg_beneficiary")
            if answers["beneficiary"] == "법정상속인":
                c3, c4 = st.columns(2)
                answers["representative_heir"] = c3.selectbox("상속인 한 명이 대표로 청구하나요?", ["선택 전", "아니요", "예", "잘 모르겠음"], key="cg_representative_heir")
                answers["minor_beneficiary"] = c4.selectbox("미성년 수익자가 포함되어 있나요?", ["선택 전", "아니요", "예"], key="cg_minor_beneficiary")
            else:
                answers["representative_heir"] = "해당 없음"
                answers["minor_beneficiary"] = st.selectbox("미성년 수익자가 포함되어 있나요?", ["선택 전", "아니요", "예"], key="cg_minor_beneficiary")

    if "태아·출산" in selected_claims:
        with st.container(border=True):
            st.markdown("#### 태아·출산 청구 추가 확인")
            st.session_state.setdefault("cg_birth_claims", [])
            answers["birth_claims"] = st.multiselect(
                "청구 내용",
                ["저체중아", "신생아 질환·입원", "신생아중환자실", "선천이상", "유산", "사산", "산모 입원·수술", "기타 출산 관련"],
                key="cg_birth_claims",
            )
            if "유산" in answers["birth_claims"]:
                answers["miscarriage_procedure"] = st.selectbox("유산과 관련해 수술 또는 처치를 받았나요?", ["선택 전", "아니요", "예", "잘 모르겠음"], key="cg_miscarriage_procedure")
            if "산모 입원·수술" in answers["birth_claims"]:
                derived.extend(["실손 입원", "입원일당", "수술"])

    return answers, list(dict.fromkeys(derived))


def conditional_documents(base_docs: list[DocumentRule], answers: dict) -> list[DocumentRule]:
    """추가 질문 답변에 맞는 추천서류를 더하고 불필요한 조건서류는 기본 해제한다."""
    docs = list(base_docs)

    def add(doc: DocumentRule) -> None:
        if not any(d.name == doc.name and d.group == doc.group for d in docs):
            docs.append(doc)

    if answers.get("blood_cancer") == "예":
        add(DocumentRule("혈액검사 결과지", "혈액암 확정진단 근거"))
        add(DocumentRule("골수검사 결과지", "골수검사 소견과 확정진단"))
    if answers.get("biopsy") == "받지 못함":
        add(DocumentRule("대체 진단자료", "영상·혈액검사 등 암 확정진단 근거"))

    if answers.get("beneficiary") == "법정상속인":
        add(DocumentRule("사망자 가족관계증명서", "사망자 기준, 상속관계 확인", "직접 준비"))
    if answers.get("representative_heir") == "예":
        add(DocumentRule("상속인 위임장", "위임하는 상속인의 서명 또는 인감", "직접 준비"))
        add(DocumentRule("인감증명서 또는 본인서명사실확인서", "위임하는 상속인 기준", "직접 준비"))
    if answers.get("minor_beneficiary") == "예":
        add(DocumentRule("미성년자 기본증명서", "친권자 확인", "직접 준비"))
        add(DocumentRule("미성년자 가족관계증명서", "미성년자와 친권자 관계", "직접 준비"))
        add(DocumentRule("친권자 신분증 사본", "친권자 기준", "직접 준비"))

    birth_claims = set(answers.get("birth_claims", []))
    if "유산" in birth_claims:
        add(DocumentRule("유산진단서", "유산 진단명, 진단일"))
    if "사산" in birth_claims:
        add(DocumentRule("사산증명서", "사산 사실과 일자"))
    if "신생아중환자실" in birth_claims:
        add(DocumentRule("신생아중환자실 사용확인서", "입원 시작일·종료일과 이용기간"))

    group_order = {"병원 발급": 0, "직접 준비": 1, "상황별 추가": 2}
    return sorted(docs, key=lambda d: (group_order.get(d.group, 9), d.name))


def compact_required_info(value: str) -> str:
    replacements = {
        "질병분류코드": "질병코드", "급여·비급여 치료내역": "급여·비급여 내역",
        "청구인 또는 보험수익자 계좌": "청구인·수익자 계좌", "보험회사별 양식, ": "",
        "최종 병리진단과 조직검사 결과": "최종 병리진단",
    }
    result = value
    for old, new in replacements.items():
        result = result.replace(old, new)
    return result


def make_customer_message(docs: list[DocumentRule], customer_name: str = "") -> str:
    clean_name = re.sub(r"\s*고객님\s*$", "", customer_name.strip())
    greeting = f"{clean_name} 고객님" if clean_name else "고객님"
    lines = [f"{greeting}, 보험금 청구에 필요한 서류를 안내드립니다."]
    labels = {"병원 발급": "병원 발급서류", "직접 준비": "직접 준비서류", "상황별 추가": "추가 준비서류"}
    for group in ["병원 발급", "직접 준비", "상황별 추가"]:
        group_docs = [d for d in docs if d.group == group]
        if not group_docs:
            continue
        lines.extend(["", f"[{labels[group]}]"])
        for i, doc in enumerate(group_docs, start=1):
            info = compact_required_info(doc.required_info)
            lines.append(f"{i}. {doc.name}" + (f"({info})" if info else ""))
    lines.extend(["", "보험회사 심사 과정에서 추가서류가 요청될 수 있습니다."])
    return "\n".join(lines)


def build_accident_narrative(accident_date: str, place: str, course: str, body_part: str, visit_date: str = "", treatment: str = "") -> str:
    if not all([accident_date.strip(), place.strip(), course.strip(), body_part.strip()]):
        return ""
    text = f"{accident_date.strip()} {place.strip()}에서 {course.strip()} 사고로 {body_part.strip()} 부위를 다쳤습니다."
    if visit_date.strip():
        text += f" 이후 {visit_date.strip()} 병원에 내원하였습니다."
        if treatment.strip():
            text += f" 진단·치료 내용은 다음과 같습니다: {treatment.strip()}."
    elif treatment.strip():
        text += f" 이후 병원에서 진단 또는 치료를 받았으며, 내용은 다음과 같습니다: {treatment.strip()}."
    return text


def _register_korean_font() -> str:
    """프로젝트의 Pretendard TTF를 등록하고, 없으면 기본 한글 글꼴을 사용합니다."""
    font_name = "Pretendard"
    module_dir = Path(__file__).resolve().parent
    font_candidates = [
        module_dir.parent / "assets" / "fonts" / "PretendardVariable.ttf",
        module_dir / "assets" / "fonts" / "PretendardVariable.ttf",
        Path.cwd() / "assets" / "fonts" / "PretendardVariable.ttf",
    ]

    font_path = next((path for path in font_candidates if path.is_file()), None)
    if font_path is not None:
        if font_name not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont(font_name, str(font_path)))
        return font_name

    fallback_name = "HYGoThic-Medium"
    if fallback_name not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(UnicodeCIDFont(fallback_name))
    return fallback_name


def build_guide_pdf(selected_claims: list[str], docs: list[DocumentRule], accident_narrative: str = "", include_accident: bool = False) -> bytes:
    font_name = _register_korean_font()
    buffer = BytesIO()

    class NumberedCanvasMixin:
        pass

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(font_name, 7.5)
        canvas.setFillColor(colors.HexColor("#64748B"))
        canvas.drawString(15 * mm, 9 * mm, f"보험금 청구 가이드 · {date.today():%Y.%m.%d}")
        canvas.drawRightString(A4[0] - 15 * mm, 9 * mm, f"{doc.page}페이지")
        canvas.restoreState()

    doc = BaseDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=13 * mm,
        bottomMargin=15 * mm,
        title="보험금 청구 준비 안내",
        author="보험금 청구 가이드",
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal")
    doc.addPageTemplates(PageTemplate(id="guide", frames=[frame], onPage=footer))

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("KTitle", parent=styles["Title"], fontName=font_name, fontSize=18, leading=23, alignment=TA_CENTER, textColor=colors.HexColor("#183B67"), spaceAfter=7 * mm)
    subtitle_style = ParagraphStyle("KSub", parent=styles["Normal"], fontName=font_name, fontSize=9.2, leading=13, alignment=TA_CENTER, textColor=colors.HexColor("#475569"), spaceAfter=5 * mm)
    section_style = ParagraphStyle("KSection", parent=styles["Heading2"], fontName=font_name, fontSize=11.5, leading=15, textColor=colors.HexColor("#1D4E89"), spaceBefore=3 * mm, spaceAfter=2 * mm)
    body_style = ParagraphStyle("KBody", parent=styles["BodyText"], fontName=font_name, fontSize=8.6, leading=12, textColor=colors.HexColor("#26384A"))
    small_style = ParagraphStyle("KSmall", parent=body_style, fontSize=7.8, leading=10.5, textColor=colors.HexColor("#52677A"))

    story = [
        Paragraph("보험금 청구 준비 안내", title_style),
        Paragraph(f"청구 항목: {' · '.join(html.escape(x) for x in selected_claims)}<br/>안내서 작성일: {date.today():%Y년 %m월 %d일}", subtitle_style),
    ]

    def add_doc_table(group: str, heading: str):
        rows = [d for d in docs if d.group == group]
        if not rows:
            return
        story.append(Paragraph(heading, section_style))
        data = [[Paragraph("준비", small_style), Paragraph("서류", small_style), Paragraph("반드시 포함될 정보", small_style)]]
        for item in rows:
            level = "" if item.level == "기본 준비" else f"<br/><font color='#64748B'>({html.escape(item.level)})</font>"
            data.append([
                Paragraph("□", body_style),
                Paragraph(html.escape(item.name) + level, body_style),
                Paragraph(html.escape(item.required_info), body_style),
            ])
        table = Table(data, colWidths=[12 * mm, 48 * mm, 120 * mm], repeatRows=1, hAlign="LEFT")
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAF2FB")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1D4E89")),
            ("FONTNAME", (0, 0), (-1, -1), font_name),
            ("ALIGN", (0, 0), (0, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#C9D6E4")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.extend([table, Spacer(1, 2 * mm)])

    add_doc_table("병원 발급", "1. 병원에서 발급받을 서류")
    add_doc_table("직접 준비", "2. 직접 준비할 서류")
    add_doc_table("상황별 추가", "3. 상황별 추가서류")

    injury_selected = any(x in selected_claims for x in ["일반 상해", "교통사고", "골절", "화상", "후유장해"])
    if injury_selected:
        story.append(Paragraph("상해사고 필수 정보", section_style))
        story.append(Paragraph("□ 사고일자　□ 사고 장소　□ 사고 당시 상황　□ 구체적인 사고경위　□ 다친 신체 부위　□ 최초 진료일", body_style))

    if include_accident and accident_narrative.strip():
        story.append(Paragraph("사고경위", section_style))
        accident_table = Table([[Paragraph(html.escape(accident_narrative.strip()), body_style)]], colWidths=[180 * mm])
        accident_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F4F8FD")),
            ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#B8CCE4")),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ]))
        story.extend([accident_table, Spacer(1, 1.5 * mm), Paragraph("실제 사고내용과 일치하는지 확인한 후 보험금 청구서에 사용해 주세요.", small_style)])

    story.extend([
        Spacer(1, 3 * mm),
        Paragraph("보험회사와 가입 담보에 따라 필요서류가 달라지거나 추가될 수 있습니다. 발급비용이 큰 진단서와 후유장해진단서는 발급 전에 해당 보험회사에 확인해 주세요.", small_style),
        Spacer(1, 1.5 * mm),
        Paragraph("본 안내서는 선택한 청구 항목을 기준으로 작성된 서류 준비 가이드입니다. 실제 보험금 지급 여부는 가입 약관과 보험회사의 심사 결과에 따라 결정됩니다.", small_style),
    ])
    doc.build(story)
    return buffer.getvalue()


def clear_state() -> None:
    for key in list(st.session_state.keys()):
        if key.startswith(STATE_PREFIX):
            del st.session_state[key]


def toggle_claim(claim: str) -> None:
    selected = set(st.session_state.get("cg_selected_claims", []))
    if claim in selected:
        selected.remove(claim)
    else:
        selected.add(claim)
    ordered = [item for group in CLAIM_GROUPS.values() for item in group if item in selected]
    st.session_state["cg_selected_claims"] = ordered
    st.session_state.pop("cg_result_claims", None)


def inject_styles() -> None:
    st.markdown(
        """
        <style>
        .cg-selected-summary {padding:.72rem .88rem;border:1px solid #CBDDF0;border-radius:.72rem;background:#F6FAFF;margin:.4rem 0 .8rem;}
        .cg-selected-summary b {color:#1D4E89;}
        .cg-status-direct,.cg-status-condition,.cg-status-related {display:inline-flex;padding:.16rem .45rem;border-radius:999px;font-size:.75rem;font-weight:700;}
        .cg-hospital-view {padding:1.4rem;border:2px solid #B7CCE3;border-radius:1rem;background:#FFF;color:#17334B;}
        .cg-hospital-view h3 {font-size:1.35rem!important;color:#183B67!important;margin-top:0!important;}
        .cg-hospital-view li {font-size:1.04rem;line-height:1.72;margin-bottom:.45rem;}
        @media(max-width:700px){.cg-hospital-view{padding:1rem}.cg-hospital-view li{font-size:.96rem}}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_claim_buttons() -> list[str]:
    selected = st.session_state.setdefault("cg_selected_claims", [])
    for group, claims in CLAIM_GROUPS.items():
        st.markdown(f"#### {group}")
        columns = st.columns(4)
        for index, claim in enumerate(claims):
            with columns[index % 4]:
                if st.button(
                    claim,
                    key=f"cg_claim_{claim}",
                    type="primary" if claim in selected else "secondary",
                    use_container_width=True,
                ):
                    toggle_claim(claim)
                    st.rerun()
    selected = st.session_state.get("cg_selected_claims", [])
    if selected:
        chips = " · ".join(html.escape(x) for x in selected)
        st.markdown(f'<div class="cg-selected-summary"><b>선택한 청구 항목 {len(selected)}개</b><br>{chips}</div>', unsafe_allow_html=True)
        if st.button("전체 선택 해제", key="cg_clear_claims"):
            st.session_state["cg_selected_claims"] = []
            st.session_state.pop("cg_result_claims", None)
            st.rerun()
    return selected


def _coverage_key(row: dict) -> str:
    return "|".join(str(row.get(x, "")) for x in ["보험회사", "상품명", "관련 담보", "계약일"])


def _coverage_editor_table(df: pd.DataFrame, key: str) -> pd.DataFrame:
    return st.data_editor(
        df, key=key, hide_index=True, use_container_width=True,
        disabled=["보험회사", "상품명", "분류", "추출상태", "계약일", "만기일", "원본쪽", "매칭근거"],
        column_config={
            "포함": st.column_config.CheckboxColumn("포함"),
            "관련 담보": st.column_config.TextColumn("관련 담보", width="large"),
            "가입금액": st.column_config.TextColumn("가입금액", width="small"),
            "확인사항": st.column_config.TextColumn("확인사항", width="large"),
            "원본쪽": None, "매칭근거": None,
        },
    )


def render_coverage_editor(matched_df: pd.DataFrame, all_coverages: list[dict]) -> pd.DataFrame:
    manual_df = pd.DataFrame(st.session_state.get("cg_manual_coverages", []))
    searched_df = pd.DataFrame(st.session_state.get("cg_searched_coverages", []))
    frames = [matched_df]
    if not searched_df.empty:
        frames.append(searched_df)
    if not manual_df.empty:
        frames.append(manual_df)
    display_df = pd.concat(frames, ignore_index=True).drop_duplicates(
        subset=["보험회사", "상품명", "관련 담보", "계약일"], keep="last"
    )

    direct_df = display_df[display_df["분류"] != "함께 확인"].copy()
    related_df = display_df[display_df["분류"] == "함께 확인"].copy()
    edited_frames: list[pd.DataFrame] = []

    if direct_df.empty:
        st.info("선택한 청구와 직접 일치하는 담보를 찾지 못했습니다. 아래에서 담보명을 검색하거나 직접 추가해 주세요.")
    else:
        c1, c2 = st.columns(2)
        c1.metric("보험회사", f"{direct_df['보험회사'].nunique()}개")
        c2.metric("직접 관련 담보", f"{len(direct_df)}개")
        edited_frames.append(_coverage_editor_table(direct_df, "cg_coverage_direct_editor"))

    if not related_df.empty:
        related_df["포함"] = related_df["포함"].fillna(False)
        with st.expander(f"함께 확인할 담보 {len(related_df)}개 보기"):
            edited_frames.append(_coverage_editor_table(related_df, "cg_coverage_related_editor"))

    search_expanded = direct_df.empty
    with st.expander("다른 담보 검색해서 추가", expanded=search_expanded):
        search_term = st.text_input("담보명 검색", key="cg_coverage_search", placeholder="예: 표적항암, 뇌혈관, 질병수술")
        if search_term.strip():
            needle = normalize_text(search_term)
            candidates = []
            existing_keys = {_coverage_key(r) for r in display_df.to_dict("records")}
            for row in all_coverages:
                searchable = normalize_text(f"{row.get('company', '')} {row.get('product', '')} {row.get('coverage', '')}")
                if needle and needle in searchable:
                    candidate = {
                        "포함": True, "보험회사": row.get("company", "확인 필요"), "상품명": row.get("product", "확인 필요"),
                        "관련 담보": row.get("coverage", "확인 필요"), "가입금액": row.get("amount", "확인 필요"), "분류": "검색 추가",
                        "확인사항": "담보 검색으로 추가", "추출상태": row.get("extraction_status", "정상 추출"),
                        "계약일": row.get("contract_date", ""), "만기일": row.get("expiry_date", ""),
                        "원본쪽": row.get("source_page", ""), "매칭근거": search_term.strip(),
                    }
                    if _coverage_key(candidate) not in existing_keys:
                        candidates.append(candidate)
            if not candidates:
                st.caption("추가할 수 있는 새로운 담보를 찾지 못했습니다.")
            else:
                candidate_df = pd.DataFrame(candidates)
                for company, company_df in candidate_df.groupby("보험회사", sort=True):
                    st.markdown(f"**{company} · {len(company_df)}개**")
                    for idx, row in company_df.iterrows():
                        cols = st.columns([4, 1])
                        cols[0].markdown(f"{row['관련 담보']}　·　{row['가입금액']}  \n<small>{row['상품명']}</small>", unsafe_allow_html=True)
                        if cols[1].button("추가", key=f"cg_add_search_{hashlib.md5(_coverage_key(row.to_dict()).encode()).hexdigest()}"):
                            st.session_state.setdefault("cg_searched_coverages", []).append(row.to_dict())
                            st.rerun()

        searched = st.session_state.get("cg_searched_coverages", [])
        if searched:
            st.markdown("##### 검색으로 추가한 담보")
            for row in searched:
                cols = st.columns([5, 1])
                cols[0].write(f"{row['보험회사']} · {row['관련 담보']} · {row['가입금액']}")
                if cols[1].button("제외", key=f"cg_remove_search_{hashlib.md5(_coverage_key(row).encode()).hexdigest()}"):
                    st.session_state["cg_searched_coverages"] = [x for x in searched if _coverage_key(x) != _coverage_key(row)]
                    st.rerun()

    with st.expander("보장분석에 없는 담보 직접 추가", expanded=not all_coverages):
        with st.form("cg_manual_coverage_form", clear_on_submit=True):
            cols = st.columns([1, 1.4, 1.5, .7])
            company = cols[0].text_input("보험회사")
            product = cols[1].text_input("상품명")
            coverage = cols[2].text_input("담보명")
            amount = cols[3].text_input("가입금액")
            note = st.text_input("확인사항", value="사용자 직접 추가")
            submitted = st.form_submit_button("담보 추가", type="primary")
        if submitted:
            manual = st.session_state.setdefault("cg_manual_coverages", [])
            manual.append({
                "포함": True, "보험회사": company or "직접 입력", "상품명": product or "확인 필요",
                "관련 담보": coverage or "확인 필요", "가입금액": amount or "확인 필요", "분류": "직접 추가",
                "확인사항": note, "추출상태": "사용자 입력", "계약일": "", "만기일": "", "원본쪽": "", "매칭근거": "직접 추가",
            })
            st.rerun()
        manual = st.session_state.get("cg_manual_coverages", [])
        for row in manual:
            cols = st.columns([5, 1])
            cols[0].write(f"{row['보험회사']} · {row['관련 담보']} · {row['가입금액']}")
            if cols[1].button("삭제", key=f"cg_remove_manual_cov_{hashlib.md5(_coverage_key(row).encode()).hexdigest()}"):
                st.session_state["cg_manual_coverages"] = [x for x in manual if _coverage_key(x) != _coverage_key(row)]
                st.rerun()

    return pd.concat(edited_frames, ignore_index=True) if edited_frames else pd.DataFrame(columns=display_df.columns)


def render_document_editor(docs: list[DocumentRule], recommendation_token: str) -> list[DocumentRule]:
    manual_docs = [DocumentRule(**x) for x in st.session_state.get("cg_manual_documents", [])]
    all_docs = docs + manual_docs
    recommended_keys = [f"{d.group}|{d.name}" for d in docs if d.default_selected]
    manual_keys = [f"{d.group}|{d.name}" for d in manual_docs]

    if st.session_state.get("cg_doc_recommendation_token") != recommendation_token:
        st.session_state["cg_selected_docs"] = list(dict.fromkeys(recommended_keys + manual_keys))
        st.session_state["cg_doc_recommendation_token"] = recommendation_token
        for key in list(st.session_state):
            if key.startswith("cg_doc_") and key != "cg_doc_recommendation_token":
                st.session_state.pop(key, None)

    title_cols = st.columns([4, 1.4])
    title_cols[0].caption("필요하지 않은 서류는 자유롭게 해제할 수 있습니다.")
    if title_cols[1].button("추천 서류로 되돌리기", key="cg_restore_docs", use_container_width=True):
        st.session_state["cg_selected_docs"] = list(dict.fromkeys(recommended_keys + manual_keys))
        for key in list(st.session_state):
            if key.startswith("cg_doc_") and key != "cg_doc_recommendation_token":
                st.session_state.pop(key, None)
        st.success("현재 답변을 기준으로 추천 서류를 복원했습니다.")
        st.rerun()

    selected_keys = st.session_state.setdefault("cg_selected_docs", recommended_keys + manual_keys)
    current_keys = {f"{d.group}|{d.name}" for d in all_docs}
    selected_keys[:] = [key for key in selected_keys if key in current_keys]
    result: list[DocumentRule] = []
    for group in ["병원 발급", "직접 준비", "상황별 추가"]:
        group_docs = [d for d in all_docs if d.group == group]
        if not group_docs:
            continue
        st.markdown(f"#### {group}")
        for doc in group_docs:
            key = f"{doc.group}|{doc.name}"
            checked = st.checkbox(
                f"{doc.name} · {doc.required_info}",
                value=key in selected_keys,
                key=f"cg_doc_{hashlib.md5(key.encode()).hexdigest()}",
                help=doc.level,
            )
            if checked:
                result.append(doc)
                if key not in selected_keys:
                    selected_keys.append(key)
            elif key in selected_keys:
                selected_keys.remove(key)

    with st.expander("필요서류 직접 추가"):
        with st.form("cg_manual_document_form", clear_on_submit=True):
            c1, c2 = st.columns([1, 1.6])
            doc_name = c1.text_input("서류명")
            required_info = c2.text_input("필수 기재사항")
            doc_group = st.selectbox("구분", ["병원 발급", "직접 준비", "상황별 추가"])
            submitted = st.form_submit_button("서류 추가", type="primary")
        if submitted and doc_name.strip():
            item = asdict(DocumentRule(doc_name.strip(), required_info.strip(), doc_group, "사용자 직접 추가"))
            st.session_state.setdefault("cg_manual_documents", []).append(item)
            st.session_state.setdefault("cg_selected_docs", []).append(f"{doc_group}|{doc_name.strip()}")
            st.rerun()

        manual_items = st.session_state.get("cg_manual_documents", [])
        for item in manual_items:
            item_key = f"{item['group']}|{item['name']}"
            cols = st.columns([5, 1])
            cols[0].write(f"{item['name']} · {item['required_info'] or '기재사항 없음'}")
            if cols[1].button("삭제", key=f"cg_delete_doc_{hashlib.md5(item_key.encode()).hexdigest()}"):
                st.session_state["cg_manual_documents"] = [x for x in manual_items if f"{x['group']}|{x['name']}" != item_key]
                st.session_state["cg_selected_docs"] = [x for x in selected_keys if x != item_key]
                st.rerun()
    return result


def render_accident_helper(selected_claims: list[str]) -> str:
    injury_claims = {"일반 상해", "교통사고", "골절", "화상", "후유장해"}
    if not injury_claims.intersection(selected_claims):
        return ""
    with st.expander("사고경위 작성 도우미 · 필요할 때만 사용"):
        c1, c2 = st.columns(2)
        accident_date = c1.text_input("사고일자", key="cg_accident_date", placeholder="예: 2026년 8월 13일")
        place = c2.text_input("사고 장소", key="cg_accident_place", placeholder="예: 자택 화장실")
        course = st.text_area("사고 당시 상황과 발생 과정", key="cg_accident_course", placeholder="예: 화장실에서 나오다가 문턱에 발가락을 부딪힘")
        c3, c4 = st.columns(2)
        body_part = c3.text_input("다친 부위", key="cg_accident_body", placeholder="예: 오른쪽 네 번째 발가락")
        visit_date = c4.text_input("최초 병원 방문일 · 선택", key="cg_visit_date")
        treatment = st.text_input("진단 또는 치료 내용 · 선택", key="cg_treatment")
        if st.button("사고경위 생성", key="cg_generate_accident", type="primary"):
            generated = build_accident_narrative(accident_date, place, course, body_part, visit_date, treatment)
            if generated:
                st.session_state["cg_accident_final"] = generated
            else:
                st.warning("사고일자, 장소, 사고 과정과 다친 부위를 입력해 주세요.")
        narrative = st.text_area("최종 사고경위", key="cg_accident_final", height=120)
        st.session_state["cg_accident_narrative"] = narrative
        if narrative.strip():
            st.checkbox("A4 청구 준비 안내서에 사고경위 포함", key="cg_include_accident_pdf", value=False)
            st.caption("입력한 사실과 일치하는지 확인한 후 사용해 주세요. 문자 안내문에는 자동으로 포함되지 않습니다.")
        return narrative


def run() -> None:
    inject_styles()
    page_header("고객 상담", "보험금 청구 가이드", "청구 항목별 필요서류를 확인하고 보장분석 PDF에서 관련 담보와 가입금액을 찾습니다.", "CG")

    top_left, top_right = st.columns([5, 1])
    with top_left:
        st.caption(f"청구서류 가이드 기준일 · {GUIDE_STANDARD_DATE}　|　버전 {GUIDE_VERSION}")
    with top_right:
        if st.button("처음부터 다시", key="cg_reset", use_container_width=True):
            clear_state()
            st.rerun()

    pending_customer = st.session_state.pop("cg_pending_customer_name", "")
    existing_parsed = st.session_state.get("cg_parsed_pdf") or {}
    existing_pdf_customer = str(existing_parsed.get("customer", "")).strip()
    if pending_customer:
        st.session_state["cg_customer_name"] = pending_customer
    elif not st.session_state.get("cg_customer_name") and existing_pdf_customer not in {"", "확인 필요"}:
        st.session_state["cg_customer_name"] = existing_pdf_customer
    customer_name = st.text_input("고객 이름 · 선택사항", key="cg_customer_name", placeholder="이름을 입력하지 않아도 안내문을 만들 수 있습니다.")

    section_intro("청구 유형", "무엇을 청구하시나요?", "해당하는 항목을 여러 개 선택할 수 있습니다.")
    selected_claims = render_claim_buttons()
    if not selected_claims:
        st.info("청구 항목을 선택하면 추천서류와 문자 안내문을 바로 확인할 수 있습니다.")
        return

    answers, derived_claims = render_conditional_questions(selected_claims)
    effective_claims = list(dict.fromkeys(selected_claims + derived_claims))

    section_intro("선택사항", "보장분석 PDF로 관련 담보도 확인하기")
    with st.expander("보장분석 PDF 첨부"):
        st.caption("첨부하지 않아도 서류 안내와 문자 안내문을 이용할 수 있습니다.")
        uploaded = st.file_uploader("프로보장분석 PDF 선택", type=["pdf"], key="cg_uploader")
    parsed = st.session_state.get("cg_parsed_pdf")
    if uploaded:
        pdf_bytes = uploaded.getvalue()
        file_hash = hashlib.sha256(pdf_bytes).hexdigest()
        if st.session_state.get("cg_pdf_hash") != file_hash:
            try:
                with st.spinner("보장분석 PDF에서 가입내용을 확인하고 있습니다..."):
                    parsed = extract_pdf(pdf_bytes)
                st.session_state["cg_pdf_hash"] = file_hash
                st.session_state["cg_parsed_pdf"] = parsed
                st.session_state.pop("cg_coverage_direct_editor", None)
                st.session_state.pop("cg_coverage_related_editor", None)
                st.session_state["cg_parsed_just_now"] = True
            except Exception as exc:
                st.session_state.pop("cg_parsed_pdf", None)
                parsed = None
                st.warning(f"지원되는 형식으로 가입내용을 확인하지 못했습니다. 서류 가이드는 계속 이용할 수 있습니다. ({exc})")
    elif st.session_state.get("cg_pdf_hash"):
        st.session_state.pop("cg_pdf_hash", None)
        st.session_state.pop("cg_parsed_pdf", None)
        parsed = None

    if st.session_state.pop("cg_parsed_just_now", False):
        st.rerun()

    if parsed:
        coverages = parsed.get("coverages", [])
        pdf_customer = str(parsed.get("customer", "")).strip()
        if pdf_customer and pdf_customer != "확인 필요":
            typed_customer = str(st.session_state.get("cg_customer_name", "")).strip()
            if typed_customer and normalize_text(typed_customer) != normalize_text(pdf_customer):
                st.warning(f"입력한 고객 이름({typed_customer})과 보장분석 PDF의 고객 이름({pdf_customer})이 다릅니다.")
                if st.button("PDF 이름 적용", key="cg_apply_pdf_customer"):
                    st.session_state["cg_pending_customer_name"] = pdf_customer
                    st.rerun()
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("고객", parsed.get("customer", "확인 필요"))
        c2.metric("작성 기준일", parsed.get("report_date", "확인 필요"))
        c3.metric("보험회사", f"{len({x['company'] for x in coverages})}개")
        c4.metric("추출 담보", f"{len(coverages)}개")
        try:
            report_dt = datetime.strptime(parsed.get("report_date", ""), "%Y.%m.%d").date()
            if (date.today() - report_dt).days >= 180:
                st.warning("보장분석 작성일로부터 6개월 이상 지났습니다. 사고일 당시 계약 유지 여부를 확인해 주세요.")
        except ValueError:
            pass
        with st.expander("추출된 가입내용 확인"):
            if coverages:
                st.dataframe(pd.DataFrame(coverages), hide_index=True, use_container_width=True)
            else:
                st.info("담보 상세표를 추출하지 못했습니다. 관련 담보는 직접 추가할 수 있습니다.")

    if parsed:
        section_intro("FP 확인", "회사별 관련 담보", "직접 관련 담보만 우선 표시하며, 필요한 담보를 검색해 추가할 수 있습니다.")
        matched_df = refine_matches_with_answers(match_coverages(parsed.get("coverages", []), effective_claims), answers)
        included_coverages = render_coverage_editor(matched_df, parsed.get("coverages", []))
        st.session_state["cg_included_coverages"] = included_coverages.to_dict("records") if not included_coverages.empty else []
        st.caption("가입금액은 보장분석 PDF 표시 기준이며 실제 지급액을 의미하지 않습니다.")

    section_intro("준비서류", "필요서류와 필수 기재사항", "선택 결과는 문자 안내문과 안내서 PDF에 함께 반영됩니다.")
    docs = conditional_documents(merged_documents(effective_claims), answers)
    recommendation_token = hashlib.sha256(repr((effective_claims, answers, [(d.group, d.name, d.default_selected) for d in docs])).encode()).hexdigest()
    selected_docs = render_document_editor(docs, recommendation_token)
    accident_narrative = render_accident_helper(effective_claims)

    section_intro("청구 안내자료", "고객·병원 전달")
    tab_message, tab_hospital, tab_pdf = st.tabs(["문자 안내문", "병원 발급 요청", "안내서 PDF"])
    with tab_message:
        default_message = make_customer_message(selected_docs, st.session_state.get("cg_customer_name", ""))
        if st.session_state.pop("cg_pending_message_restore", False):
            st.session_state["cg_customer_message"] = default_message
            st.session_state["cg_message_source"] = default_message
        elif "cg_customer_message" not in st.session_state:
            st.session_state["cg_customer_message"] = default_message
            st.session_state["cg_message_source"] = default_message
        elif st.session_state.get("cg_message_source") != default_message:
            st.info("고객 이름 또는 선택한 서류가 변경되었습니다. 변경 내용을 반영하려면 기본 안내문으로 복원해 주세요.")
        st.text_area("문자 안내문", key="cg_customer_message", height=280)
        st.caption("필요한 문장을 자유롭게 추가하거나 수정할 수 있습니다. 회사명·담보명·가입금액은 포함되지 않습니다.")
        if st.button("기본 안내문으로 복원", key="cg_restore_message"):
            st.session_state["cg_pending_message_restore"] = True
            st.rerun()

    with tab_hospital:
        hospital_docs = [d for d in selected_docs if d.group == "병원 발급"]
        if hospital_docs:
            items = "".join(f"<li><b>{html.escape(d.name)}</b><br>{html.escape(d.required_info)}</li>" for d in hospital_docs)
            st.markdown(f'<div class="cg-hospital-view"><h3>보험금 청구용 서류 발급 요청</h3><ol>{items}</ol></div>', unsafe_allow_html=True)
        else:
            st.info("선택된 병원 발급서류가 없습니다.")

    with tab_pdf:
        include_accident = bool(st.session_state.get("cg_include_accident_pdf", False))
        try:
            pdf_bytes = build_guide_pdf(effective_claims, selected_docs, accident_narrative, include_accident)
            st.download_button(
                "보험금 청구 준비 안내서 PDF 다운로드",
                data=pdf_bytes,
                file_name=f"보험금_청구_준비_안내서_{date.today():%Y%m%d}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True,
            )
            if include_accident and accident_narrative.strip():
                st.success("최종 수정된 사고경위가 A4 안내서에 포함됩니다.")
            else:
                st.caption("사고경위는 기본적으로 포함되지 않습니다. 사고경위 도우미에서 선택하면 포함할 수 있습니다.")
        except Exception as exc:
            st.error(f"PDF 안내서를 생성하지 못했습니다: {exc}")

    st.divider()
    st.caption("이 가이드는 보장분석 자료와 선택한 청구 항목을 기준으로 관련 담보와 준비서류를 안내합니다. 실제 지급 여부와 추가서류는 가입 약관 및 보험회사의 심사 결과에 따라 달라질 수 있습니다.")
