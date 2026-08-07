from __future__ import annotations

# 전달용 파일: 보험회사·상품·세부 조건 단계형 선택 적용본 v4

import hashlib
import io
import re
from dataclasses import dataclass
from typing import Any

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


DEFAULT_PAYOUT_RATE = 65.0
FIRST_YEAR_HEADERS = {"1차년계", "1차년도합계", "1차년합계"}
TOTAL_HEADERS = {"총수수료", "총계", "총합계", "총수수료계"}
PRODUCT_HEADERS = {"상품명"}
PRODUCT_FALLBACK_HEADERS = {"구분"}


@dataclass(frozen=True)
class ProductRate:
    key: str
    source_type: str
    insurer: str
    product: str
    conditions: str
    first_year_rate: float
    total_rate: float
    sheet_name: str
    row_number: int

    @property
    def label(self) -> str:
        detail = f" · {self.conditions}" if self.conditions else ""
        return f"{self.product}{detail}"


def _normalize(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("計", "계")
    return re.sub(r"[\s\n\r\t:()\[\]·ㆍ_-]+", "", text).lower()


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def _number(value: Any) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if text in {"", "-"}:
            return None
        if text.endswith("%"):
            try:
                return float(text[:-1]) / 100
            except ValueError:
                return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def _effective_max_col(ws) -> int:
    """서식만 남아 XFD까지 확장된 시트의 불필요한 탐색을 막습니다."""
    upper = min(ws.max_column, 120)
    last = 1
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 40), max_col=upper):
        for cell in row:
            if cell.value not in (None, ""):
                last = max(last, cell.column)
    return min(max(last + 4, 20), upper)


def _is_header(value: Any, aliases: set[str]) -> bool:
    normalized = _normalize(value)
    return normalized in aliases


def _header_positions(ws, max_col: int) -> list[tuple[int, int, int, int]]:
    """상품명/1차년계/총수수료 열로 구성된 표 구간을 찾습니다."""
    positions: list[tuple[int, int, int, int]] = []
    for row_no in range(1, ws.max_row + 1):
        product_cols = [
            col for col in range(1, max_col + 1)
            if _is_header(ws.cell(row_no, col).value, PRODUCT_HEADERS)
        ]
        if not product_cols:
            product_cols = [
                col for col in range(1, max_col + 1)
                if _is_header(ws.cell(row_no, col).value, PRODUCT_FALLBACK_HEADERS)
            ]

        first_candidates: list[tuple[int, int]] = []
        total_candidates: list[tuple[int, int]] = []
        for header_row in range(row_no, min(row_no + 4, ws.max_row + 1)):
            for col in range(1, max_col + 1):
                value = ws.cell(header_row, col).value
                if _is_header(value, FIRST_YEAR_HEADERS):
                    first_candidates.append((header_row, col))
                if _is_header(value, TOTAL_HEADERS):
                    total_candidates.append((header_row, col))

        if not product_cols and first_candidates and total_candidates:
            product_cols = [1]
        if not product_cols:
            continue

        for product_col in product_cols:
            first_after = [item for item in first_candidates if item[1] > product_col]
            total_after = [item for item in total_candidates if item[1] > product_col]
            if not first_after or not total_after:
                continue
            first = min(first_after, key=lambda item: item[1])
            total = max(total_after, key=lambda item: item[1])
            if first[1] < total[1]:
                data_start = max(row_no, first[0], total[0]) + 1
                positions.append((data_start, product_col, first[1], total[1]))
                break
    return positions


def _condition_text(ws, row_no: int, product_col: int, first_col: int) -> str:
    ignored = {"-", "상품별상이", "해당없음"}
    values: list[str] = []
    for col in range(product_col + 1, first_col):
        value = ws.cell(row_no, col).value
        if isinstance(value, str) and value.startswith("="):
            continue
        text = _clean_text(value)
        if not text or text in ignored or text in values or _number(text) is not None:
            continue
        values.append(text)
    return " / ".join(values[:6])


def _source_payout_rate(formula_ws, value_ws, max_col: int) -> float | None:
    for row_no in range(1, min(formula_ws.max_row, 5) + 1):
        for col in range(1, max_col + 1):
            label = _normalize(formula_ws.cell(row_no, col).value)
            if label not in {"지급율", "지급률"}:
                continue
            for offset in range(1, 4):
                rate = _number(value_ws.cell(row_no, col + offset).value)
                if rate is not None:
                    return rate
    return None


def _extract_sheet(
    formula_ws,
    value_ws,
    source_type: str,
) -> tuple[list[ProductRate], list[str]]:
    insurer = formula_ws.title.strip()
    max_col = _effective_max_col(formula_ws)
    tables = _header_positions(formula_ws, max_col)
    results: list[ProductRate] = []
    warnings: list[str] = []
    source_payout = _source_payout_rate(formula_ws, value_ws, max_col)

    if not tables:
        return results, warnings
    if source_payout == 0:
        warnings.append(
            f"{insurer}: 원본 예시표의 지급율이 0%입니다. 해당 시트를 100%로 저장한 뒤 다시 올려 주세요."
        )
        return results, warnings

    for table_index, (data_start, product_col, first_col, total_col) in enumerate(tables):
        next_start = tables[table_index + 1][0] if table_index + 1 < len(tables) else formula_ws.max_row + 1
        end_row = next_start - 2
        current_product = ""

        for row_no in range(data_start, end_row + 1):
            raw_product = formula_ws.cell(row_no, product_col).value
            product_text = _clean_text(raw_product)

            if product_text:
                normalized_product = _normalize(product_text)
                if (
                    normalized_product in PRODUCT_HEADERS
                    or product_text.startswith("■")
                    or "수수료타입변경" in normalized_product
                ):
                    current_product = ""
                    continue
                current_product = product_text

            if not current_product:
                continue

            first_rate = _number(value_ws.cell(row_no, first_col).value)
            total_rate = _number(value_ws.cell(row_no, total_col).value)
            if first_rate is None or total_rate is None or first_rate < 0 or total_rate < 0:
                continue
            if first_rate == 0 and total_rate == 0:
                continue

            if source_payout not in (None, 0):
                first_rate /= source_payout
                total_rate /= source_payout

            conditions = _condition_text(formula_ws, row_no, product_col, first_col)
            identity = f"{source_type}|{insurer}|{current_product}|{conditions}|{row_no}"
            key = hashlib.sha1(identity.encode("utf-8")).hexdigest()[:16]
            results.append(
                ProductRate(
                    key=key,
                    source_type=source_type,
                    insurer=insurer,
                    product=current_product,
                    conditions=conditions,
                    first_year_rate=first_rate,
                    total_rate=total_rate,
                    sheet_name=formula_ws.title,
                    row_number=row_no,
                )
            )

    if tables and not results:
        warnings.append(f"{insurer}: 표는 찾았지만 계산된 수수료율을 읽지 못했습니다.")
    return results, warnings


@st.cache_data(show_spinner=False)
def parse_commission_workbook(file_bytes: bytes, source_type: str) -> tuple[list[dict], list[str]]:
    """예시표의 저장된 계산 결과를 읽습니다. 원본 파일은 변경하지 않습니다."""
    formula_book = load_workbook(io.BytesIO(file_bytes), data_only=False, read_only=False)
    value_book = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=False)
    products: list[dict] = []
    warnings: list[str] = []

    for sheet_name in formula_book.sheetnames:
        if "변경" in sheet_name or sheet_name not in value_book.sheetnames:
            continue
        extracted, sheet_warnings = _extract_sheet(
            formula_book[sheet_name], value_book[sheet_name], source_type
        )
        products.extend(item.__dict__ for item in extracted)
        warnings.extend(sheet_warnings)

    formula_book.close()
    value_book.close()
    return products, warnings


def _to_product_rate(item: dict) -> ProductRate:
    return ProductRate(**item)


def _format_rate(multiplier: float) -> str:
    return f"{multiplier * 100:,.1f}%"


def _format_won(value: float) -> str:
    return f"{round(value):,}원"


def _make_excel(contracts: list[dict], payout_rate: float) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "수수료 계산"
    headers = [
        "구분", "고객명", "보험회사", "상품명", "세부 조건", "월보험료", "지급율",
        "1차년계", "총수수료율", "잔여수수료율", "예상 익월수당", "예상 총수당", "예상 잔여수당",
        "출처 시트", "출처 행",
    ]
    ws.append(headers)

    for index, contract in enumerate(contracts, start=1):
        first_rate = contract["first_year_rate"] * payout_rate
        total_rate = contract["total_rate"] * payout_rate
        remaining_rate = total_rate - first_rate
        premium = contract["premium"]
        ws.append([
            index,
            contract.get("customer", ""),
            contract["insurer"],
            contract["product"],
            contract.get("conditions", ""),
            premium,
            payout_rate,
            first_rate,
            total_rate,
            remaining_rate,
            round(premium * first_rate),
            round(premium * total_rate),
            round(premium * remaining_rate),
            contract["sheet_name"],
            contract["row_number"],
        ])

    header_fill = PatternFill("solid", fgColor="2563D9")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in range(2, ws.max_row + 1):
        ws.cell(row, 6).number_format = "#,##0"
        for col in range(7, 11):
            ws.cell(row, col).number_format = "0.0%"
        for col in range(11, 14):
            ws.cell(row, col).number_format = "#,##0"

    widths = [8, 12, 16, 42, 34, 14, 11, 12, 14, 16, 17, 17, 17, 16, 10]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 26

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def _initialize_state() -> None:
    st.session_state.setdefault("commission_contracts", [])
    st.session_state.setdefault("commission_payout_rate", DEFAULT_PAYOUT_RATE)


def run() -> None:
    _initialize_state()

    st.title("수수료 계산기")
    st.caption("생보·손보 수수료 예시표에서 상품별 1차년계와 총수수료율을 불러옵니다.")

    with st.expander("① 수수료 예시표 불러오기", expanded=True):
        life_file = st.file_uploader(
            "생보 수수료 예시표", type=["xlsx"], key="commission_life_file"
        )
        nonlife_file = st.file_uploader(
            "손보 수수료 예시표", type=["xlsx"], key="commission_nonlife_file"
        )

    all_products: list[ProductRate] = []
    parse_warnings: list[str] = []
    for uploaded, source_type in ((life_file, "생보"), (nonlife_file, "손보")):
        if uploaded is None:
            continue
        try:
            parsed, warnings = parse_commission_workbook(uploaded.getvalue(), source_type)
            all_products.extend(_to_product_rate(item) for item in parsed)
            parse_warnings.extend(warnings)
        except Exception as exc:
            st.error(f"{source_type} 예시표를 읽지 못했습니다: {exc}")

    if all_products:
        insurer_count = len({product.insurer for product in all_products})
        st.success(f"보험회사 {insurer_count}개 · 수수료 조건 {len(all_products):,}개를 불러왔습니다.")
    else:
        st.info("생보 또는 손보 수수료 예시표를 올리면 상품을 선택할 수 있습니다.")

    for warning in parse_warnings:
        st.warning(warning)

    st.markdown("### ② 지급율 및 계약 입력")
    payout_rate_percent = st.number_input(
        "공통 지급율 (%)",
        min_value=0.0,
        max_value=100.0,
        value=float(st.session_state["commission_payout_rate"]),
        step=0.1,
        format="%.1f",
        help="변경한 지급율은 현재 작성 중인 모든 계약에 일괄 적용됩니다.",
    )
    st.session_state["commission_payout_rate"] = payout_rate_percent
    payout_rate = payout_rate_percent / 100

    source_options = [source for source in ("생보", "손보") if any(
        product.source_type == source for product in all_products
    )]
    source_type: str | None = None
    selected_product: ProductRate | None = None

    if source_options:
        source_type = st.radio(
            "보험 구분",
            options=source_options,
            format_func=lambda value: "생명보험" if value == "생보" else "손해보험",
            horizontal=True,
            key="commission_source_type",
        )
        insurer_names = sorted({
            product.insurer for product in all_products
            if product.source_type == source_type
        })
        insurer = st.selectbox(
            "보험회사",
            options=insurer_names,
            index=None,
            placeholder="보험회사를 선택하거나 검색해 주세요.",
            key="commission_insurer",
        )
    else:
        st.radio("보험 구분", options=["예시표를 먼저 올려 주세요."], disabled=True)
        insurer = None
        st.selectbox("보험회사", options=["예시표를 먼저 올려 주세요."], disabled=True)

    insurer_products = [
        product for product in all_products
        if insurer and product.source_type == source_type and product.insurer == insurer
    ]
    product_names = sorted({product.product for product in insurer_products})

    if product_names:
        product_name = st.selectbox(
            "상품",
            options=product_names,
            index=None,
            placeholder="상품명을 선택하거나 검색해 주세요.",
            key="commission_product_name",
        )
    else:
        product_name = None
        st.selectbox(
            "상품",
            options=["보험회사를 먼저 선택해 주세요."],
            disabled=True,
            key="commission_product_disabled",
        )

    condition_products = [
        product for product in insurer_products if product.product == product_name
    ]
    if condition_products:
        condition_map = {product.key: product for product in condition_products}
        selected_key = st.selectbox(
            "세부 조건",
            options=list(condition_map),
            format_func=lambda key: condition_map[key].conditions or "기본 조건",
            key="commission_condition",
        )
        selected_product = condition_map[selected_key]
        first_applied = selected_product.first_year_rate * payout_rate
        total_applied = selected_product.total_rate * payout_rate
        st.caption(
            f"적용 결과 · 1차년계 {_format_rate(first_applied)} · "
            f"총수수료율 {_format_rate(total_applied)}"
        )
    else:
        st.selectbox(
            "세부 조건",
            options=["상품을 먼저 선택해 주세요."],
            disabled=True,
            key="commission_condition_disabled",
        )

    customer_col, premium_col = st.columns(2)
    with customer_col:
        customer = st.text_input("고객명", key="commission_customer")
    with premium_col:
        premium = st.number_input(
            "월보험료", min_value=0, step=1000, value=0, format="%d",
            key="commission_premium",
        )

    if st.button("계약 추가", type="primary", use_container_width=True):
        if selected_product is None:
            st.warning("보험회사와 상품을 선택해 주세요.")
        elif premium <= 0:
            st.warning("월보험료를 입력해 주세요.")
        else:
            st.session_state["commission_contracts"].append({
                "customer": customer.strip(),
                "insurer": selected_product.insurer,
                "product": selected_product.product,
                "conditions": selected_product.conditions,
                "premium": int(premium),
                "first_year_rate": selected_product.first_year_rate,
                "total_rate": selected_product.total_rate,
                "sheet_name": selected_product.sheet_name,
                "row_number": selected_product.row_number,
            })
            st.rerun()

    contracts = st.session_state["commission_contracts"]
    st.markdown("### ③ 작성 중인 계약")
    if not contracts:
        st.info("추가된 계약이 없습니다.")
        return

    total_premium = sum(contract["premium"] for contract in contracts)
    total_first = sum(
        contract["premium"] * contract["first_year_rate"] * payout_rate
        for contract in contracts
    )
    total_commission = sum(
        contract["premium"] * contract["total_rate"] * payout_rate
        for contract in contracts
    )
    metric_cols = st.columns(3)
    metric_cols[0].metric("월보험료 합계", _format_won(total_premium))
    metric_cols[1].metric("예상 익월수당", _format_won(total_first))
    metric_cols[2].metric("예상 총수당", _format_won(total_commission))

    for index, contract in enumerate(contracts):
        first_rate = contract["first_year_rate"] * payout_rate
        total_rate = contract["total_rate"] * payout_rate
        title = f"{index + 1}. {contract.get('customer') or '고객명 없음'} · {contract['insurer']}"
        with st.container(border=True):
            st.markdown(f"**{title}**")
            st.caption(contract["product"] + (f" · {contract['conditions']}" if contract["conditions"] else ""))
            columns = st.columns(4)
            columns[0].metric("월보험료", _format_won(contract["premium"]))
            columns[1].metric("1차년계", _format_rate(first_rate))
            columns[2].metric("총수수료율", _format_rate(total_rate))
            columns[3].metric("예상 총수당", _format_won(contract["premium"] * total_rate))
            if st.button("삭제", key=f"delete_commission_{index}"):
                contracts.pop(index)
                st.rerun()

    st.divider()
    clear_col, download_col = st.columns([1, 2])
    with clear_col:
        if st.button("전체 계약 지우기", use_container_width=True):
            st.session_state["commission_contracts"] = []
            st.rerun()
    with download_col:
        excel_bytes = _make_excel(contracts, payout_rate)
        st.download_button(
            "엑셀 다운로드",
            data=excel_bytes,
            file_name="수수료_계산결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
