import streamlit as st
from dataclasses import dataclass
from typing import Tuple

EOK = 10_000  # 내부 계산 단위: 만원


def won_text(amount_manwon: float) -> str:
    won = max(0, round(float(amount_manwon) * 10_000))
    if won == 0:
        return "0원"
    eok, remainder = divmod(won, 100_000_000)
    man, one = divmod(remainder, 10_000)
    parts = []
    if eok:
        parts.append(f"{eok:,}억")
    if man:
        parts.append(f"{man:,}만")
    if one:
        parts.append(f"{one:,}")
    return " ".join(parts) + "원"


def tax_rate_and_deduction(tax_base: float) -> Tuple[float, float]:
    if tax_base <= 1 * EOK:
        return 0.10, 0
    if tax_base <= 5 * EOK:
        return 0.20, 1_000
    if tax_base <= 10 * EOK:
        return 0.30, 6_000
    if tax_base <= 30 * EOK:
        return 0.40, 16_000
    return 0.50, 46_000


def financial_asset_deduction(net_financial_assets: float) -> float:
    x = max(0.0, net_financial_assets)
    if x <= 2_000:
        return x
    if x <= 10_000:
        return 2_000
    if x <= 100_000:
        return x * 0.20
    return 20_000


def spouse_statutory_share(group: str, count: int) -> float:
    if group == "배우자 단독":
        return 1.0
    count = max(1, int(count))
    return 1.5 / (count + 1.5)


@dataclass
class Result:
    gross_estate: float
    taxable_estate: float
    personal_or_lump: float
    spouse_deduction: float
    financial_deduction: float
    home_deduction: float
    other_deduction: float
    deduction_before_limit: float
    deduction_limit: float
    allowed_deduction: float
    tax_base: float
    rate: float
    progressive_deduction: float
    calculated_tax: float
    generation_skip_surcharge: float
    tax_credits: float
    filing_credit: float
    estimated_tax_due: float


def calculate(**v) -> Result:
    gross = max(0, v["gross_estate"]) + max(0, v["deemed_estate"])
    expenses = max(0, v["public_dues"]) + max(0, v["funeral_expense"]) + max(0, v["liabilities"])
    prior_gifts = max(0, v["prior_gifts_heirs"]) + max(0, v["prior_gifts_non_heirs"])
    taxable_estate = max(0, gross - max(0, v["non_taxable"]) - expenses + prior_gifts)

    personal = (
        20_000
        + max(0, v["children_count"]) * 5_000
        + max(0, v["minor_deduction"])
        + max(0, v["elderly_count"]) * 5_000
        + max(0, v["disability_deduction"])
    )
    personal_or_lump = max(50_000, personal) if v["lump_mode"] else personal
    if v["spouse_exists"] and v["spouse_share"] >= 0.999:
        personal_or_lump = personal

    spouse_deduction = 0.0
    if v["spouse_exists"]:
        if v["spouse_actual_inheritance"] < 50_000:
            spouse_deduction = 50_000
        else:
            spouse_limit_base = max(
                0,
                gross + max(0, v["prior_gifts_heirs"])
                - max(0, v["non_heir_bequest"])
                - max(0, v["non_taxable"])
                - max(0, v["public_dues"])
                - max(0, v["liabilities"]),
            )
            spouse_limit = min(300_000, spouse_limit_base * min(1.0, max(0.0, v["spouse_share"])))
            spouse_deduction = min(max(0, v["spouse_actual_inheritance"]), spouse_limit)

    financial = financial_asset_deduction(v["net_financial_assets"])
    home = min(max(0, v["cohabiting_home_value"]), 60_000)
    deduction_before_limit = personal_or_lump + spouse_deduction + financial + home + max(0, v["other_deduction"])
    deduction_limit = max(
        0,
        taxable_estate
        - max(0, v["non_heir_bequest"])
        - max(0, v["inheritance_waiver_next_rank"])
        - max(0, v["prior_gift_tax_base_for_limit"]),
    )
    allowed = min(deduction_before_limit, deduction_limit)
    tax_base = max(0, taxable_estate - allowed - max(0, v["appraisal_fee"]))
    rate, progressive = tax_rate_and_deduction(tax_base)
    calculated_tax = max(0, tax_base * rate - progressive)

    gen_amount = min(max(0, v["generation_skip_amount"]), taxable_estate)
    gen_ratio = gen_amount / taxable_estate if taxable_estate else 0
    gen_rate = 0.40 if v["generation_skip_minor_over_2b"] else 0.30
    gen_surcharge = calculated_tax * gen_ratio * gen_rate

    before_credit = calculated_tax + gen_surcharge
    tax_credits = min(before_credit, max(0, v["gift_tax_credit"]) + max(0, v["other_tax_credit"]))
    after_credit = max(0, before_credit - tax_credits)
    filing_credit = after_credit * 0.03 if v["apply_filing_credit"] else 0

    return Result(
        gross, taxable_estate, personal_or_lump, spouse_deduction, financial, home,
        max(0, v["other_deduction"]), deduction_before_limit, deduction_limit, allowed,
        tax_base, rate, progressive, calculated_tax, gen_surcharge, tax_credits,
        filing_credit, max(0, after_credit - filing_credit)
    )


def run():
    st.title("🧾 상속세 예상 계산기")
    st.caption("금액 입력 단위: 만원 · 상담 및 사전 검토용 예상 계산기")

    with st.expander("계산 기준 및 유의사항"):
        st.markdown("""
        - 총상속재산에서 비과세 재산과 공과금·장례비용·채무를 차감하고 사전증여재산을 합산합니다.
        - 상속공제에는 일괄·인적공제, 배우자공제, 금융재산공제, 동거주택공제 등을 반영합니다.
        - 실제 신고세액은 상속관계, 재산평가, 사전증여 내역과 적용 요건에 따라 달라질 수 있습니다.
        """)

    tab1, tab2, tab3 = st.tabs(["① 기본 입력", "② 공제·상속인", "③ 고급 입력"])

    with tab1:
        c1, c2 = st.columns(2)
        with c1:
            gross_estate = st.number_input("상속재산가액", 0, value=100_000, step=1_000, key="it_gross")
            deemed_estate = st.number_input("추정·간주상속재산", 0, value=0, step=1_000, key="it_deemed")
            non_taxable = st.number_input("비과세·과세가액 불산입액", 0, value=0, step=1_000, key="it_nontax")
        with c2:
            public_dues = st.number_input("공과금", 0, value=0, step=100, key="it_dues")
            funeral_expense = st.number_input("공제대상 장례비용", 0, value=1_000, step=100, key="it_funeral")
            liabilities = st.number_input("피상속인 채무", 0, value=0, step=1_000, key="it_liab")
        c3, c4 = st.columns(2)
        with c3:
            prior_gifts_heirs = st.number_input("10년 이내 상속인 증여재산", 0, value=0, step=1_000, key="it_gift_h")
        with c4:
            prior_gifts_non_heirs = st.number_input("5년 이내 상속인 외 증여재산", 0, value=0, step=1_000, key="it_gift_n")

    with tab2:
        lump_mode = st.radio("공제 방식", ["일괄공제와 인적공제 중 큰 금액", "기초공제 + 인적공제"], horizontal=True, key="it_mode") == "일괄공제와 인적공제 중 큰 금액"
        a1, a2 = st.columns(2)
        with a1:
            children_count = st.number_input("자녀 수", 0, value=1, step=1, key="it_children")
            elderly_count = st.number_input("65세 이상 연로자 수", 0, value=0, step=1, key="it_elderly")
        with a2:
            minor_deduction = st.number_input("미성년자공제 합계액", 0, value=0, step=1_000, key="it_minor")
            disability_deduction = st.number_input("장애인공제 합계액", 0, value=0, step=1_000, key="it_disability")

        spouse_exists = st.toggle("배우자가 생존해 있음", value=True, key="it_spouse_exists")
        spouse_actual_inheritance, spouse_share = 0, 0.0
        if spouse_exists:
            b1, b2 = st.columns(2)
            with b1:
                group = st.selectbox("배우자와 공동상속하는 상속인", ["직계비속", "직계존속", "배우자 단독"], key="it_group")
                count = 0 if group == "배우자 단독" else st.number_input(f"{group} 공동상속인 수", 1, value=1, step=1, key="it_count")
            with b2:
                spouse_actual_inheritance = st.number_input("배우자가 실제 상속받는 금액", 0, value=50_000, step=1_000, key="it_spouse_amount")
                spouse_share = st.number_input("배우자 법정상속지분", 0.0, 1.0, float(spouse_statutory_share(group, count)), 0.01, format="%.4f", key="it_share")

        d1, d2 = st.columns(2)
        with d1:
            net_financial_assets = st.number_input("순금융재산가액", 0, value=0, step=1_000, key="it_fin")
            cohabiting_home_value = st.number_input("공제대상 동거주택가액", 0, value=0, step=1_000, key="it_home")
        with d2:
            other_deduction = st.number_input("가업·영농 등 기타 공제", 0, value=0, step=1_000, key="it_other_ded")
            appraisal_fee = st.number_input("감정평가수수료 공제액", 0, value=0, step=50, key="it_appraisal")

    with tab3:
        e1, e2, e3 = st.columns(3)
        with e1:
            non_heir_bequest = st.number_input("상속인 아닌 자에 대한 유증 등", 0, value=0, step=1_000, key="it_bequest")
        with e2:
            inheritance_waiver_next_rank = st.number_input("상속포기로 다음 순위가 받은 재산", 0, value=0, step=1_000, key="it_waiver")
        with e3:
            prior_gift_tax_base_for_limit = st.number_input("공제한도 차감 사전증여 과세표준", 0, value=0, step=1_000, key="it_prior_base")
        f1, f2 = st.columns(2)
        with f1:
            generation_skip_amount = st.number_input("세대를 건너뛴 상속재산가액", 0, value=0, step=1_000, key="it_gen")
            generation_skip_minor_over_2b = st.checkbox("미성년자가 세대생략으로 20억원 초과 상속", key="it_gen_minor")
        with f2:
            gift_tax_credit = st.number_input("사전증여 관련 증여세액공제", 0, value=0, step=100, key="it_gift_credit")
            other_tax_credit = st.number_input("기타 세액공제", 0, value=0, step=100, key="it_other_credit")
            apply_filing_credit = st.checkbox("기한 내 신고세액공제 3% 적용", value=True, key="it_filing")

    result = calculate(**locals())

    st.divider()
    st.subheader("계산 결과")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("상속세 과세가액", won_text(result.taxable_estate))
    m2.metric("상속공제 적용액", won_text(result.allowed_deduction))
    m3.metric("과세표준", won_text(result.tax_base))
    m4.metric("예상 납부세액", won_text(result.estimated_tax_due))

    if result.deduction_before_limit > result.allowed_deduction:
        st.warning(f"공제 합계 {won_text(result.deduction_before_limit)} 중 종합한도에 따라 {won_text(result.allowed_deduction)}만 적용되었습니다.")

    left, right = st.columns([1.25, 1])
    with left:
        st.dataframe({
            "항목": ["총상속재산", "상속세 과세가액", "일괄·인적공제", "배우자공제", "금융재산공제", "동거주택공제", "기타 공제", "공제 종합한도", "실제 공제액", "과세표준"],
            "금액": [won_text(x) for x in [result.gross_estate, result.taxable_estate, result.personal_or_lump, result.spouse_deduction, result.financial_deduction, result.home_deduction, result.other_deduction, result.deduction_limit, result.allowed_deduction, result.tax_base]],
        }, hide_index=True, use_container_width=True)
    with right:
        st.write(f"**적용세율:** {result.rate:.0%}")
        st.write(f"**누진공제:** {won_text(result.progressive_deduction)}")
        st.write(f"**산출세액:** {won_text(result.calculated_tax)}")
        st.write(f"**세대생략 할증:** {won_text(result.generation_skip_surcharge)}")
        st.write(f"**세액공제:** - {won_text(result.tax_credits)}")
        st.write(f"**신고세액공제:** - {won_text(result.filing_credit)}")
        st.success(f"예상 납부세액: {won_text(result.estimated_tax_due)}")

    st.info("상담용 예상치입니다. 실제 신고 시 재산평가와 공제 요건을 별도로 확인해야 합니다.")
