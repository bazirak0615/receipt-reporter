"""
Categorizer Module — 비용 카테고리 자동 분류 및 세무 검증
"""


# 카테고리 대분류
CATEGORY_GROUPS = {
    "T": "교통비",
    "A": "숙박비",
    "M": "식비",
    "E": "접대비",
    "C": "통신비",
    "R": "등록/참가비",
    "S": "업무용품",
    "O": "기타",
}


def get_category_group(category_code: str) -> str:
    """카테고리 코드에서 대분류 반환"""
    if not category_code:
        return "기타"
    prefix = category_code[0]
    return CATEGORY_GROUPS.get(prefix, "기타")


def is_vat_deductible(receipt: dict, trip_type: str = "domestic") -> bool:
    """
    부가세 매입세액 공제 가능 여부 판별

    Args:
        receipt: 파싱된 영수증 데이터
        trip_type: "domestic" 또는 "overseas"
    """
    # 해외 출장 → 전액 불공제
    if trip_type == "overseas":
        return False

    # 적격증빙 없음 → 불공제
    receipt_type = receipt.get("receipt_type", "NONE")
    if receipt_type in ("SIMPLIFIED", "NONE"):
        return False

    # 접대비 → 불공제
    category = receipt.get("category", "")
    if category.startswith("E"):
        return False

    # 부가세 0원 → 불공제 (면세 거래)
    vat = receipt.get("vat_amount")
    if not vat or vat == 0:
        return False

    return True


def check_qualified_receipt(receipt: dict) -> dict:
    """
    적격증빙 검증

    Returns:
        {"is_qualified": bool, "risk_level": str, "message": str}
    """
    receipt_type = receipt.get("receipt_type", "NONE")
    total = receipt.get("total_amount", 0) or 0
    category = receipt.get("category", "")

    # 접대비 1만원 기준, 일반 3만원 기준
    threshold = 10000 if category.startswith("E") else 30000

    qualified_types = {"TAX_INVOICE", "CASH_RECEIPT", "CARD_SLIP"}

    if receipt_type in qualified_types:
        return {
            "is_qualified": True,
            "risk_level": "safe",
            "message": "적격증빙 수취 완료"
        }

    if receipt_type == "SIMPLIFIED" and total <= threshold:
        return {
            "is_qualified": True,
            "risk_level": "safe",
            "message": f"간이영수증 ({total:,}원 ≤ {threshold:,}원 이하 인정)"
        }

    if total > threshold:
        return {
            "is_qualified": False,
            "risk_level": "warning",
            "message": f"적격증빙 미수취 ({total:,}원 > {threshold:,}원) — 가산세 2% 위험"
        }

    return {
        "is_qualified": False,
        "risk_level": "info",
        "message": "증빙 미확인"
    }


def calculate_summary(receipts: list, trip_type: str = "domestic", exchange_rate: float = 1.0, exchange_rates: dict = None) -> dict:
    """
    영수증 목록에서 요약 정보 산출 (복수 환율 지원)

    Args:
        exchange_rate: 기본 환율 (단일 환율 폴백용)
        exchange_rates: 통화별 환율 dict (예: {"JPY": 9.2, "USD": 1350})
    """
    if exchange_rates is None:
        exchange_rates = {}

    category_totals = {}
    payment_totals = {}
    total_amount = 0
    total_vat_deductible = 0
    total_vat_non_deductible = 0
    qualified_count = 0
    risk_items = []

    for r in receipts:
        amount = r.get("total_amount", 0) or 0
        currency = r.get("currency", "KRW")
        # 통화별 환율 적용 (복수 환율 우선, 없으면 기본 환율)
        if currency == "KRW":
            krw_amount = amount
        else:
            rate = exchange_rates.get(currency, exchange_rate)
            krw_amount = round(amount * rate)

        # 카테고리별 합산
        cat_group = get_category_group(r.get("category", "O99"))
        category_totals[cat_group] = category_totals.get(cat_group, 0) + krw_amount

        # 결제수단별 합산
        payment = r.get("payment_method", "기타")
        payment_totals[payment] = payment_totals.get(payment, 0) + krw_amount

        total_amount += krw_amount

        # 세무 검증
        vat = r.get("vat_amount", 0) or 0
        if is_vat_deductible(r, trip_type):
            total_vat_deductible += vat
        else:
            total_vat_non_deductible += vat

        qual = check_qualified_receipt(r)
        if qual["is_qualified"]:
            qualified_count += 1
        if qual["risk_level"] == "warning":
            risk_items.append(r)

    return {
        "total_amount": total_amount,
        "total_count": len(receipts),
        "category_totals": category_totals,
        "payment_totals": payment_totals,
        "vat_deductible": total_vat_deductible,
        "vat_non_deductible": total_vat_non_deductible,
        "qualified_count": qualified_count,
        "qualified_rate": round(qualified_count / max(len(receipts), 1) * 100, 1),
        "risk_items_count": len(risk_items),
        "trip_type": trip_type,
    }
