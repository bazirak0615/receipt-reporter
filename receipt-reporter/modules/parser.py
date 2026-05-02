"""
Parser Module — OCR 텍스트에서 구조화된 영수증 데이터 추출
날짜, 금액, 업체명, 통화, 결제수단 등을 정규식 + 키워드로 파싱
"""
import re
from datetime import datetime


# 통화 기호/코드 매핑
CURRENCY_PATTERNS = {
    "KRW": [r"₩", r"원", r"KRW"],
    "USD": [r"\$", r"USD", r"US\$"],
    "JPY": [r"¥", r"円", r"JPY"],
    "EUR": [r"€", r"EUR"],
    "CNY": [r"¥", r"元", r"CNY", r"RMB"],
    "THB": [r"฿", r"THB", r"บาท"],
    "VND": [r"₫", r"VND", r"đ"],
    "GBP": [r"£", r"GBP"],
    "SGD": [r"S\$", r"SGD"],
    "HKD": [r"HK\$", r"HKD"],
    "TWD": [r"NT\$", r"TWD"],
    "AUD": [r"A\$", r"AUD"],
}

# 카테고리 키워드 매핑 (통합 버전: 조식/중식/석식 → 식비)
CATEGORY_KEYWORDS = {
    "T01": ["항공", "airline", "air", "flight", "대한항공", "아시아나", "제주항공", "진에어",
            "delta", "united", "ana", "jal", "peach", "空港", "airport", "搭乗"],
    "T02": ["ktx", "srt", "기차", "열차", "고속버스", "시외버스", "train", "rail", "bus",
            "jr ", "新幹線", "鉄道", "乗車券", "station"],
    "T03": ["택시", "taxi", "uber", "lyft", "grab", "카카오택시", "타다",
            "タクシー", "迎車"],
    "T04": ["렌터카", "렌트카", "rental", "rent-a-car", "hertz", "avis", "budget",
            "レンタカー", "orix", "times"],
    "T05": ["주유", "주유소", "gas", "fuel", "통행료", "toll", "하이패스", "주차",
            "ガソリン", "parking", "駐車", "高速", "etс"],
    "T06": ["지하철", "metro", "subway", "교통카드", "suica", "pasmo", "icoca", "交通"],
    "A01": ["호텔", "hotel", "hilton", "marriott", "hyatt", "inn", "resort", "리조트", "숙박",
            "ホテル", "旅館", "宿泊", "チェックイン", "check-in", "room", "sancoinn",
            "toyoko", "apa ", "dormy", "villa", "lodge", "accommodation"],
    "A02": ["모텔", "게스트하우스", "airbnb", "에어비앤비", "민박", "hostel", "ゲストハウス"],
    "M01": ["식비", "식사", "음식", "레스토랑", "restaurant", "food", "meal", "dining",
            "조식", "중식", "석식", "breakfast", "lunch", "dinner", "아침", "점심", "저녁",
            "커피", "카페", "cafe", "coffee", "스타벅스", "starbucks", "음료",
            "편의점", "convenience", "コンビニ", "セブン", "ファミリ", "ローソン",
            "食事", "飲食", "弁当", "ランチ", "ディナー", "レストラン", "居酒屋",
            "ラーメン", "寿司", "うどん", "そば", "カフェ", "喫茶",
            "bakery", "パン", "le pan", "sweets", "菓子"],
    "M02": ["자판기", "vending", "自販機", "간식", "snack"],
    "E01": ["접대", "회식", "entertainment", "거래처", "接待"],
    "C01": ["로밍", "국제전화", "roaming", "wifi", "와이파이", "인터넷", "ポケット"],
    "R01": ["등록비", "참가비", "컨퍼런스", "세미나", "conference", "seminar", "registration"],
    "S01": ["인쇄", "복사", "문구", "사무용품", "print", "copy", "stationery"],
    "O99": [],
}

# 날짜 패턴 (우선순위 순)
DATE_PATTERNS = [
    (r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", "YMD"),           # 2026-03-25, 2026.03.25
    (r"(\d{2})[-./](\d{1,2})[-./](\d{1,2})\s+\d{1,2}:\d{2}", "YMD_SHORT"),  # 25/09/13 14:24
    (r"(\d{1,2})[-./](\d{1,2})[-./](\d{4})", "XYZ"),           # 03/25/2026 또는 25/03/2026
    (r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日", "YMD"),          # 2025年09月13日 (일본어)
    (r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일", "YMD"),          # 2026년 3월 25일
    (r"(\d{1,2})월\s*(\d{1,2})일", "MD"),                       # 3월 25일
    (r"(\w{3})\s+(\d{1,2}),?\s+(\d{4})", "MDY_EN"),            # Mar 25, 2026
    (r"(\d{1,2})\s+(\w{3})\s+(\d{4})", "DMY_EN"),              # 25 Mar 2026
]

MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

# 금액 패턴 (한국어 + 영어 + 일본어)
AMOUNT_PATTERNS = [
    # 한국어
    r"합\s*계\s*[:\s]*([0-9,]+)",
    r"합계금액\s*[:\s]*([0-9,]+)",
    r"결제\s*금액\s*[:\s]*([0-9,]+)",
    r"청구\s*금액\s*[:\s]*([0-9,]+)",
    r"카드\s*결제\s*[:\s]*([0-9,]+)",
    r"승인\s*금액\s*[:\s]*([0-9,]+)",
    # 일본어
    r"合\s*計\s*[:\s]*¥?\s*([0-9,]+)",
    r"合計金額\s*[:\s]*¥?\s*([0-9,]+)",
    r"お支払[い]?\s*[:\s]*¥?\s*([0-9,]+)",
    r"お買上合計\s*[:\s]*¥?\s*([0-9,]+)",
    r"税込\s*[:\s]*¥?\s*([0-9,]+)",
    r"金\s*額\s*[:\s]*¥?\s*([0-9,]+)",
    r"TOTAL\s*[:\s]*¥?\s*([0-9,]+)",
    # ¥ 기호 뒤의 금액 (일본 영수증 핵심 패턴)
    r"¥\s*([0-9,]{3,})",
    r"\\([0-9,]{3,})",
    # 영어
    r"total\s*[:\s]*\$?\s*([0-9,.]+)",
    r"amount\s*[:\s]*\$?\s*([0-9,.]+)",
    r"grand\s*total\s*[:\s]*\$?\s*([0-9,.]+)",
]

# 결제수단 키워드
PAYMENT_KEYWORDS = {
    "법인카드": ["법인", "법인카드", "corporate"],
    "개인카드": ["신용카드", "체크카드", "개인카드", "credit", "debit", "visa", "master", "amex",
                "mastercard", "クレジット", "カード", "card"],
    "현금": ["현금", "cash", "현금영수증", "現金"],
    "계좌이체": ["이체", "transfer", "계좌", "振込"],
}

# 증빙 유형 키워드
RECEIPT_TYPE_KEYWORDS = {
    "TAX_INVOICE": ["세금계산서", "tax invoice"],
    "CASH_RECEIPT": ["현금영수증"],
    "CARD_SLIP": ["카드매출전표", "카드전표", "승인번호", "approval"],
    "SIMPLIFIED": ["간이영수증", "간이"],
}


class ReceiptParser:
    """OCR 텍스트에서 구조화된 영수증 데이터를 추출하는 파서"""

    def __init__(self, default_currency="KRW", date_format_hint="YMD"):
        self.default_currency = default_currency
        self.date_format_hint = date_format_hint

    def parse(self, ocr_result: dict) -> dict:
        """
        OCR 결과를 구조화된 영수증 데이터로 변환

        Args:
            ocr_result: OCREngine.extract_from_image()의 반환값

        Returns:
            구조화된 영수증 데이터 dict
        """
        full_text = ocr_result.get("full_text", "")
        lines = ocr_result.get("lines", [])
        text_lower = full_text.lower()

        parsed = {
            "image_path": ocr_result.get("image_path", ""),
            "vendor_name": self._extract_vendor(lines),
            "date": self._extract_date(full_text),
            "total_amount": self._extract_amount(full_text),
            "currency": self._detect_currency(full_text),
            "supply_amount": self._extract_supply_amount(full_text),
            "vat_amount": self._extract_vat(full_text),
            "payment_method": self._detect_payment(text_lower),
            "receipt_type": self._detect_receipt_type(text_lower),
            "category": self._classify_category(text_lower),
            "business_reg_no": self._extract_business_no(full_text),
            "card_approval_no": self._extract_approval_no(full_text),
            "confidence": ocr_result.get("avg_confidence", 0),
            "raw_text": full_text,
        }

        # 공급가액/부가세 미추출 시 역산
        if parsed["total_amount"] and not parsed["supply_amount"]:
            if parsed["currency"] == "KRW":
                parsed["supply_amount"] = round(parsed["total_amount"] / 1.1)
                parsed["vat_amount"] = parsed["total_amount"] - parsed["supply_amount"]

        return parsed

    def _extract_vendor(self, lines: list) -> str:
        """업체명 추출 — 보통 영수증 상단 첫 몇 줄에 위치"""
        if not lines:
            return ""
        # 상단 3줄 중 가장 긴 텍스트를 업체명으로 추정
        top_lines = lines[:5]
        candidates = []
        for line in top_lines:
            text = line["text"]
            # 숫자만 있거나, 날짜/금액 패턴이면 건너뜀
            if re.match(r"^[\d\s,./:-]+$", text):
                continue
            if len(text) < 2:
                continue
            candidates.append(text)
        return candidates[0] if candidates else ""

    def _extract_date(self, text: str) -> str:
        """날짜 추출 — 다중 포맷 지원"""
        for pattern, fmt in DATE_PATTERNS:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                try:
                    return self._normalize_date(match, fmt)
                except (ValueError, KeyError):
                    continue
        return ""

    def _normalize_date(self, match, fmt: str) -> str:
        """날짜를 YYYY-MM-DD 형식으로 정규화"""
        groups = match.groups()

        if fmt == "YMD":
            y, m, d = int(groups[0]), int(groups[1]), int(groups[2])
        elif fmt == "YMD_SHORT":
            # 25/09/13 형식 → 2025/09/13
            y, m, d = int(groups[0]), int(groups[1]), int(groups[2])
            if y < 100:
                y += 2000
        elif fmt == "XYZ":
            a, b, y = int(groups[0]), int(groups[1]), int(groups[2])
            if self.date_format_hint == "MDY":
                m, d = a, b
            else:
                d, m = a, b
            # 자동 보정: 월이 12 초과면 반전
            if m > 12:
                m, d = d, m
        elif fmt == "MD":
            m, d = int(groups[0]), int(groups[1])
            y = datetime.now().year
        elif fmt == "MDY_EN":
            m = MONTH_MAP.get(groups[0].lower()[:3], 1)
            d, y = int(groups[1]), int(groups[2])
        elif fmt == "DMY_EN":
            d = int(groups[0])
            m = MONTH_MAP.get(groups[1].lower()[:3], 1)
            y = int(groups[2])
        else:
            return ""

        if y < 100:
            y += 2000

        date_obj = datetime(y, m, d)
        return date_obj.strftime("%Y-%m-%d")

    def _parse_amount_str(self, amount_str: str) -> float | None:
        """금액 문자열을 숫자로 변환 (소수점 통화 지원)"""
        amount_str = amount_str.strip()
        # 쉼표가 천단위 구분자인지 소수점인지 판별
        # 예: 1,234.56 (영미권) / 1.234,56 (유럽) / 12,500 (한국)
        if re.match(r"^\d{1,3}(,\d{3})*(\.\d{1,2})?$", amount_str):
            # 영미권: 1,234.56 또는 12,500
            return float(amount_str.replace(",", ""))
        elif re.match(r"^\d{1,3}(\.\d{3})*(,\d{1,2})?$", amount_str):
            # 유럽식: 1.234,56
            return float(amount_str.replace(".", "").replace(",", "."))
        else:
            # 기타: 쉼표/점 모두 제거 후 정수 변환
            cleaned = amount_str.replace(",", "").replace(".", "")
            return float(cleaned) if cleaned.isdigit() else None

    def _to_int_safe(self, amount_str: str) -> int | None:
        """금액 문자열 → 정수 (소수점 반올림)"""
        val = self._parse_amount_str(amount_str)
        return round(val) if val is not None else None

    def _extract_amount(self, text: str) -> int | None:
        """합계 금액 추출 (소수점 통화 대응)"""
        for pattern in AMOUNT_PATTERNS:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                val = self._parse_amount_str(match.group(1))
                if val is not None:
                    # KRW는 정수, 외화는 소수점 유지 위해 *100 하지 않고 반올림
                    return round(val) if self.default_currency == "KRW" else round(val, 2)

        # 폴백: 텍스트에서 가장 큰 숫자를 합계로 추정
        all_numbers = re.findall(r"(\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?)", text)
        if all_numbers:
            numbers = []
            for n in all_numbers:
                v = self._parse_amount_str(n)
                if v is not None:
                    numbers.append(v)
            if numbers:
                val = max(numbers)
                return round(val) if self.default_currency == "KRW" else round(val, 2)

        return None

    def _extract_supply_amount(self, text: str) -> int | None:
        """공급가액 추출"""
        patterns = [
            r"공급가[액]\s*[:\s]*([0-9,]+)",
            r"과세\s*물품\s*[:\s]*([0-9,]+)",
            r"subtotal\s*[:\s]*([0-9,.]+)",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                val = self._parse_amount_str(match.group(1))
                if val is not None:
                    return round(val)
        return None

    def _extract_vat(self, text: str) -> int | None:
        """부가세액 추출 (소수점 통화 대응)"""
        patterns = [
            r"부가세\s*[:\s]*([0-9,]+)",
            r"부가가치세\s*[:\s]*([0-9,]+)",
            r"vat\s*[:\s]*([0-9,.]+)",
            r"tax\s*[:\s]*([0-9,.]+)",
            r"消費税\s*[:\s]*¥?\s*([0-9,]+)",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                val = self._parse_amount_str(match.group(1))
                if val is not None:
                    return round(val)
        return None

    def _detect_currency(self, text: str) -> str:
        """통화 감지"""
        for currency, patterns in CURRENCY_PATTERNS.items():
            for pattern in patterns:
                if re.search(pattern, text):
                    # CNY와 JPY 모두 ¥ 사용 → 한자/글자로 구분
                    if currency in ("CNY", "JPY") and pattern == r"¥":
                        if re.search(r"[元人民币]", text):
                            return "CNY"
                        if re.search(r"[円税込]", text):
                            return "JPY"
                        continue
                    return currency
        return self.default_currency

    def _detect_payment(self, text: str) -> str:
        """결제수단 감지"""
        for method, keywords in PAYMENT_KEYWORDS.items():
            for kw in keywords:
                if kw in text:
                    return method
        return "법인카드"  # 기본값

    def _detect_receipt_type(self, text: str) -> str:
        """증빙 유형 감지"""
        for rtype, keywords in RECEIPT_TYPE_KEYWORDS.items():
            for kw in keywords:
                if kw in text:
                    return rtype
        return "CARD_SLIP"  # 기본값

    def _classify_category(self, text: str) -> str:
        """비용 카테고리 자동 분류"""
        best_match = "O99"
        best_score = 0
        for category, keywords in CATEGORY_KEYWORDS.items():
            score = sum(1 for kw in keywords if kw in text)
            if score > best_score:
                best_score = score
                best_match = category
        return best_match

    def _extract_business_no(self, text: str) -> str:
        """사업자등록번호 추출 (XXX-XX-XXXXX)"""
        match = re.search(r"(\d{3})-?(\d{2})-?(\d{5})", text)
        if match:
            return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
        return ""

    def _extract_approval_no(self, text: str) -> str:
        """카드 승인번호 추출"""
        patterns = [
            r"승인\s*번호\s*[:\s]*(\d{6,10})",
            r"approval\s*#?\s*[:\s]*(\d{6,10})",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1)
        return ""
