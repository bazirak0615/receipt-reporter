"""
OCR Engine Module — EasyOCR 기반 영수증 텍스트 추출
이미지 전처리(회전 보정, 대비 개선) 포함
"""
import easyocr
from pathlib import Path
from PIL import Image, ImageEnhance, ImageFilter, ExifTags
import tempfile
import os


def preprocess_image(image_path: str) -> str:
    """
    영수증 이미지 전처리 — OCR 정확도 향상
    1. EXIF 기반 자동 회전 보정
    2. 그레이스케일 변환
    3. 대비 향상
    4. 샤프닝
    Returns: 전처리된 이미지 임시 파일 경로
    """
    img = Image.open(image_path)

    # 1. EXIF 회전 보정 (스마트폰 사진)
    try:
        exif = img._getexif()
        if exif:
            for tag, value in exif.items():
                if ExifTags.TAGS.get(tag) == "Orientation":
                    if value == 3:
                        img = img.rotate(180, expand=True)
                    elif value == 6:
                        img = img.rotate(270, expand=True)
                    elif value == 8:
                        img = img.rotate(90, expand=True)
                    break
    except (AttributeError, KeyError, IndexError):
        pass

    # 2. 크기 조정 (너무 크면 OCR 느려짐, 너무 작으면 인식 안 됨)
    max_dim = 2000
    if max(img.size) > max_dim:
        ratio = max_dim / max(img.size)
        new_size = (int(img.size[0] * ratio), int(img.size[1] * ratio))
        img = img.resize(new_size, Image.LANCZOS)

    # 3. 그레이스케일 변환
    if img.mode != "L":
        img_gray = img.convert("L")
    else:
        img_gray = img

    # 4. 대비 향상
    enhancer = ImageEnhance.Contrast(img_gray)
    img_enhanced = enhancer.enhance(1.8)

    # 5. 샤프닝
    img_sharp = img_enhanced.filter(ImageFilter.SHARPEN)

    # 임시 파일로 저장
    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False, dir=tempfile.gettempdir())
    img_sharp.save(tmp.name)
    tmp.close()
    return tmp.name


class OCREngine:
    """EasyOCR 기반 영수증 인식 엔진"""

    def __init__(self, languages=None, gpu=False):
        self.languages = languages or ["ko", "en"]
        self.gpu = gpu
        self._reader = None

    @property
    def reader(self):
        if self._reader is None:
            self._reader = easyocr.Reader(self.languages, gpu=self.gpu)
        return self._reader

    def set_languages(self, languages: list):
        """OCR 언어 변경 (변경 시 리더 재초기화)"""
        if set(languages) != set(self.languages):
            self.languages = languages
            self._reader = None

    def extract_text(self, image_path: str) -> list[dict]:
        """이미지에서 텍스트 추출 (전처리 포함)"""
        # 전처리
        processed_path = preprocess_image(image_path)

        try:
            # 전처리된 이미지로 OCR
            results = self.reader.readtext(processed_path)

            # 결과가 적으면 원본으로 재시도
            if len(results) < 3:
                results_orig = self.reader.readtext(image_path)
                if len(results_orig) > len(results):
                    results = results_orig
        finally:
            # 임시 파일 정리
            try:
                os.unlink(processed_path)
            except OSError:
                pass

        extracted = []
        for bbox, text, confidence in results:
            text = text.strip()
            if text:
                extracted.append({
                    "text": text,
                    "confidence": round(confidence, 4),
                    "bbox": bbox
                })
        return extracted

    def extract_from_image(self, image_path: str) -> dict:
        """영수증 이미지에서 전체 텍스트 추출"""
        raw_results = self.extract_text(image_path)
        full_text = "\n".join([r["text"] for r in raw_results])

        return {
            "image_path": image_path,
            "full_text": full_text,
            "lines": raw_results,
            "avg_confidence": round(
                sum(r["confidence"] for r in raw_results) / max(len(raw_results), 1), 4
            )
        }
