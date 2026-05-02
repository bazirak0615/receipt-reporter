# Receipt Reporter — 출장 영수증 자동 보고서 생성기

출장 후 영수증 사진/PDF를 한꺼번에 업로드하면 OCR로 금액·통화·일자·가맹점을 인식하고 카테고리를 자동 분류해 **부가세 매입세액 공제 가능 여부까지 판정한 정산 보고서**(Excel/Word/PDF 동시)를 만들어주는 로컬 FastAPI 도구.

> 12개 통화 · 17개 카테고리 · 국내/해외 출장 모두 대응. EasyOCR 로컬 동작으로 영수증 정보가 외부로 나가지 않음.

자세한 기획은 [docs/기획안.md](docs/기획안.md). 사용 흐름 상세는 [`receipt-reporter/사용자 가이드.html`](receipt-reporter/사용자%20가이드.html).

## 핵심 기능

- 이미지/PDF 다중 업로드 (한 번에 최대 50건/10MB)
- 통화에 따라 OCR 언어팩 자동 전환 (KRW→ko, JPY→ja, CNY→ch_sim 등)
- 17개 비용 카테고리 자동 분류 (교통·숙박·식비·접대비·통신·등록·업무용품·기타)
- 부가세 공제 룰 엔진 (해외/접대비/면세/적격증빙 자동 판정)
- 적격증빙 검증 (세금계산서·현금영수증·카드매출전표 / 일반 30,000원 · 접대비 10,000원 임계)
- 보고서 3종 동시 출력 (Excel 카테고리별 시트 / Word 결재 양식 / PDF 인쇄용)
- 세션 영속화로 중단 후 이어 작업 가능

## 빠른 시작

```bash
cd receipt-reporter
pip install -r requirements.txt
python app.py
# → http://127.0.0.1:8500 자동 오픈
```

또는 Windows에서는 `run.bat` 더블클릭.

> 첫 실행 시 EasyOCR이 모델 파일(약 100MB)을 자동 다운로드합니다 (1~2분 소요).

## 설정

[`receipt-reporter/config.json`](receipt-reporter/config.json)에서 통화 매핑·카테고리·업로드 한도 조정 가능. 기본값 그대로 사용해도 충분합니다.

## 데이터 보관

- `data/uploads/` — 업로드된 영수증 (gitignore 처리, 외부 노출 금지)
- `data/output/` — 생성된 보고서 (gitignore 처리)
- `data/sessions.json` — 세션 영속화 (gitignore 처리)

## 기술 스택

- Python 3.10+ · FastAPI · Uvicorn
- EasyOCR (PyTorch 기반, CPU/GPU 옵션)
- openpyxl · python-docx · reportlab
- Jinja2 + Vanilla JS

## 한계

- EasyOCR은 클라우드 OCR(Google Vision 등) 대비 정확도가 낮음 → 일부 영수증은 수동 보정 필요
- 한국 세법 기준 (해외 세무 미지원)
- 단일 사용자 로컬 도구 (멀티유저 동시성 미지원)

## 라이선스

개인 프로젝트 — 자유롭게 참고·재사용 가능.
