## 실행 방법

### 1) 백엔드 (FastAPI)

프로젝트 루트에서:

`py -3 -m uvicorn backend.main:app --reload --port 8000`

정상 실행 확인:

`http://localhost:8000/health`

### 2) 프론트 (React)

`cd frontend`

최초 1회:

`npm install`

실행:

`npm run dev`

브라우저:

`http://localhost:5173`

### 3) 주요 API

- `GET /api/reorder/template` : 엑셀 템플릿 다운로드
- `POST /api/reorder/plan` : 발주 추천(JSON)
- `POST /api/reorder/plan.xlsx` : 발주 추천 엑셀 다운로드
- `POST /api/inventory/aggregate` : 일자별 재고 파일 통합/피벗

