## 실행 방법

### 1) 백엔드 (FastAPI)

`cd backend` 후:

`py -3 -m uvicorn main:app --reload --port 8000`

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

### 3) 주요 API (접두사 `/api/adaptscm`)

- `GET /view` : 저장된 일자·국가별 뷰
- `POST /aggregate` : 업로드 엑셀 집계
- `POST /shipment/aggregate` : 출고·파일 집계
- `GET /files` : 업로드 파일 목록
- `GET /purchase-orders` : 발주 목록
- (기타) SKU 매핑, S3 직접 업로드, 입고·발주 CRUD 엔드포인트

## 배포 메모

### 1) 백엔드 환경변수

프로젝트 루트 `.env`에 최소 아래 값이 필요합니다.

`DATABASE_URL`

`AWS_ACCESS_KEY_ID`

`AWS_SECRET_ACCESS_KEY`

`AWS_REGION`

`S3_BUCKET`

`FRONTEND_ORIGINS`

`ADMIN_ACCESS_CODE` (관리자 진입용 인증번호, 필수)

`JWT_SECRET` (운영 환경에서 긴 무작위 문자열 권장)

`JWT_EXPIRE_MINUTES` (선택, 기본 180 = 3시간)

예시:

`FRONTEND_ORIGINS=http://13.124.10.10,https://your-domain.com`

### 2) 프론트 환경변수

`frontend/.env.production` 파일에 아래처럼 둡니다.

`VITE_API_BASE_URL=`

빈 값으로 두면 프론트가 같은 도메인의 `/api/...`를 호출하므로 Nginx reverse proxy 배포에 맞습니다.

### 3) 프론트 빌드

`cd frontend`

`npm install`

`npm run build`

### 4) 백엔드 실행

`cd backend` 후:

`py -3 -m uvicorn main:app --host 0.0.0.0 --port 8000`

운영에서는 `systemd`로 상시 실행하는 것을 권장합니다.

### 5) Nginx 예시

프론트 빌드 결과물은 정적으로 서빙하고, `/api`는 백엔드로 프록시합니다.

```nginx
server {
    listen 80;
    server_name _;

    root /var/www/adaptscm;
    index index.html;

    location / {
        try_files $uri /index.html;
    }

    location /api/ {
        proxy_pass http://127.0.0.1:8000/api/;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }

    location /health {
        proxy_pass http://127.0.0.1:8000/health;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

### 6) 배포 후 확인

- `http://서버주소/health`
- 프론트 접속 후 파일 업로드
- `S3`, `uploaded_files`, `inventory_rows`, `inventory_aggregates` 반영 확인

