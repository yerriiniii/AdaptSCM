# AdaptSCM 데스크톱 (Tauri) 배포

프론트만 Windows 설치형으로 포장하고, API·DB·S3는 기존 서버를 그대로 사용합니다.

## 사전 요구사항 (빌드 PC)

- Node.js 18+
- Rust (`rustup`) — 이미 설치됨
- Visual Studio 2022 (C++ 빌드 도구)
- WebView2 (Windows 10/11에 보통 기본 포함)

## 1) API 주소 설정

```powershell
cd frontend
copy .env.desktop.example .env.desktop
```

`.env.desktop`의 `VITE_API_BASE`를 팀 서버 주소로 바꿉니다.

```
VITE_API_BASE=https://adaptscm.com
```

웹 Nginx처럼 동일 출처가 아니므로 **반드시 절대 URL**이어야 합니다.

## 2) 서버 CORS

서버 `.env`의 `FRONTEND_ORIGINS`에 아래를 포함하세요.

```
http://tauri.localhost,https://tauri.localhost,tauri://localhost
```

백엔드 기본값에도 이미 포함되어 있지만, `.env`로 덮어쓰는 경우 직접 추가해야 합니다.

## 3) 개발 실행 (로컬 창)

백엔드가 `8000`에서 떠 있는 상태에서:

```powershell
cd frontend
npm install
npm run tauri:dev
```

개발 모드에서는 Vite 프록시(`/api` → `127.0.0.1:8000`)를 그대로 씁니다.

## 4) 설치 파일 빌드

```powershell
cd frontend
npm run tauri:build
```

산출물:

- `frontend/src-tauri/target/release/bundle/nsis/*.exe` — NSIS 설치 프로그램
- `frontend/src-tauri/target/release/bundle/msi/*.msi` — MSI 설치 프로그램

팀 PC에는 `.exe` 또는 `.msi`를 배포하면 됩니다.

## 참고

- 서버 URL이 바뀌면 `.env.desktop`을 수정한 뒤 **다시 빌드**해야 합니다.
- 아이콘은 `frontend/src-tauri/icons/`에서 교체할 수 있습니다.
