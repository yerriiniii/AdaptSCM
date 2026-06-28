import axios from "axios";

/** Vite dev: `vite.config`에서 `/api` → `http://127.0.0.1:8000` 프록시. 배포: nginx 동일 출처. */
export const API_BASE = import.meta.env.VITE_API_BASE ?? "";

const TOKEN_KEY = "adaptscm_access_token";

/** 초기 `/api/auth/me` 등: 무한 대기 방지 */
export const AUTH_SESSION_CHECK_TIMEOUT_MS = 8_000;

export function getAccessToken() {
  if (typeof localStorage === "undefined") return null;
  return localStorage.getItem(TOKEN_KEY);
}

export function setAccessToken(token) {
  if (typeof localStorage === "undefined") return;
  if (token) localStorage.setItem(TOKEN_KEY, token);
  else localStorage.removeItem(TOKEN_KEY);
}

export function clearAccessToken() {
  setAccessToken(null);
}

/** JWT payload `exp`(초) → 만료 시각(ms). 파싱 실패 시 null. */
export function parseJwtExpiryMs(token) {
  if (!token || typeof token !== "string") return null;
  try {
    const parts = token.split(".");
    if (parts.length < 2) return null;
    let b64 = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    const pad = b64.length % 4 ? "=".repeat(4 - (b64.length % 4)) : "";
    const json = JSON.parse(atob(b64 + pad));
    const exp = json?.exp;
    return typeof exp === "number" && Number.isFinite(exp) ? exp * 1000 : null;
  } catch {
    return null;
  }
}

function isPresignedS3Url(url) {
  if (!url || typeof url !== "string") return false;
  // Presigned PUT: 서명이 쿼리에 있음. JWT Authorization 헤더를 넣으면 S3가 400 InvalidArgument 반환
  return url.includes("amazonaws.com") || url.includes("X-Amz-Algorithm=");
}

axios.interceptors.request.use((config) => {
  const url = String(config.url || "");
  if (isPresignedS3Url(url)) {
    if (config.headers) {
      delete config.headers.Authorization;
    }
    return config;
  }
  const t = getAccessToken();
  if (t) {
    config.headers = config.headers || {};
    config.headers.Authorization = `Bearer ${t}`;
  }
  return config;
});

function isAuthPublicRequestUrl(url) {
  const u = String(url || "");
  return u.includes("/api/auth/access-code");
}

axios.interceptors.response.use(
  (res) => res,
  (err) => {
    if (err?.response?.status !== 401) return Promise.reject(err);
    if (isAuthPublicRequestUrl(err?.config?.url)) return Promise.reject(err);
    if (err?.config?.headers?.Authorization) {
      clearAccessToken();
      if (typeof window !== "undefined") {
        window.dispatchEvent(new CustomEvent("adaptscm-auth-expired"));
      }
    }
    return Promise.reject(err);
  },
);
