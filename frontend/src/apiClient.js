import axios from "axios";

/** Vite dev: `vite.config`에서 `/api` → `http://127.0.0.1:8000` 프록시. 배포: nginx 동일 출처. */
export const API_BASE = import.meta.env.VITE_API_BASE ?? "";

const TOKEN_KEY = "inventory_access_token";

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
