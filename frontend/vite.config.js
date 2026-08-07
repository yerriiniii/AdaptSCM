import path from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig, loadEnv } from "vite";
import react from "@vitejs/plugin-react";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const host = process.env.TAURI_DEV_HOST;

// 프로젝트 루트(`AdaptSCM/.env`) — uvicorn을 8001 등으로 쓰면
//   VITE_DEV_PROXY_API=http://127.0.0.1:8001
// 넣고 Vite(5173) 개발 서버를 재시작.

export default defineConfig(({ mode }) => {
  const projectRoot = path.resolve(__dirname, "..");
  const env = loadEnv(mode, projectRoot, "");
  const localEnv = loadEnv(mode, __dirname, "");
  const apiTarget =
    localEnv.VITE_DEV_PROXY_API || env.VITE_DEV_PROXY_API || "http://127.0.0.1:8000";

  return {
    plugins: [react()],
    clearScreen: false,
    envPrefix: ["VITE_", "TAURI_ENV_"],
    server: {
      port: 5173,
      strictPort: true,
      host: host || false,
      hmr: host
        ? {
            protocol: "ws",
            host,
            port: 1421,
          }
        : undefined,
      watch: {
        ignored: ["**/src-tauri/**"],
      },
      proxy: {
        "/api": {
          target: apiTarget,
          changeOrigin: true,
          // 상품마스터 대량 업로드 등 장시간 요청
          timeout: 300_000,
          proxyTimeout: 300_000,
        },
      },
    },
    build: {
      target: process.env.TAURI_ENV_PLATFORM === "windows" ? "chrome105" : "safari13",
      minify: !process.env.TAURI_ENV_DEBUG ? "esbuild" : false,
      sourcemap: !!process.env.TAURI_ENV_DEBUG,
    },
  };
});
