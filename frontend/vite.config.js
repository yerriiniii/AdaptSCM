import path from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig, loadEnv } from "vite";
import react from "@vitejs/plugin-react";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// 프로젝트 루트(`InventorySystem/.env`) — uvicorn을 8001 등으로 쓰면
//   VITE_DEV_PROXY_API=http://127.0.0.1:8001
// 넣고 Vite(5173) 개발 서버를 재시작.

export default defineConfig(({ mode }) => {
  const projectRoot = path.resolve(__dirname, "..");
  const env = loadEnv(mode, projectRoot, "");
  const apiTarget = env.VITE_DEV_PROXY_API || "http://127.0.0.1:8000";

  return {
    plugins: [react()],
    server: {
      port: 5173,
      proxy: {
        "/api": {
          target: apiTarget,
          changeOrigin: true,
        },
      },
    },
  };
});
