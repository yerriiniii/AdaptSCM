import React from "react";
import ReactDOM from "react-dom/client";
import "./apiClient";
import App from "./App";
import AuthGate from "./AuthGate";
import "./styles.css";

/** Tauri(WebView2) — sticky 레이아웃·스크롤 루트 보정용 */
if (typeof window !== "undefined" && ("__TAURI_INTERNALS__" in window || "__TAURI__" in window)) {
  document.documentElement.classList.add("tauri-desktop");
}

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <AuthGate>
      <App />
    </AuthGate>
  </React.StrictMode>
);

