import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import axios from "axios";
import { Eye, EyeOff, KeyRound, Package } from "lucide-react";
import {
  API_BASE,
  AUTH_SESSION_CHECK_TIMEOUT_MS,
  getAccessToken,
  parseJwtExpiryMs,
  setAccessToken,
} from "./apiClient";

const SESSION_HOURS = 3;

function setSessionBannerHeightCssVar(px) {
  if (typeof document === "undefined") return;
  document.documentElement.style.setProperty(
    "--auth-session-banner-height",
    `${Math.max(0, Math.round(px) || 0)}px`
  );
}

function parseErrorMessage(err) {
  const st = err?.response?.status;
  const d = err?.response?.data?.detail;
  if (st === 422 && Array.isArray(d) && d.length) {
    const parts = d.map((x) => (x && (x.msg || x.message) ? String(x.msg || x.message) : "")).filter(Boolean);
    if (parts.length) return parts.join(" ");
  }
  if (d == null) return err?.message || "요청에 실패했습니다.";
  if (typeof d === "string") return d;
  if (typeof d === "object" && d.message) return String(d.message);
  return String(d);
}

function formatSessionRemaining(expMs) {
  const remSec = Math.max(0, Math.floor((expMs - Date.now()) / 1000));
  const hh = String(Math.floor(remSec / 3600)).padStart(2, "0");
  const mm = String(Math.floor((remSec % 3600) / 60)).padStart(2, "0");
  const ss = String(remSec % 60).padStart(2, "0");
  return `${hh}:${mm}:${ss}`;
}

function AuthGateHeroTitle() {
  return (
    <div className="heroHead">
      <h1 className="heroTitle heroTitleDash authGateHeroTitleDash">
        <span className="heroTitleMark authGateHeroTitleMark" aria-hidden="true">
          <Package size={40} strokeWidth={2} />
        </span>
        <span className="heroTitleText">
          <span className="heroTitleBrand">Adapt</span>
          <span className="heroTitleSub">상품 관리 보드</span>
        </span>
      </h1>
    </div>
  );
}

export default function AuthGate({ children }) {
  const [phase, setPhase] = useState("checking");
  const [bootError, setBootError] = useState("");
  const [accessCode, setAccessCode] = useState("");
  const [error, setError] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [showAccessCode, setShowAccessCode] = useState(false);
  const [sessionTick, setSessionTick] = useState(0);
  const sessionBannerRef = useRef(null);

  const expireSession = useCallback((message) => {
    setAccessToken(null);
    setPhase("unauth");
    setBootError(message);
  }, []);

  const trySession = useCallback(async () => {
    setBootError("");
    const t = getAccessToken();
    if (!t) {
      setPhase("unauth");
      return;
    }
    const expMs = parseJwtExpiryMs(t);
    if (expMs != null && Date.now() >= expMs) {
      expireSession("세션이 만료되었습니다. 인증번호를 다시 입력해 주세요.");
      return;
    }
    try {
      await axios.get(`${API_BASE}/api/auth/me`, { timeout: AUTH_SESSION_CHECK_TIMEOUT_MS });
      setPhase("authed");
    } catch (e) {
      const st = e?.response?.status;
      const aborted = e?.code === "ECONNABORTED" || /timeout/i.test(String(e?.message || ""));
      setAccessToken(null);
      setPhase("unauth");
      if (aborted) {
        setBootError("서버 응답이 지연되었습니다.\n서버 상태와 네트워크를 확인한 뒤 다시 시도해 주세요.");
      } else if (st === 401) {
        setBootError("세션이 만료되었거나 유효하지 않습니다. 인증번호를 다시 입력해 주세요.");
      } else if (e?.response == null) {
        setBootError("서버에 연결할 수 없습니다. API 주소·네트워크를 확인해 주세요.");
      } else {
        setBootError("접근 확인에 실패했습니다. 잠시 후 다시 시도해 주세요.");
      }
    }
  }, [expireSession]);

  useEffect(() => {
    void trySession();
  }, [trySession]);

  useEffect(() => {
    const onExpired = () => {
      expireSession("세션이 만료되었습니다. 인증번호를 다시 입력해 주세요.");
    };
    window.addEventListener("adaptscm-auth-expired", onExpired);
    return () => window.removeEventListener("adaptscm-auth-expired", onExpired);
  }, [expireSession]);

  useEffect(() => {
    if (phase !== "authed") return undefined;
    const t = getAccessToken();
    const expMs = parseJwtExpiryMs(t);
    if (expMs == null) return undefined;
    const delay = expMs - Date.now();
    if (delay <= 0) {
      expireSession("세션이 만료되었습니다. 인증번호를 다시 입력해 주세요.");
      return undefined;
    }
    const id = setTimeout(() => {
      expireSession("세션이 만료되었습니다. 인증번호를 다시 입력해 주세요.");
      window.dispatchEvent(new CustomEvent("adaptscm-auth-expired"));
    }, delay);
    return () => clearTimeout(id);
  }, [phase, expireSession]);

  useEffect(() => {
    if (phase !== "authed") return undefined;
    const id = setInterval(() => setSessionTick((n) => n + 1), 1000);
    return () => clearInterval(id);
  }, [phase]);

  const sessionRemainingLabel = useMemo(() => {
    if (phase !== "authed") return "";
    void sessionTick;
    const expMs = parseJwtExpiryMs(getAccessToken());
    if (expMs == null) return "";
    return formatSessionRemaining(expMs);
  }, [phase, sessionTick]);

  useEffect(() => {
    if (!sessionRemainingLabel) {
      setSessionBannerHeightCssVar(0);
      return undefined;
    }
    const el = sessionBannerRef.current;
    if (!el) {
      setSessionBannerHeightCssVar(0);
      return undefined;
    }
    const measure = () => setSessionBannerHeightCssVar(el.offsetHeight || 0);
    measure();
    let ro;
    if (typeof ResizeObserver !== "undefined") {
      ro = new ResizeObserver(measure);
      ro.observe(el);
    }
    window.addEventListener("resize", measure);
    return () => {
      window.removeEventListener("resize", measure);
      if (ro) ro.disconnect();
      setSessionBannerHeightCssVar(0);
    };
  }, [sessionRemainingLabel]);

  const onSubmitAccessCode = async (e) => {
    e.preventDefault();
    setError("");
    setBootError("");
    const code = String(accessCode || "").trim();
    if (!code) {
      setError("인증번호를 입력해 주세요.");
      return;
    }
    setSubmitting(true);
    try {
      const { data } = await axios.post(`${API_BASE}/api/auth/access-code`, { code });
      if (data?.access_token) {
        setAccessToken(data.access_token);
        setAccessCode("");
        setPhase("authed");
      } else {
        setError("접근 토큰을 받지 못했습니다.");
      }
    } catch (err) {
      setError(parseErrorMessage(err));
    } finally {
      setSubmitting(false);
    }
  };

  if (phase === "checking") {
    return (
      <div className="pageSplit authGatePage">
        <div className="authGateStack">
          <div className="dashboardHeroBand">
            <div className="headerArea">
              <section className="hero">
                <AuthGateHeroTitle />
                <p className="authGateSub">접근 확인 중…</p>
              </section>
            </div>
          </div>
        </div>
      </div>
    );
  }

  if (phase === "unauth") {
    return (
      <div className="pageSplit authGatePage">
        <div className="authGateStack">
          <div className="dashboardHeroBand">
            <div className="headerArea">
              <section className="hero">
                <AuthGateHeroTitle />
                <div className="authGateHeroText">
                  <p className="authGateSub" style={{ marginBottom: 0 }}>
                    관리자 인증번호를 입력하면 대시보드를 사용할 수 있습니다.
                  </p>
                  <p className="authGateSub authGateSubSecondary">
                    세션 유지 시간은 {SESSION_HOURS}시간이며, 만료 후에는 인증번호를 다시 입력해야 합니다.
                  </p>
                </div>
              </section>
            </div>
          </div>

          <div className="authGatePane">
            <div className="tableCard authGateCard">
              {bootError ? (
                <p
                  className="authGateMsgErr authGateBootMsg"
                  role="alert"
                  style={{ margin: "0 0 12px", whiteSpace: "pre-line" }}
                >
                  {bootError}
                </p>
              ) : null}
              <form onSubmit={onSubmitAccessCode} className="authGateForm authGateFormAccessCode">
                <div className="authGateField">
                  <span className="authGateFieldLabel">관리자 인증번호</span>
                  <div className="authGateInputInset">
                    <input
                      type={showAccessCode ? "text" : "password"}
                      autoComplete="off"
                      value={accessCode}
                      onChange={(ev) => setAccessCode(ev.target.value)}
                      placeholder="인증번호 입력"
                      required
                    />
                    <button
                      type="button"
                      className="authGateInputInsetToggle"
                      onClick={() => setShowAccessCode((prev) => !prev)}
                      aria-label={showAccessCode ? "인증번호 숨기기" : "인증번호 표시"}
                      aria-pressed={showAccessCode}
                    >
                      {showAccessCode ? (
                        <EyeOff size={22} strokeWidth={1.8} aria-hidden="true" />
                      ) : (
                        <Eye size={22} strokeWidth={1.8} aria-hidden="true" />
                      )}
                    </button>
                  </div>
                </div>
                {error ? <p className="authGateMsgErr" role="alert">{error}</p> : null}
                <button type="submit" className="primary authGateSubmit authGateSubmitWithIcon" disabled={submitting}>
                  <KeyRound size={18} strokeWidth={2} aria-hidden="true" />
                  {submitting ? "확인 중…" : "접속하기"}
                </button>
              </form>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <>
      {sessionRemainingLabel ? (
        <div ref={sessionBannerRef} className="authGateSessionBanner" role="status" aria-live="polite">
          세션 남은 시간 {sessionRemainingLabel}
        </div>
      ) : null}
      {children}
    </>
  );
}
