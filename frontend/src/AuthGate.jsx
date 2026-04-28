import { useCallback, useEffect, useMemo, useState } from "react";
import axios from "axios";
import { Check, Eye, EyeOff, LogIn, UserPlus } from "lucide-react";
import {
  API_BASE,
  AUTH_SESSION_CHECK_TIMEOUT_MS,
  getAccessToken,
  parseJwtExpiryMs,
  setAccessToken,
} from "./apiClient";

const SIGNUP_TOKEN_KEY = "inventory_signup_token";
const LS_POST_REGISTER_EMAIL = "inventory_post_register_login_email";
const LS_POST_REGISTER_STATE = "inventory_post_register_state";
/** `RegisterResponse` / 서버와 동일 */
const PENDING_OK_MSG = "가입 요청이 접수되었습니다. 관리자 승인 후 로그인할 수 있습니다.";
const APPROVED_OK_MSG = "승인 완료되었습니다.\n\n로그인을 진행해주세요.";
/** `password_policy._PASSWORD_SPECIALS` 와 동기화(변경 시 백엔드·MyPage와 같이 수정) */
const PASSWORD_SPECIALS_ALLOWED = `!"#$%&'()*+,-./:;<=>?@[]^_\`{|}~`;

function readPostRegisterEmailFromStorage() {
  try {
    return (localStorage.getItem(LS_POST_REGISTER_EMAIL) || "").trim();
  } catch {
    return "";
  }
}

function setPostRegisterStorage(email, state) {
  const em = String(email || "").trim();
  if (!em) {
    try {
      localStorage.removeItem(LS_POST_REGISTER_EMAIL);
      localStorage.removeItem(LS_POST_REGISTER_STATE);
    } catch {
      /* ignore */
    }
    return;
  }
  try {
    localStorage.setItem(LS_POST_REGISTER_EMAIL, em);
    if (state) localStorage.setItem(LS_POST_REGISTER_STATE, state);
  } catch {
    /* ignore */
  }
}

function clearPostRegisterStorage() {
  try {
    localStorage.removeItem(LS_POST_REGISTER_EMAIL);
    localStorage.removeItem(LS_POST_REGISTER_STATE);
  } catch {
    /* ignore */
  }
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
  if (typeof d === "object" && d.code === "pending_approval") return "관리자 승인 대기 중입니다. 승인 후 다시 로그인하세요.";
  if (typeof d === "object" && d.code === "rejected") return "가입이 거절된 계정입니다.";
  if (typeof d === "object" && d.code === "email_not_verified") {
    return "EMAIL_NOT_VERIFIED";
  }
  return String(d);
}

/**
 * regStep: request_email = 인증 링크 요청만 / password = 링크 클릭 후 비밀번호 입력
 */
export default function AuthGate({ children }) {
  const [phase, setPhase] = useState("checking");
  /** 로그인 확인 실패·만료 시 로그인 화면 상단 안내 */
  const [bootError, setBootError] = useState("");
  const [mode, setMode] = useState("login");
  const [regStep, setRegStep] = useState("request_email");
  const [email, setEmail] = useState(() => readPostRegisterEmailFromStorage());
  const [signupToken, setSignupToken] = useState("");
  const [password, setPassword] = useState("");
  const [password2, setPassword2] = useState("");
  const [error, setError] = useState("");
  const [okMessage, setOkMessage] = useState("");
  const [verifyBanner, setVerifyBanner] = useState({ kind: null, text: "" });
  const [manualSignupUrl, setManualSignupUrl] = useState("");
  const [lastSentToEmail, setLastSentToEmail] = useState("");
  const [linkSendBlockedEmail, setLinkSendBlockedEmail] = useState(null);
  /** 서버 `expires_at` 밀리초(인증 링크 보낸 뒤 남은 시간 표시) */
  const [linkDeadlineMs, setLinkDeadlineMs] = useState(null);
  const [linkTick, setLinkTick] = useState(0);
  const [submitting, setSubmitting] = useState(false);
  const [sendSubmitting, setSendSubmitting] = useState(false);
  const [showLoginPassword, setShowLoginPassword] = useState(false);
  const [showRegPassword, setShowRegPassword] = useState(false);
  const [showRegPassword2, setShowRegPassword2] = useState(false);
  const [postRegStatusSync, setPostRegStatusSync] = useState(0);

  const trySession = useCallback(async () => {
    setBootError("");
    const t = getAccessToken();
    if (!t) {
      setPhase("unauth");
      return;
    }
    const expMs = parseJwtExpiryMs(t);
    if (expMs != null && Date.now() >= expMs) {
      setAccessToken(null);
      setPhase("unauth");
      setBootError("로그인 유효 시간이 지났습니다. 다시 로그인해 주세요.");
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
        setBootError(
          "서버 응답이 지연되었습니다.\n서버 상태와 네트워크를 확인한 뒤 다시 로그인해 주세요.",
        );
      } else if (st === 401) {
        setBootError("로그인이 만료되었거나 유효하지 않습니다. 다시 로그인해 주세요.");
      } else if (e?.response == null) {
        setBootError("서버에 연결할 수 없습니다. API 주소·네트워크를 확인해 주세요.");
      } else {
        setBootError("로그인 확인에 실패했습니다. 잠시 후 다시 시도해 주세요.");
      }
    }
  }, []);

  useEffect(() => {
    void trySession();
  }, [trySession]);

  useEffect(() => {
    const onExpired = () => {
      setPhase("unauth");
      setBootError("로그인 유효 시간이 지났습니다. 다시 로그인해 주세요.");
    };
    window.addEventListener("inventory-auth-expired", onExpired);
    return () => window.removeEventListener("inventory-auth-expired", onExpired);
  }, []);

  useEffect(() => {
    if (linkDeadlineMs == null) return undefined;
    const id = setInterval(() => { setLinkTick((x) => x + 1); }, 1000);
    return () => clearInterval(id);
  }, [linkDeadlineMs]);

  useEffect(() => {
    if (phase !== "unauth" || mode !== "login") return undefined;
    const onVis = () => {
      if (document.visibilityState === "visible" && readPostRegisterEmailFromStorage()) {
        setPostRegStatusSync((n) => n + 1);
      }
    };
    document.addEventListener("visibilitychange", onVis);
    return () => { document.removeEventListener("visibilitychange", onVis); };
  }, [phase, mode]);

  /** 가입 직후·재방문 시 로그인 탭: localStorage 이메일 + 승인 상태에 맞는 안내 */
  useEffect(() => {
    if (phase !== "unauth" || mode !== "login") return undefined;
    const stored = readPostRegisterEmailFromStorage();
    if (stored && !email.trim()) {
      setEmail(stored);
      return undefined;
    }
    const em = email.trim().toLowerCase();
    if (!stored) {
      setOkMessage("");
      return undefined;
    }
    if (em !== stored.toLowerCase()) {
      setOkMessage("");
      return undefined;
    }
    const target = stored;
    let cancelled = false;
    (async () => {
      try {
        const { data } = await axios.get(`${API_BASE}/api/auth/account-status`, { params: { email: target } });
        if (cancelled) return;
        const st = data?.status;
        if (st === "active") {
          setOkMessage(APPROVED_OK_MSG);
          setPostRegisterStorage(target, "approved");
        } else if (st === "pending") {
          setOkMessage(PENDING_OK_MSG);
          setPostRegisterStorage(target, "pending");
        } else {
          setOkMessage("");
          clearPostRegisterStorage();
        }
      } catch {
        if (cancelled) return;
        try {
          const s = localStorage.getItem(LS_POST_REGISTER_STATE);
          if (s === "pending") setOkMessage(PENDING_OK_MSG);
          else if (s === "approved") setOkMessage(APPROVED_OK_MSG);
        } catch {
          setOkMessage(PENDING_OK_MSG);
        }
      }
    })();
    return () => { cancelled = true; };
  }, [phase, mode, email, postRegStatusSync]);

  const signupLinkCountdownEl = useMemo(() => {
    if (mode !== "register" || regStep !== "request_email") return null;
    if (!lastSentToEmail || email.trim().toLowerCase() !== lastSentToEmail) return null;
    if (linkDeadlineMs == null) return null;
    void linkTick;
    const rem = Math.max(0, Math.floor((linkDeadlineMs - Date.now()) / 1000));
    const mm = String(Math.floor(rem / 60)).padStart(2, "0");
    const ss = String(rem % 60).padStart(2, "0");
    if (rem <= 0) {
      return (
        <p
          className="authGateLinkCountdown authGateLinkCountdownExpired"
          role="status"
          style={{ marginTop: 8 }}
        >
          인증 링크 유효 시간이 지났습니다. 「인증 링크 재전송」을 눌러 주세요.
        </p>
      );
    }
    return (
      <p className="authGateLinkCountdown" role="status" style={{ marginTop: 8 }} aria-live="polite">
        링크 유효 남은 시간:{" "}
        <span className="authGateLinkCountdownClock" aria-label={`${mm}분 ${ss}초`}>
          {mm}:{ss}
        </span>
        {" "}(5분이 지나면 링크를 다시 보내야 합니다)
      </p>
    );
  }, [mode, regStep, lastSentToEmail, email, linkDeadlineMs, linkTick]);

  /** URL ?signup_token= — 이메일 링크로 들어온 경우: 토큰 확인 후 비밀번호 단계 */
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const tok = params.get("signup_token");
    if (!tok) {
      return;
    }
    let cancelled = false;
    (async () => {
      try {
        const { data } = await axios.get(`${API_BASE}/api/auth/signup-email/check`, { params: { token: tok } });
        if (cancelled) return;
        if (data?.valid && data?.email) {
          setMode("register");
          setRegStep("password");
          setEmail(String(data.email));
          setSignupToken(tok);
          try {
            sessionStorage.setItem(SIGNUP_TOKEN_KEY, tok);
          } catch {
            /* ignore */
          }
        }
      } catch (err) {
        if (!cancelled) {
          setMode("register");
          setVerifyBanner({ kind: "err", text: parseErrorMessage(err) || "인증 링크가 유효하지 않습니다." });
        }
      } finally {
        const u = new URL(window.location.href);
        u.searchParams.delete("signup_token");
        window.history.replaceState({}, "", `${u.pathname}${u.search}${u.hash}`);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  const onLogin = async (e) => {
    e.preventDefault();
    setError("");
    setOkMessage("");
    setBootError("");
    setSubmitting(true);
    try {
      const { data } = await axios.post(`${API_BASE}/api/auth/login`, { email, password });
      if (data?.access_token) {
        clearPostRegisterStorage();
        setAccessToken(data.access_token);
        setPhase("authed");
      } else {
        setError("토큰을 받지 못했습니다.");
      }
    } catch (err) {
      const p = parseErrorMessage(err);
      if (p === "EMAIL_NOT_VERIFIED") {
        setError("계정이 아직 이전 방식(미검증)으로 남아 있을 수 있습니다. 관리자에게 문의하세요.");
      } else {
        setError(p);
      }
    } finally {
      setSubmitting(false);
    }
  };

  const onSendSignupLink = async (e) => {
    e.preventDefault();
    setError("");
    setOkMessage("");
    const em = String(email).trim();
    if (!em) {
      setError("이메일을 입력하세요.");
      return;
    }
    const emLower = em.toLowerCase();
    if (linkSendBlockedEmail && emLower === linkSendBlockedEmail) {
      return;
    }
    setSendSubmitting(true);
    setManualSignupUrl("");
    try {
      const { data } = await axios.post(`${API_BASE}/api/auth/signup-email/send`, { email: em });
      setOkMessage(data?.message || "처리되었습니다.");
      setManualSignupUrl(data?.signup_url && !data?.mail_sent ? String(data.signup_url) : "");
      setLastSentToEmail(emLower);
      setLinkSendBlockedEmail(null);
      const rawExp = data?.expires_at;
      if (rawExp) {
        const ms = new Date(String(rawExp)).getTime();
        if (Number.isFinite(ms)) {
          setLinkDeadlineMs(ms);
        } else {
          setLinkDeadlineMs(Date.now() + 5 * 60 * 1000);
        }
      } else {
        setLinkDeadlineMs(Date.now() + 5 * 60 * 1000);
      }
    } catch (err) {
      const st = err?.response?.status;
      if (st === 429) {
        setLinkSendBlockedEmail(emLower);
      }
      setError(parseErrorMessage(err));
    } finally {
      setSendSubmitting(false);
    }
  };

  const onRegister = async (e) => {
    e.preventDefault();
    setError("");
    setOkMessage("");
    if (password !== password2) {
      setError("비밀번호가 일치하지 않습니다.");
      return;
    }
    const tok = signupToken || (typeof sessionStorage !== "undefined" ? sessionStorage.getItem(SIGNUP_TOKEN_KEY) : "");
    if (!tok) {
      setError("이메일 인증 링크를 먼저 받고, 메일에 있는 링크를 눌러 이 화면으로 들어온 뒤 가입을 완료하세요.");
      return;
    }
    setSubmitting(true);
    try {
      const { data } = await axios.post(`${API_BASE}/api/auth/register`, {
        email,
        password,
        password_confirm: password2,
        signup_token: tok,
      });
      setPassword("");
      setPassword2("");
      try {
        sessionStorage.removeItem(SIGNUP_TOKEN_KEY);
      } catch {
        /* ignore */
      }
      setSignupToken("");
      setRegStep("request_email");
      setPostRegisterStorage(email, "pending");
      setOkMessage(PENDING_OK_MSG);
      setMode("login");
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
                <div className="heroHead">
                  <h1 className="heroTitle authGateHeroTitle">재고 분석 대시보드</h1>
                </div>
                <p className="authGateSub">로그인 확인 중…</p>
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
                <div className="heroHead">
                  <h1 className="heroTitle authGateHeroTitle">재고 분석 대시보드</h1>
                </div>
                <div className="authGateHeroText">
                  <p className="authGateSub" style={{ marginBottom: 0 }}>
                    로그인 후 재고 분석 대시보드의 기능을 사용할 수 있습니다.
                  </p>
                  <p className="authGateSub authGateStepFlow">
                    <strong>1. 이메일 인증</strong>
                    <span className="authGateFlowArrow" aria-hidden="true">→</span>
                    <strong>2. 비밀번호 입력 후, 관리자 승인 대기</strong>
                    <span className="authGateFlowArrow" aria-hidden="true">→</span>
                    <strong>3. 관리자 승인 후 회원가입 완료</strong>
                  </p>
                </div>
              </section>
            </div>
          </div>

          {verifyBanner.kind ? (
            <div
              className="tableCard authGateCard authGateVerifyBanner"
              style={{
                borderColor: verifyBanner.kind === "ok" ? "#86efac" : "#fecaca",
                background: verifyBanner.kind === "ok" ? "#f0fdf4" : "#fef2f2",
              }}
            >
              <p
                className="authGateMsgOk"
                style={{ color: verifyBanner.kind === "ok" ? "#0b8a5d" : "#b91c1c", textAlign: "center" }}
                role="status"
              >
                {verifyBanner.text}
              </p>
            </div>
          ) : null}

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
              <div className="tabs">
                <button
                  type="button"
                  className={`tab tabWithIcon ${mode === "login" ? "active" : ""}`}
                  onClick={() => { setMode("login"); setError(""); setBootError(""); }}
                >
                  <LogIn className="tabIcon" size={24} strokeWidth={2} aria-hidden />
                  로그인
                </button>
                <button
                  type="button"
                  className={`tab tabWithIcon ${mode === "register" ? "active" : ""}`}
                  onClick={() => { setMode("register"); setError(""); setOkMessage(""); setBootError(""); }}
                >
                  <UserPlus className="tabIcon" size={24} strokeWidth={2} aria-hidden />
                  회원가입
                </button>
              </div>

              {mode === "login" ? (
                <form onSubmit={onLogin} className="authGateForm">
                  <div className="authGateField">
                    <span className="authGateFieldLabel">이메일</span>
                    <input
                      type="email"
                      value={email}
                      onChange={(ev) => setEmail(ev.target.value)}
                      autoComplete="email"
                      required
                    />
                  </div>
                  <div className="authGateField">
                    <span className="authGateFieldLabel">비밀번호</span>
                    <div className="authGatePasswordRow">
                      <input
                        type={showLoginPassword ? "text" : "password"}
                        value={password}
                        onChange={(ev) => setPassword(ev.target.value)}
                        autoComplete="current-password"
                        minLength={1}
                        required
                      />
                      <button
                        type="button"
                        className="authGatePasswordToggle"
                        onClick={() => { setShowLoginPassword((s) => !s); }}
                        aria-label={showLoginPassword ? "비밀번호 숨기기" : "비밀번호 표시"}
                        aria-pressed={showLoginPassword}
                      >
                        {showLoginPassword ? <EyeOff size={22} strokeWidth={1.8} /> : <Eye size={22} strokeWidth={1.8} />}
                      </button>
                    </div>
                  </div>
                  {okMessage ? <p className="authGateMsgOk" role="status">{okMessage}</p> : null}
                  {error ? <p className="authGateMsgErr" role="alert">{error}</p> : null}
                  <button type="submit" className="primary authGateSubmit" disabled={submitting}>
                    {submitting ? "처리 중…" : "로그인"}
                  </button>
                </form>
              ) : regStep === "request_email" ? (
                <div className="authGateForm">
                  <p className="authGateSub" style={{ marginTop: 0, maxWidth: "100%", marginBottom: 20 }}>
                    1. 먼저 이메일 주소를 인증해주세요
                  </p>
                  <div className="authGateField">
                    <span className="authGateFieldLabel">이메일</span>
                    <input
                      type="email"
                      value={email}
                      onChange={(ev) => {
                        const v = ev.target.value;
                        setEmail(v);
                        const n = v.trim().toLowerCase();
                        if (linkSendBlockedEmail && n !== linkSendBlockedEmail) {
                          setLinkSendBlockedEmail(null);
                        }
                        if (lastSentToEmail && n !== lastSentToEmail) {
                          setLastSentToEmail("");
                          setLinkDeadlineMs(null);
                        }
                      }}
                      autoComplete="email"
                      required
                    />
                  </div>
                  {okMessage ? <p className="authGateMsgOk" role="status" style={{ marginTop: 8 }}>{okMessage}</p> : null}
                  {signupLinkCountdownEl}
                  {manualSignupUrl ? (
                    <div className="authGateManualUrl" style={{ marginTop: 12 }}>
                      <span className="authGateFieldLabel">인증 링크(복사)</span>
                      <div className="authGateManualUrlRow">
                        <input type="text" readOnly value={manualSignupUrl} className="authGateManualUrlInput" />
                        <button
                          type="button"
                          className="ghost authGateResendBtn"
                          onClick={() => {
                            void navigator.clipboard.writeText(manualSignupUrl);
                          }}
                        >
                          복사
                        </button>
                      </div>
                    </div>
                  ) : null}
                  {error ? <p className="authGateMsgErr" role="alert">{error}</p> : null}
                  <div className="authGateResendRow" style={{ marginTop: 24 }}>
                    <button
                      type="button"
                      className="primary authGateSubmit"
                      style={{ maxWidth: 360 }}
                      disabled={sendSubmitting || (
                        !!linkSendBlockedEmail
                        && email.trim().toLowerCase() === linkSendBlockedEmail
                      )}
                      onClick={onSendSignupLink}
                    >
                      {sendSubmitting
                        ? "보내는 중…"
                        : (lastSentToEmail && email.trim().toLowerCase() === lastSentToEmail
                          ? "인증 링크 재전송"
                          : "인증 링크 전송")}
                    </button>
                  </div>
                </div>
              ) : (
                <form onSubmit={onRegister} className="authGateForm">
                    <div className="authGateField">
                      <span className="authGateFieldLabel authGateFieldLabelWithCheck">
                        <Check className="authGateFieldCheckIcon" size={18} strokeWidth={2.2} aria-hidden />
                        인증된 이메일
                      </span>
                      <input type="text" value={email} readOnly className="authGateReadonlyEmail" />
                      <button
                        type="button"
                        className="authGateRegReemailLink"
                        onClick={() => {
                          setRegStep("request_email");
                          setSignupToken("");
                          setPassword("");
                          setPassword2("");
                          setOkMessage("");
                          setError("");
                          try {
                            sessionStorage.removeItem(SIGNUP_TOKEN_KEY);
                          } catch {
                            /* ignore */
                          }
                        }}
                      >
                        다른 이메일로 재인증 하기
                      </button>
                    </div>
                    <div className="authGateField">
                      <span className="authGateFieldLabel authGateFieldLabelWithCheck">
                        <Check className="authGateFieldCheckIcon" size={18} strokeWidth={2.2} aria-hidden />
                        새 비밀번호
                      </span>
                      <div className="authGatePasswordRow">
                        <input
                          type={showRegPassword ? "text" : "password"}
                          value={password}
                          onChange={(ev) => setPassword(ev.target.value)}
                          autoComplete="new-password"
                          minLength={8}
                          maxLength={16}
                          required
                        />
                        <button
                          type="button"
                          className="authGatePasswordToggle"
                          onClick={() => { setShowRegPassword((s) => !s); }}
                          aria-label={showRegPassword ? "비밀번호 숨기기" : "비밀번호 표시"}
                          aria-pressed={showRegPassword}
                        >
                          {showRegPassword ? <EyeOff size={22} strokeWidth={1.8} /> : <Eye size={22} strokeWidth={1.8} />}
                        </button>
                      </div>
                    </div>
                    <div className="authGateRegPwHint" aria-label="비밀번호 정책">
                      <p className="authGateRegPwHintRule">
                        8~16자, 영문 대·소문자, 숫자, 아래 특수문자(각 1자 이상), 공백 사용 불가
                      </p>
                      <p className="authGateRegPwHintChars" title={PASSWORD_SPECIALS_ALLOWED}>
                        {PASSWORD_SPECIALS_ALLOWED}
                      </p>
                    </div>
                    <div className="authGateField">
                      <span className="authGateFieldLabel authGateFieldLabelWithCheck">
                        <Check className="authGateFieldCheckIcon" size={18} strokeWidth={2.2} aria-hidden />
                        비밀번호 확인
                      </span>
                      <div className="authGatePasswordRow">
                        <input
                          type={showRegPassword2 ? "text" : "password"}
                          value={password2}
                          onChange={(ev) => setPassword2(ev.target.value)}
                          autoComplete="new-password"
                          minLength={8}
                          maxLength={16}
                          required
                        />
                        <button
                          type="button"
                          className="authGatePasswordToggle"
                          onClick={() => { setShowRegPassword2((s) => !s); }}
                          aria-label={showRegPassword2 ? "비밀번호 숨기기" : "비밀번호 표시"}
                          aria-pressed={showRegPassword2}
                        >
                          {showRegPassword2 ? <EyeOff size={22} strokeWidth={1.8} /> : <Eye size={22} strokeWidth={1.8} />}
                        </button>
                      </div>
                    </div>
                    {error ? <p className="authGateMsgErr" role="alert">{error}</p> : null}
                    <button type="submit" className="primary authGateSubmit" disabled={submitting}>
                      {submitting ? "처리 중…" : "가입 요청"}
                    </button>
                </form>
              )}
            </div>
          </div>
        </div>
      </div>
    );
  }

  return children;
}
