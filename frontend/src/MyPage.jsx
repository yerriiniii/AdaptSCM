import { useCallback, useEffect, useState } from "react";
import axios from "axios";
import { API_BASE } from "./apiClient";
import { ArrowLeft, Eye, EyeOff, KeyRound, Shield, User } from "lucide-react";

/** `password_policy._PASSWORD_SPECIALS` 와 동기화(변경 시 백엔드와 같이 수정) */
const PASSWORD_SPECIALS_ALLOWED = `!"#$%&'()*+,-./:;<=>?@[]^_\`{|}~`;

function parseDetail(err) {
  const st = err?.response?.status;
  const d = err?.response?.data?.detail;
  if (st === 422 && Array.isArray(d) && d.length) {
    const parts = d.map((x) => (x && (x.msg || x.message) ? String(x.msg || x.message) : "")).filter(Boolean);
    if (parts.length) return parts.join(" ");
  }
  if (d == null) return err?.message || "요청에 실패했습니다.";
  if (typeof d === "string") return d;
  return String(d);
}

/**
 * @param {{ onBack: () => void }} props
 */
export default function MyPage({ onBack }) {
  const [me, setMe] = useState(null);
  const [meError, setMeError] = useState("");
  const [pendingItems, setPendingItems] = useState([]);
  const [pendingLoadError, setPendingLoadError] = useState("");
  const [pendingLoading, setPendingLoading] = useState(false);
  const [actioningId, setActioningId] = useState(null);

  const [curPw, setCurPw] = useState("");
  const [newPw, setNewPw] = useState("");
  const [newPw2, setNewPw2] = useState("");
  const [showCur, setShowCur] = useState(false);
  const [showNew, setShowNew] = useState(false);
  const [showNew2, setShowNew2] = useState(false);
  const [pwSubmitting, setPwSubmitting] = useState(false);
  const [pwMsg, setPwMsg] = useState("");
  const [pwErr, setPwErr] = useState("");

  const loadMe = useCallback(async () => {
    setMeError("");
    try {
      const { data } = await axios.get(`${API_BASE}/api/auth/me`);
      setMe(data);
    } catch (err) {
      setMeError(parseDetail(err));
    }
  }, []);

  const loadPending = useCallback(async () => {
    setPendingLoadError("");
    setPendingLoading(true);
    try {
      const { data } = await axios.get(`${API_BASE}/api/auth/admin/pending-users`);
      setPendingItems(data?.items || []);
    } catch (err) {
      setPendingLoadError(parseDetail(err));
    } finally {
      setPendingLoading(false);
    }
  }, []);

  useEffect(() => {
    void loadMe();
  }, [loadMe]);

  useEffect(() => {
    if (me?.is_admin) {
      void loadPending();
    } else {
      setPendingItems([]);
    }
  }, [me, loadPending]);

  const onChangePassword = async (e) => {
    e.preventDefault();
    setPwMsg("");
    setPwErr("");
    setPwSubmitting(true);
    try {
      const { data } = await axios.post(`${API_BASE}/api/auth/change-password`, {
        current_password: curPw,
        new_password: newPw,
        new_password_confirm: newPw2,
      });
      setPwMsg(data?.message || "비밀번호가 변경되었습니다.");
      setCurPw("");
      setNewPw("");
      setNewPw2("");
    } catch (err) {
      setPwErr(parseDetail(err));
    } finally {
      setPwSubmitting(false);
    }
  };

  const onApprove = async (userId) => {
    setActioningId(userId);
    setPendingLoadError("");
    try {
      await axios.post(`${API_BASE}/api/auth/admin/users/${userId}/approve`);
      await loadPending();
    } catch (err) {
      setPendingLoadError(parseDetail(err));
    } finally {
      setActioningId(null);
    }
  };

  const onReject = async (userId) => {
    if (!window.confirm("이 가입 요청을 거절할까요?")) return;
    setActioningId(userId);
    setPendingLoadError("");
    try {
      await axios.post(`${API_BASE}/api/auth/admin/users/${userId}/reject`);
      await loadPending();
    } catch (err) {
      setPendingLoadError(parseDetail(err));
    } finally {
      setActioningId(null);
    }
  };

  return (
    <div className="mypagePageShell">
      <div className="mypageDashboard">
        <div className="dashboardHeroBand">
          <div className="headerArea mypageHeaderArea">
            <section className="hero mypageHeroTop">
              <div className="mypageHeroRow">
                <div className="mypageHeroRowStart">
                  <button type="button" className="cautionBtn mypageBackBtn" onClick={onBack}>
                    <ArrowLeft size={16} strokeWidth={2} aria-hidden />
                    대시보드로
                  </button>
                </div>
                <h1 className="heroTitle mypageMainTitle">마이페이지</h1>
                <div className="mypageHeroRowEnd" aria-hidden="true" />
              </div>
            </section>
          </div>
        </div>

        {meError ? (
          <div className="mypageBannerErr" role="alert">
            {meError}
          </div>
        ) : null}

        {me && !meError ? (
          <div className="mypageContent">
            <div className="mypageContentGrid">
            <section className="tableCard mypagePanel mypageAccountPanel">
              <div className="mypageSubSection">
                <div className="mypageProfileTitleRow">
                  <h3 className="mypageBlockTitle">
                    <User className="mypageKickerIcon" size={18} strokeWidth={2} aria-hidden />
                    프로필
                  </h3>
                  {me.is_admin ? (
                    <span className="mypageAdminBadge" role="status">
                      <Shield size={13} strokeWidth={2} aria-hidden />
                      관리자
                    </span>
                  ) : null}
                </div>
                <dl className="mypageDl">
                  <div className="mypageDlRow">
                    <dt className="mypageFieldLabel mypageEmailDt">이메일</dt>
                    <dd title={me.email}>{me.email}</dd>
                  </div>
                </dl>
              </div>

              <div className="mypageSectionRule" role="separator" />

              <div className="mypageSubSection mypageSubSectionForm">
                <h3 className="mypageBlockTitle">
                  <KeyRound className="mypageKickerIcon" size={18} strokeWidth={2} aria-hidden />
                  비밀번호 변경
                </h3>
                <form onSubmit={onChangePassword} className="mypagePwForm">
                  <div className="authGateField mypageFormField">
                    <span className="authGateFieldLabel mypageFieldLabel">현재 비밀번호</span>
                    <div className="authGatePasswordRow">
                      <input
                        type={showCur ? "text" : "password"}
                        value={curPw}
                        onChange={(ev) => setCurPw(ev.target.value)}
                        autoComplete="current-password"
                        required
                      />
                      <button
                        type="button"
                        className="authGatePasswordToggle"
                        onClick={() => { setShowCur((s) => !s); }}
                        aria-label={showCur ? "비밀번호 숨기기" : "비밀번호 표시"}
                        aria-pressed={showCur}
                      >
                        {showCur ? <EyeOff size={22} strokeWidth={1.8} /> : <Eye size={22} strokeWidth={1.8} />}
                      </button>
                    </div>
                  </div>
                  <div className="authGateField mypageFormField">
                    <span className="authGateFieldLabel mypageFieldLabel">새 비밀번호</span>
                    <div className="authGatePasswordRow">
                      <input
                        type={showNew ? "text" : "password"}
                        value={newPw}
                        onChange={(ev) => setNewPw(ev.target.value)}
                        autoComplete="new-password"
                        minLength={8}
                        maxLength={16}
                        required
                      />
                      <button
                        type="button"
                        className="authGatePasswordToggle"
                        onClick={() => { setShowNew((s) => !s); }}
                        aria-label={showNew ? "비밀번호 숨기기" : "비밀번호 표시"}
                        aria-pressed={showNew}
                      >
                        {showNew ? <EyeOff size={22} strokeWidth={1.8} /> : <Eye size={22} strokeWidth={1.8} />}
                      </button>
                    </div>
                  </div>
                  <div className="authGateField mypageFormField">
                    <span className="authGateFieldLabel mypageFieldLabel">새 비밀번호 확인</span>
                    <div className="authGatePasswordRow">
                      <input
                        type={showNew2 ? "text" : "password"}
                        value={newPw2}
                        onChange={(ev) => setNewPw2(ev.target.value)}
                        autoComplete="new-password"
                        minLength={8}
                        maxLength={16}
                        required
                      />
                      <button
                        type="button"
                        className="authGatePasswordToggle"
                        onClick={() => { setShowNew2((s) => !s); }}
                        aria-label={showNew2 ? "비밀번호 숨기기" : "비밀번호 표시"}
                        aria-pressed={showNew2}
                      >
                        {showNew2 ? <EyeOff size={22} strokeWidth={1.8} /> : <Eye size={22} strokeWidth={1.8} />}
                      </button>
                    </div>
                  </div>
                  {pwMsg ? <p className="mypageFormMsgOk" role="status">{pwMsg}</p> : null}
                  {pwErr ? <p className="mypageFormMsgErr" role="alert">{pwErr}</p> : null}
                  <p className="mypageHint">
                    8~16자, 영문 대·소문자·숫자, 아래 특수문자(최소 1자), 공백 없음
                  </p>
                  <p className="mypageHint mypageHintSpecials" title={PASSWORD_SPECIALS_ALLOWED}>
                    <span className="mypageHintSpecialsLabel">허용 특수문자: </span>
                    <span className="mypageHintSpecialsChars">{PASSWORD_SPECIALS_ALLOWED}</span>
                  </p>
                  <div className="mypageFormActions">
                    <button type="submit" className="primary mypagePrimaryBtn" disabled={pwSubmitting}>
                      {pwSubmitting ? "처리 중…" : "비밀번호 변경"}
                    </button>
                  </div>
                </form>
              </div>
            </section>

            {me.is_admin ? (
              <section className="tableCard mypagePanel mypagePanelAdmin">
                <header className="mypagePanelHeadAdmin">
                  <h2 className="mypagePanelKicker mypagePanelKickerAdmin">
                    <Shield className="mypageKickerIcon" size={18} strokeWidth={2} aria-hidden />
                    가입 승인
                  </h2>
                </header>
                {pendingLoadError ? <p className="mypageFormMsgErr" role="alert">{pendingLoadError}</p> : null}
                {pendingLoading && !pendingItems.length ? (
                  <p className="mypageEmpty">불러오는 중…</p>
                ) : !pendingItems.length ? (
                  <p className="mypageEmpty">승인 대기 중인 계정이 없습니다.</p>
                ) : (
                  <div className="tableWrap mypageTableWrap">
                    <table className="mypageDataTable">
                      <thead>
                        <tr>
                          <th>이메일</th>
                          <th>가입 요청일</th>
                          <th className="mypageColActions"> </th>
                        </tr>
                      </thead>
                      <tbody>
                        {pendingItems.map((row) => (
                          <tr key={row.id}>
                            <td className="mypageCellEmail">{row.email}</td>
                            <td>
                              {row.created_at
                                ? new Date(row.created_at).toLocaleString("ko-KR", {
                                    year: "numeric",
                                    month: "2-digit",
                                    day: "2-digit",
                                    hour: "2-digit",
                                    minute: "2-digit",
                                  })
                                : "—"}
                            </td>
                            <td className="mypageCellActions">
                              <div className="mypageRowActions">
                                <button
                                  type="button"
                                  className="primary"
                                  disabled={actioningId === row.id}
                                  onClick={() => { void onApprove(row.id); }}
                                >
                                  {actioningId === row.id ? "처리 중…" : "승인"}
                                </button>
                                <button
                                  type="button"
                                  className="ghost"
                                  disabled={actioningId === row.id}
                                  onClick={() => { void onReject(row.id); }}
                                >
                                  거절
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </section>
            ) : null}
            </div>
          </div>
        ) : null}
      </div>
    </div>
  );
}
