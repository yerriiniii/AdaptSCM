import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from html import escape

from shared.config import get_runtime_settings

_log = logging.getLogger(__name__)


def send_signup_email_proof_link(*, to_email: str, signup_url: str) -> bool:
    """가입 1단계: 이메일 소유 확인 링크(비밀번호는 링크 후 화면에서 입력)."""
    s = get_runtime_settings()
    subject = "[재고 시스템] 이메일 인증 — 가입을 계속하세요"
    plain = (
        "안녕하세요.\n\n"
        "아래 버튼을 눌러 이메일 주소를 인증해 주세요.\n"
        "이후 회원가입을 요청 하시면 관리자 승인 후 회원가입 처리가 완료됩니다.\n\n"
        f"인증 링크: {signup_url}"
    )
    html = f"""<p>안녕하세요.</p>
<p>아래 버튼을 눌러 이메일 주소를 인증해 주세요.</p>
<p>이후 회원가입을 요청 하시면 관리자 승인 후 회원가입 처리가 완료됩니다.</p>
<div style="margin-top:32px;">
<a href="{escape(signup_url)}" style="display:inline-block;padding:12px 18px;background:#0f8a5f;color:#fff;
text-decoration:none;border-radius:8px;font-weight:600;">이메일 인증하고 가입 이어가기</a>
</div>"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    from_addr = s.email_smtp_from or s.email_smtp_user or ""
    if from_addr:
        msg["From"] = from_addr
    else:
        msg["From"] = to_email
    msg["To"] = to_email
    msg.attach(MIMEText(plain, "plain", "utf-8"))
    msg.attach(MIMEText(html, "html", "utf-8"))

    if not s.email_smtp_host or not s.email_smtp_user:
        _log.warning(
            "EMAIL_SMTP_HOST / EMAIL_SMTP_USER 가 없어 메일을 보내지 않습니다. "
            "가입 인증 URL(프론트 JSON 또는 아래): %s",
            signup_url,
        )
        return False

    try:
        with smtplib.SMTP(s.email_smtp_host, s.email_smtp_port, timeout=30) as smtp:
            if s.email_smtp_use_tls:
                smtp.starttls()
            if s.email_smtp_password:
                smtp.login(s.email_smtp_user, s.email_smtp_password)
            out_from = s.email_smtp_from or s.email_smtp_user
            if not out_from:
                _log.error("EMAIL_SMTP_FROM / EMAIL_SMTP_USER 이 없어 발신 주소를 정할 수 없습니다.")
                return False
            smtp.sendmail(out_from, [to_email], msg.as_string())
    except Exception as exc:  # noqa: BLE001 — SMTP·소켓 등
        _log.error("SMTP 전송 실패: %s", exc, exc_info=True)
        return False
    return True


def send_signup_verification_email(*, to_email: str, verify_url: str) -> None:
    """(레거시 호환) — 필요 시 `send_signup_email_proof_link` 사용."""
    s = get_runtime_settings()
    subject = "[재고 시스템] 이메일 주소를 확인해 주세요"
    plain = (
        "안녕하세요.\n\n"
        "아래 링크를 눌러 이메일 인증을 완료한 뒤, 관리자 승인을 기다려 주세요.\n"
        f"{verify_url}\n\n"
        "본인이 요청한 가입이 아니면 이 메일을 무시하시면 됩니다."
    )
    html = f"""<p>안녕하세요.</p>
<p>아래 버튼(또는 링크)로 이메일 인증을 완료한 뒤, <strong>관리자 승인</strong>을 기다려 주세요.</p>
<p><a href="{escape(verify_url)}" style="display:inline-block;padding:10px 16px;background:#0f8a5f;color:#fff;
text-decoration:none;border-radius:8px;font-weight:600;">이메일 인증하기</a></p>
<p style="font-size:12px;color:#6b7280;word-break:break-all;">{escape(verify_url)}</p>
<p style="font-size:12px;color:#6b7280;">본인이 요청한 가입이 아니면 이 메일을 무시하시면 됩니다.</p>"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    from_addr = s.email_smtp_from or s.email_smtp_user or ""
    if from_addr:
        msg["From"] = from_addr
    else:
        msg["From"] = to_email
    msg["To"] = to_email
    msg.attach(MIMEText(plain, "plain", "utf-8"))
    msg.attach(MIMEText(html, "html", "utf-8"))

    if not s.email_smtp_host or not s.email_smtp_user:
        _log.warning(
            "EMAIL_SMTP_HOST / EMAIL_SMTP_USER 가 없어 메일을 보내지 않습니다. "
            "인증 URL(직접 열기): %s",
            verify_url,
        )
        return

    with smtplib.SMTP(s.email_smtp_host, s.email_smtp_port, timeout=30) as smtp:
        if s.email_smtp_use_tls:
            smtp.starttls()
        if s.email_smtp_password:
            smtp.login(s.email_smtp_user, s.email_smtp_password)
        out_from = s.email_smtp_from or s.email_smtp_user
        smtp.sendmail(out_from, [to_email], msg.as_string())
