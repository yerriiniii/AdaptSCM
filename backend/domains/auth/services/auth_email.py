import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from html import escape

from shared.config import get_runtime_settings

_log = logging.getLogger(__name__)


def send_signup_email_proof_link(*, to_email: str, signup_url: str) -> bool:
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
    msg["From"] = from_addr if from_addr else to_email
    msg["To"] = to_email
    msg.attach(MIMEText(plain, "plain", "utf-8"))
    msg.attach(MIMEText(html, "html", "utf-8"))

    if not s.email_smtp_host or not s.email_smtp_user:
        _log.warning(
            "EMAIL_SMTP_HOST / EMAIL_SMTP_USER 가 없어 메일을 보내지 않습니다. 가입 인증 URL: %s",
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
    except Exception as exc:  # noqa: BLE001
        _log.error("SMTP 전송 실패: %s", exc, exc_info=True)
        return False
    return True


__all__ = ["send_signup_email_proof_link"]
