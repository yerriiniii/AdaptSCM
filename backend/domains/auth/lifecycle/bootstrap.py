import logging

from domains.auth.services.user_account import ensure_gate_access_user, ensure_initial_admin_user
from shared.db import get_db_session, is_database_configured

_log = logging.getLogger(__name__)


def run_initial_admin_bootstrap() -> None:
    if not is_database_configured():
        return
    gen = get_db_session()
    db = next(gen)
    try:
        ensure_gate_access_user(db)
        ensure_initial_admin_user(db)
    except Exception:
        _log.exception("초기 관리자 계정 생성에 실패했습니다.")
    finally:
        db.close()
        try:
            next(gen)
        except StopIteration:
            pass

__all__ = ["run_initial_admin_bootstrap"]
