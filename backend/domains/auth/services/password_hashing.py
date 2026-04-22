import bcrypt

# bcrypt는 입력을 72바이트로 제한. passlib과 동일하게 긴 UTF-8은 앞 72바이트만 사용.
def _to_bcrypt_secret(plain: str) -> bytes:
    b = plain.encode("utf-8")
    if len(b) > 72:
        return b[:72]
    return b


def hash_password(plain: str) -> str:
    secret = _to_bcrypt_secret(plain)
    return bcrypt.hashpw(secret, bcrypt.gensalt(rounds=12)).decode("ascii")


def verify_password(plain: str, hashed: str) -> bool:
    if not plain or not hashed:
        return False
    try:
        return bcrypt.checkpw(_to_bcrypt_secret(plain), hashed.encode("ascii"))
    except (ValueError, TypeError):
        return False
