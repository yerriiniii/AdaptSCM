# 허용 특수문자(아래 집합만 인정, 공백·한글·그 밖 기호 불가)
_PASSWORD_SPECIALS = set('!@#$%^&*()_+-={}[]:;"\'|,.<>?/~`')

# UI·에러 메시지에 그대로 쓰는 나열(검증 집합과 동일)
PASSWORD_SPECIALS_LIST_SPACED = " ".join(sorted(_PASSWORD_SPECIALS))
PASSWORD_SPECIALS_LIST_COMPACT = "".join(sorted(_PASSWORD_SPECIALS))


def validate_password_strength(plain: str) -> str | None:
    """적합하면 None, 아니면 사용자에게 보여줄 한국어 메시지."""
    if len(plain) < 8 or len(plain) > 16:
        return "비밀번호는 8~16자여야 합니다."
    t = str(plain)
    if t.strip() != t or " " in t or "\n" in t or "\r" in t or "\t" in t:
        return "비밀번호에 공백·탭·줄바꿈을 넣을 수 없습니다."
    if not any("A" <= c <= "Z" for c in plain):
        return "영문 대문자(A–Z)를 최소 1자 포함해야 합니다."
    if not any("a" <= c <= "z" for c in plain):
        return "영문 소문자(a–z)를 최소 1자 포함해야 합니다."
    if not any(c.isdigit() for c in plain):
        return "숫자(0–9)를 최소 1자 포함해야 합니다."
    if not any(c in _PASSWORD_SPECIALS for c in plain):
        return (
            "다음 특수문자만 사용할 수 있으며(이 중 최소 1자), 공백은 불가합니다. "
            f"{PASSWORD_SPECIALS_LIST_COMPACT}"
        )
    return None
