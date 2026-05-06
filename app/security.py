import hashlib
import hmac
import os

def hash_password(password: str) -> str:
    salt = os.urandom(16)
    pwd_hash = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 120000)
    return f"{salt.hex()}:{pwd_hash.hex()}"

def verify_password(password: str, stored_value: str) -> bool:
    try:
        salt_hex, hash_hex = stored_value.split(":")
        salt = bytes.fromhex(salt_hex)
        expected = bytes.fromhex(hash_hex)
        actual = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 120000)
        return hmac.compare_digest(actual, expected)
    except Exception:
        return False

def validate_new_password(password: str):
    if len(password) < 6:
        return False, "La clave debe tener mínimo 6 caracteres."
    return True, ""
