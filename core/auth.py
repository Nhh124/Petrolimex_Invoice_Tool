import hashlib
import hmac
import os

DEFAULT_USER = "boyhieu"
# Mã hóa bảo mật cao nhất: PBKDF2-HMAC-SHA256 với 600,000 vòng lặp + Cryptographic Salt
STORED_HASH = "5be088e3ca9f045a593471c6771107778d73bab676754e8996cf8efa8b2ffaae$7a02a5aa8974ff3ea435793db7f14d4ff58bd6843dbab6631b0f95b7281cb307"

def verify_login(username, password):
    """
    Xác thực tài khoản và mật khẩu bằng thuật toán PBKDF2 bảo mật cao
    """
    if not username or not password:
        return False
    
    if username.strip().lower() != DEFAULT_USER.lower():
        return False

    try:
        salt_hex, hash_hex = STORED_HASH.split('$')
        salt = bytes.fromhex(salt_hex)
        expected_hash = bytes.fromhex(hash_hex)
        
        computed_hash = hashlib.pbkdf2_hmac(
            'sha256',
            password.encode('utf-8'),
            salt,
            600000
        )
        return hmac.compare_digest(computed_hash, expected_hash)
    except Exception as e:
        print(f"Lỗi xác thực đăng nhập: {e}")
        return False
