"""內容沒變就還原舊檔。結束代碼 0 = 沒變，1 = 有變。"""
import base64
import json
import os
import subprocess
import sys

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC


def decrypt(enc, pw):
    d = lambda s: base64.b64decode(s)
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=d(enc["salt"]), iterations=enc["iter"]).derive(pw.encode())
    return json.loads(AESGCM(key).decrypt(d(enc["iv"]), d(enc["ct"]), None))


path = sys.argv[1]
pw = os.environ.get("RECIPE_PASSWORD") or (sys.argv[2] if len(sys.argv) > 2 else "")
try:
    old_raw = subprocess.run(["git", "show", "HEAD:" + path], capture_output=True, check=True).stdout
    old = decrypt(json.loads(old_raw), pw)
except Exception:
    sys.exit(1)  # 沒有舊檔或舊檔密碼不同 → 視為有變動
new = decrypt(json.load(open(path)), pw)
old.pop("updated", None)
new.pop("updated", None)
if old == new:
    open(path, "wb").write(old_raw)
    print("內容沒變，不更新")
    sys.exit(0)
print("內容有變動")
sys.exit(1)
