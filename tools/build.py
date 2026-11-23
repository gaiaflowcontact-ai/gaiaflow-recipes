"""產生加密資料。python tools/build.py <資料夾> <密碼> [輸出檔]"""
import base64
import datetime
import glob
import json
import os
import re
import sys

import openpyxl
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

ITER = 200_000
STATIONS = ["冷台", "熱台", "甜點"]


def is_checked(v):
    if v is True:
        return True
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return v == 1
    if isinstance(v, str):
        return v.strip().upper() in ("TRUE", "V", "✓", "✔", "1", "是")
    return False


def read_selection(path):
    """回傳 [(ID, [(工作站, 分類), ...]), ...]"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["配料分類"]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    hdr_i = next(i for i, r in enumerate(rows) if r and str(r[0] or "").strip() == "ID")
    hdr = rows[hdr_i]
    grp = rows[hdr_i - 1] if hdr_i > 0 else [None] * len(hdr)
    # 合併儲存格只有第一格有值 → 往右延伸
    groups, cur = [], None
    for i in range(len(hdr)):
        g = grp[i] if i < len(grp) else None
        if g not in (None, ""):
            cur = str(g).strip()
        groups.append(cur)
    web_col = next(i for i, h in enumerate(hdr) if str(h or "").strip() == "上網站")
    cat_cols = [(i, groups[i], str(hdr[i]).strip()) for i in range(len(hdr))
                if groups[i] in STATIONS and hdr[i] not in (None, "")]
    out = []
    for r in rows[hdr_i + 1:]:
        if not r or not r[0]:
            continue
        pid = str(r[0]).strip().upper()
        if not re.fullmatch(r"PREP\d+", pid):
            continue
        if not is_checked(r[web_col] if web_col < len(r) else None):
            continue
        cats = [(g, s) for i, g, s in cat_cols if i < len(r) and is_checked(r[i])]
        out.append((pid, cats))
    return out


def num(v):
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except ValueError:
        return None


def read_prep(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    rows = [list(r) for r in wb["成本"].iter_rows(values_only=True)]
    hdr = next((r for r in rows if "編號 ID" in r), [])
    def after(lbl):
        if lbl in hdr:
            i = hdr.index(lbl)
            return hdr[i + 1] if i + 1 < len(hdr) else None
    def label_row(prefix):
        return next((r for r in rows if any(isinstance(x, str) and x.strip().startswith(prefix) for x in r)), None)
    start = next(i for i, r in enumerate(rows) if "Ingredients" in r) + 1
    end = next((i for i, r in enumerate(rows) if any(isinstance(x, str) and x.startswith("材料總重量") for x in r)), len(rows))
    items = []
    for r in rows[start:end]:
        iid, name, amt, unit = (r + [None] * 5)[1:5]
        g = num(amt)
        if not name and g is None:
            continue
        items.append({
            "id": str(iid).strip() if iid else "",
            "name": str(name).strip() if name else "",
            "g": g,
            "unit": str(unit).strip() if unit else "g",
        })
    def first_num(prefix):
        r = label_row(prefix)
        if not r:
            return None
        ns = [num(x) for x in r if num(x) is not None]
        return ns[0] if ns else None
    steps = ""
    if "食譜" in wb.sheetnames:
        parts = []
        for r in wb["食譜"].iter_rows(values_only=True):
            for v in r:
                if v not in (None, ""):
                    parts.append(str(v).strip())
        steps = "\n\n".join(parts)
    return {
        "id": str(after("編號 ID") or "").strip(),
        "name": str(after("名稱") or "").strip(),
        "items": items,
        "total": first_num("材料總重量"),
        "yield": first_num("成品總重量"),
        "steps": steps,
    }


def build(root):
    sel = read_selection(os.path.join(root, "配料分類表.xlsx"))
    recipes, missing = [], []
    for pid, cats in sel:
        fs = glob.glob(os.path.join(root, "配料", pid + "_*.xlsx"))
        if not fs:
            missing.append(pid)
            continue
        try:
            p = read_prep(fs[0])
        except Exception as e:  # 單一檔案壞掉不影響其他
            print("讀取失敗", fs[0], e, file=sys.stderr)
            missing.append(pid)
            continue
        stem = os.path.basename(fs[0]).rsplit(".", 1)[0]
        p["id"] = pid
        p["name"] = p["name"] or stem.split("_", 1)[-1]
        p["cats"] = [list(c) for c in cats]
        recipes.append(p)
    recipes.sort(key=lambda r: int(r["id"][4:]))
    tz = datetime.timezone(datetime.timedelta(hours=8))
    return {
        "updated": datetime.datetime.now(tz).strftime("%Y-%m-%d %H:%M"),
        "stations": {
            "冷台": ["香料粉", "配料", "醬汁", "風味油"],
            "熱台": ["大醬", "配料", "醬汁", "風味油"],
            "甜點": ["甜點"],
        },
        "recipes": recipes,
        "missing": missing,
    }


def encrypt(obj, password):
    salt, iv = os.urandom(16), os.urandom(12)
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(password.encode("utf-8"))
    ct = AESGCM(key).encrypt(iv, json.dumps(obj, ensure_ascii=False).encode("utf-8"), None)
    b64 = lambda b: base64.b64encode(b).decode()
    return {"v": 1, "iter": ITER, "salt": b64(salt), "iv": b64(iv), "ct": b64(ct)}


if __name__ == "__main__":
    root, password = sys.argv[1], sys.argv[2]
    out = sys.argv[3] if len(sys.argv) > 3 else os.path.join(os.path.dirname(__file__), "..", "data", "recipes.enc.json")
    data = build(root)
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with open(out, "w") as f:
        json.dump(encrypt(data, password), f)
    n_steps = sum(1 for r in data["recipes"] if r["steps"])
    print(f"{len(data['recipes'])} 個配料（{n_steps} 個有做法），缺檔：{data['missing'] or '無'}")
