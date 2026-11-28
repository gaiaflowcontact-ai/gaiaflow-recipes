"""下載同步需要的檔案。python tools/fetch_onedrive.py <連結> <輸出資料夾>"""
import base64
import io
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import zipfile

share, out = (sys.argv[1] if len(sys.argv) > 1 else "").strip(), sys.argv[2] if len(sys.argv) > 2 else "src"


def write_status(lines):
    with open("sync-status.txt", "w") as f:
        f.write(time.strftime("%Y-%m-%d %H:%M UTC", time.gmtime()) + "\n" + "\n".join(lines) + "\n")


shape = f"連結格式：長度 {len(share)}，" + ("1drv.ms 新版資料夾連結" if re.match(r"https://1drv\.ms/f/c/", share) else "開頭 " + share[:12].split("//")[-1][:6] + "…")
if not share:
    write_status(["ONEDRIVE_URL 是空的（Secrets 名稱要完全一樣：ONEDRIVE_URL）"])
    sys.exit(1)
UA = {"User-Agent": "Mozilla/5.0 (gaiaflow-recipes sync)"}


def get(url, tries=3):
    last = None
    for i in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=180) as r:
                return r.read(), r.headers.get("Content-Type", ""), r.geturl()
        except Exception as e:
            last = e
            time.sleep(2 * (i + 1))
    raise last


def save(name, data):
    p = os.path.join(out, name)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "wb") as f:
        f.write(data)


def from_zip(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    n_prep, got_table = 0, False
    for info in z.infolist():
        name = info.filename
        try:  # 有些 zip 的中文檔名用 cp437 存
            name = name.encode("cp437").decode("utf-8") if not (info.flag_bits & 0x800) else name
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
        parts = [p for p in name.replace("\\", "/").split("/") if p]
        if not parts or parts[-1].startswith("~$") or not parts[-1].lower().endswith(".xlsx"):
            continue
        if parts[-1] == "食材顯示名稱.xlsx" and len(parts) <= 2:
            save("食材顯示名稱.xlsx", z.read(info))
        elif parts[-1] == "配料分類表.xlsx" and len(parts) <= 2:
            save("配料分類表.xlsx", z.read(info))
            got_table = True
        elif len(parts) >= 2 and parts[-2] == "配料" and len(parts) <= 3:
            save(os.path.join("配料", parts[-1]), z.read(info))
            n_prep += 1
    if not got_table or n_prep == 0:
        raise RuntimeError(f"zip 裡找不到需要的檔案（配料分類表：{got_table}，配料檔：{n_prep}）")
    return n_prep


def via_zip(url):
    data, ctype, _ = get(url)
    if data[:2] != b"PK":
        raise RuntimeError(f"不是 zip（{ctype}，{len(data)} bytes）")
    return from_zip(data)


def via_api(token):
    base = f"https://api.onedrive.com/v1.0/shares/{token}/driveItem"
    def children(path=""):
        url = base + (":/" + urllib.parse.quote(path) + ":" if path else "") + "/children?$top=500"
        items = []
        while url:
            d = json.loads(get(url)[0])
            items += d.get("value", [])
            url = d.get("@odata.nextLink")
        return items
    def dl(it):
        link = it.get("@content.downloadUrl") or it.get("@microsoft.graph.downloadUrl")
        return get(link)[0]
    top = {it["name"]: it for it in children()}
    if "配料分類表.xlsx" not in top:
        raise RuntimeError("資料夾裡找不到 配料分類表.xlsx")
    save("配料分類表.xlsx", dl(top["配料分類表.xlsx"]))
    if "食材顯示名稱.xlsx" in top:
        save("食材顯示名稱.xlsx", dl(top["食材顯示名稱.xlsx"]))
    n = 0
    for it in children("配料"):
        if "file" in it and it["name"].lower().endswith(".xlsx") and not it["name"].startswith("~$"):
            save(os.path.join("配料", it["name"]), dl(it))
            n += 1
    return n


BADGER_APP = "5cbed6ac-a083-4e14-b191-b4ba07653de2"


def via_badger(share):
    req = urllib.request.Request("https://api-badgerp.svc.ms/v1.0/token",
                                 data=json.dumps({"appId": BADGER_APP}).encode(),
                                 headers={**UA, "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=60) as r:
        tk = json.loads(r.read())["token"]
    H = {**UA, "Authorization": "Badger " + tk, "Prefer": "autoredeem"}
    api = "https://my.microsoftpersonalcontent.com/_api/v2.0"
    def jget(url):
        with urllib.request.urlopen(urllib.request.Request(url, headers=H), timeout=120) as r:
            return json.loads(r.read())
    def children(drive, item):
        url = f"{api}/drives/{drive}/items/{item}/children?$top=500"
        items = []
        while url:
            d = jget(url)
            items += d.get("value", [])
            url = d.get("@odata.nextLink")
        return items
    sid = "u!" + base64.urlsafe_b64encode(share.encode()).decode().rstrip("=")
    root = jget(f"{api}/shares/{sid}/driveItem")
    drive = root["parentReference"]["driveId"]
    top = {it["name"]: it for it in children(drive, root["id"])}
    def dl(it):
        tries = [lambda: urllib.request.urlopen(urllib.request.Request(f"{api}/drives/{drive}/items/{it['id']}/content", headers=H), timeout=180).read()]
        if it.get("@content.downloadUrl"):
            tries.append(lambda: get(it["@content.downloadUrl"])[0])
        size = it.get("size")
        info = []
        for attempt in range(3):
            for k, t in enumerate(tries):
                try:
                    data = t()
                except Exception as e:
                    info.append(f"{k}:{type(e).__name__}")
                    continue
                if data[:2] == b"PK" and zipfile.is_zipfile(io.BytesIO(data)) and (not size or len(data) == size):
                    return data
                info.append(f"{k}:{len(data)}B/{size}")
            time.sleep(3)
        raise RuntimeError(it["name"] + " 下載失敗：" + "；".join(info))
    if "配料分類表.xlsx" not in top or "配料" not in top:
        raise RuntimeError("資料夾內容不符")
    save("配料分類表.xlsx", dl(top["配料分類表.xlsx"]))
    if "食材顯示名稱.xlsx" in top:
        save("食材顯示名稱.xlsx", dl(top["食材顯示名稱.xlsx"]))
    n = 0
    for it in children(drive, top["配料"]["id"]):
        if "file" in it and it["name"].lower().endswith(".xlsx") and not it["name"].startswith("~$"):
            save(os.path.join("配料", it["name"]), dl(it))
            n += 1
    return n


token = "u!" + base64.urlsafe_b64encode(share.encode()).decode().rstrip("=")
attempts = [("方法A", lambda: via_badger(share))]
m = re.search(r"1drv\.ms/\w/c/([0-9a-fA-F]+)/([\w-]+)", share)
if m:
    cid, tok = m.groups()
    attempts.append(("方法B",
                     lambda: via_zip(f"https://my.microsoftpersonalcontent.com/personal/{cid}/_layouts/15/download.aspx?share={tok}")))
attempts += [
    ("方法C", lambda: via_api(token)),
    ("方法D", lambda: via_zip(f"https://api.onedrive.com/v1.0/shares/{token}/root/content")),
    ("方法E", lambda: via_zip(share + ("&" if "?" in share else "?") + "download=1")),
]
def clean(msg):
    return re.sub(r"https?://\S+", "[網址]", msg)[:150]


errors, ok = [], None
for label, fn in attempts:
    try:
        n = fn()
        ok = f"{label} 成功，{n} 個檔"
        break
    except Exception as e:
        errors.append(f"{label} 失敗：{type(e).__name__} {clean(str(e))}")
lines = [shape] + errors + ([ok] if ok else ["全部失敗"])
if ok:
    t = os.path.join(out, "配料分類表.xlsx")
    lines.append(f"配料分類表：{os.path.getsize(t)} bytes，zip={zipfile.is_zipfile(t)}")
    bad = [f for f in sorted(os.listdir(os.path.join(out, "配料"))) if not zipfile.is_zipfile(os.path.join(out, "配料", f))]
    lines.append(f"壞掉的配料檔：{len(bad)} 個 {bad[:5]}")
print("\n".join(lines))
write_status(lines)
if not ok:
    sys.exit(1)
