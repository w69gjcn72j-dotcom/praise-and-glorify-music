#!/usr/bin/env python3
"""Rebuild the Praise & Glorify music app from the shared folders.

Reads
  Music/Praise and Glorify/Song Sets.xlsx          (Sundays, Library, Songs without lead sheets)
  Music/Music files/English Choruses/*-lead-*.pdf  (SongSelect lead sheets)
  Music/Music files/Chinese Choruses/*-lead-*.pdf
Writes (inside the site folder)
  s/<id>.bin      lead sheets, AES-GCM encrypted with the team password
  s/index.json    source hashes, so unchanged sheets are not re-encrypted
  data.js         everything the page shows
It also adds any new lead sheet to the Library sheet so a Chinese title can be filled in.

The team password lives in .pg-key next to this script's site folder (never committed).
Run:  python3 tools/build.py      (from the site folder)
"""
import hashlib, json, os, re, secrets, subprocess, sys, datetime, unicodedata

SITE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PG_DIR = os.path.dirname(SITE)                       # Music/Praise and Glorify
MUSIC = os.path.dirname(PG_DIR)                      # Music
XLSX = os.path.join(PG_DIR, "Song Sets.xlsx")
FOLDERS = [("English Choruses", "en"), ("Chinese Choruses", "zh")]
OUT = os.path.join(SITE, "s")

from openpyxl import load_workbook
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes


def slug(base):
    return re.sub(r"[^a-z0-9]+", "-", re.sub(r"-lead-.*", "", base).lower()).strip("-")


def norm(t):
    t = unicodedata.normalize("NFKC", str(t or "")).lower()
    t = re.sub(r"\(.*?\)|（.*?）", "", t)
    t = t.replace("&", "and").replace("’", "'")
    return re.sub(r"[^0-9a-z一-鿿]+", "", t)


def cjk(t):
    return bool(re.search(r"[一-鿿]", str(t or "")))


def pdf_meta(path):
    ccli, pages = None, None
    try:
        txt = subprocess.run(["pdftotext", path, "-"], capture_output=True, text=True, timeout=60).stdout
        m = re.search(r"CCLI Song\s*#\s*(\d+)", txt)
        ccli = int(m.group(1)) if m else None
        info = subprocess.run(["pdfinfo", path], capture_output=True, text=True, timeout=60).stdout
        m = re.search(r"Pages:\s+(\d+)", info)
        pages = int(m.group(1)) if m else None
    except Exception:
        pass
    return ccli, pages


def main():
    keyfile = os.path.join(SITE, ".pg-key")
    K = json.load(open(keyfile))
    salt = bytes.fromhex(K["salt"])
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=K["iter"]).derive(K["password"].encode())
    aes = AESGCM(key)

    def enc(b):
        iv = secrets.token_bytes(12)
        return iv + aes.encrypt(iv, b, None)

    wb = load_workbook(XLSX)
    lib_ws, sun_ws, oth_ws = wb["Library"], wb["Sundays"], wb["Songs without lead sheets"]

    # ---- Library overrides from the sheet
    lib_rows = {}
    for r in lib_ws.iter_rows(min_row=2, values_only=True):
        if r and r[0]:
            lib_rows[str(r[0]).strip()] = {"title": (r[2] or "").strip() if r[2] else "", "zh": (r[3] or "").strip() if r[3] else ""}

    # ---- Scan lead sheets
    os.makedirs(OUT, exist_ok=True)
    manifest_path = os.path.join(OUT, "index.json")
    manifest = json.load(open(manifest_path)) if os.path.exists(manifest_path) else {}
    songs, new_files, skipped, keep = [], [], [], set()
    for folder, libcode in FOLDERS:
        path = os.path.join(MUSIC, "Music files", folder)
        if not os.path.isdir(path):
            continue
        for f in sorted(os.listdir(path)):
            if "-lead-" not in f or not f.lower().endswith(".pdf") or f.startswith("~$"):
                continue
            full = os.path.join(path, f)
            data = open(full, "rb").read()
            if not data.startswith(b"%PDF"):
                skipped.append(f + " (not a real PDF — download it again)")
                continue
            base = f[:-4]
            sid = slug(base)
            if sid in keep:
                skipped.append(f + " (same song name as another file)")
                continue
            keep.add(sid)
            m = re.search(r"-lead-([A-G][b#]?m?)", base)
            skey = m.group(1) if m else ""
            ov = lib_rows.get(f)
            if ov is None:
                new_files.append((f, folder))
                ov = {"title": "", "zh": ""}
            h = hashlib.sha256(data).hexdigest()
            meta = manifest.get(sid, {})
            if meta.get("sha") != h or not os.path.exists(os.path.join(OUT, sid + ".bin")):
                open(os.path.join(OUT, sid + ".bin"), "wb").write(enc(data))
                ccli, pages = pdf_meta(full)
                meta = {"sha": h, "ccli": ccli, "pages": pages}
                manifest[sid] = meta
            title = ov["title"] or re.sub(r"-lead-.*", "", base)
            songs.append({"id": sid, "en": title, "zh": ov["zh"], "key": skey, "ccli": meta.get("ccli"),
                          "pages": meta.get("pages") or 1, "lib": libcode, "sheet": True, "file": f})
    for sid in list(manifest):
        if sid not in keep:
            manifest.pop(sid)
            try:
                os.remove(os.path.join(OUT, sid + ".bin"))
            except OSError:
                pass
    json.dump(manifest, open(manifest_path, "w"), indent=1, sort_keys=True)

    # ---- Songs without lead sheets
    others = []
    for r in oth_ws.iter_rows(min_row=2, values_only=True):
        if r and (r[0] or r[1]):
            others.append({"id": None, "en": (r[0] or "").strip(), "zh": (r[1] or "").strip(),
                           "lib": (r[2] or "zh").strip(), "sheet": False, "note": (r[3] or "").strip() if r[3] else ""})

    # ---- Title matching
    index = {}
    for s in songs:
        for t in (s["en"], s["zh"], re.sub(r"-lead-.*", "", s["file"][:-4])):
            if norm(t):
                index.setdefault(norm(t), s["id"])
    oindex = {}
    for o in others:
        for t in (o["en"], o["zh"]):
            if norm(t):
                oindex.setdefault(norm(t), o)
    unmatched = set()

    def resolve(t):
        t = str(t).strip()
        n = norm(t)
        if n in index:
            return index[n]
        # loose match: a library title that starts with what was typed, or vice versa
        cands = [sid for k, sid in index.items() if len(n) >= 6 and (k.startswith(n) or n.startswith(k))]
        if len(set(cands)) == 1:
            return cands[0]
        unmatched.add(t)
        o = oindex.get(n)
        if o:
            return {"en": o["en"], "zh": o["zh"]}
        return {"en": "", "zh": t} if cjk(t) else {"en": t, "zh": ""}

    schedule = []
    for r in sun_ws.iter_rows(min_row=2, values_only=True):
        if not r or not r[0]:
            continue
        d = r[0]
        if isinstance(d, datetime.datetime):
            d = d.date()
        if not isinstance(d, datetime.date):
            try:
                d = datetime.datetime.strptime(str(d).strip(), "%d/%m/%Y").date()
            except ValueError:
                skipped.append(f"Sundays row with date '{r[0]}' (type the date as 4/10/2026)")
                continue
        schedule.append({
            "d": d.isoformat(), "type": (r[1] or "").strip() if r[1] else "",
            "passage": (r[2] or "").strip() if r[2] else "",
            "mEn": (r[3] or "").strip() if r[3] else "",
            "en": [resolve(x) for x in r[4:8] if x and str(x).strip()],
            "mCn": (r[8] or "").strip() if r[8] else "",
            "cn": [resolve(x) for x in r[9:13] if x and str(x).strip()],
            "notes": (r[13] or "").strip() if len(r) > 13 and r[13] else "",
        })
    schedule.sort(key=lambda w: w["d"])

    # ---- Add new lead sheets to the Library sheet
    if new_files:
        for f, folder in new_files:
            lib_ws.append([f, folder, re.sub(r"-lead-.*", "", f[:-4]), ""])
        try:
            wb.save(XLSX)
        except PermissionError:
            skipped.append("Song Sets.xlsx is open — new lead sheets were not added to the Library sheet this time")

    if not K.get("check"):
        K["check"] = enc(b"praise-and-glorify").hex()
        json.dump(K, open(keyfile, "w"))
    check = K["check"]
    for s in songs:
        s.pop("file", None)
    data = {"CRYPTO": {"salt": K["salt"], "iter": K["iter"], "check": check},
            "SONGS": songs, "SCHEDULE": schedule, "NOSHEET": others,
            "updated": datetime.date.today().isoformat()}
    js = "/* Generated by tools/build.py from Song Sets.xlsx and the Choruses folders. Do not edit by hand. */\n"
    js += "window.PG = " + json.dumps(data, ensure_ascii=False, indent=1) + ";\n"
    open(os.path.join(SITE, "data.js"), "w", encoding="utf-8").write(js)

    print(f"Lead sheets: {len(songs)}  Sundays: {len(schedule)}  New lead sheets: {len(new_files)}")
    for f, _ in new_files:
        print("  + new:", f)
    for t in sorted(unmatched):
        print("  ? no lead sheet for:", t)
    for s in skipped:
        print("  ! skipped:", s)


if __name__ == "__main__":
    main()
