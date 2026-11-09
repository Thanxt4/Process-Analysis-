"""
Process Control – ฟอร์มตรวจพื้นที่หน้างาน
- กรอกข้อมูลรายพื้นที่ (ข้อความ + รูป)
- Checklist ว่าพื้นที่ไหนตรวจแล้ว / ยังไม่ตรวจ
- Generate สไลด์ (.pptx) จากข้อมูลและรูป
- ตัดรอบทุกวันเวลา 07:00 น. (รอบ = 07:00 วันนี้ ถึง 06:59 วันถัดไป) และสร้างสไลด์เก็บไว้อัตโนมัติ
"""
import hashlib
import re
import secrets
import io
import json
import mimetypes
import os
import sqlite3
import threading
import time as _time
import shutil
import uuid
from contextlib import contextmanager
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import List, Optional
from zoneinfo import ZoneInfo

from fastapi import FastAPI, File, Form, HTTPException, Query, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.middleware.gzip import GZipMiddleware
from fastapi.staticfiles import StaticFiles
from starlette.concurrency import run_in_threadpool
from PIL import Image, ImageOps

import slides as _slides_mod
from slides import SLIDES_VERSION, build_pptx, build_pptx_multi

# เวลาแก้ไขล่าสุดของโค้ดสร้างสไลด์: ไฟล์สไลด์ที่เก็บไว้ก่อนหน้านี้ถือว่าเก่า → สร้างใหม่อัตโนมัติเมื่อมีคนกดดาวน์โหลด
SLIDES_CODE_MTIME = Path(_slides_mod.__file__).stat().st_mtime

# ---------------------------------------------------------------- settings
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("DATA_DIR", BASE_DIR.parent / "data"))
UPLOAD_DIR = DATA_DIR / "uploads"
SLIDE_DIR = DATA_DIR / "slides"
DB_PATH = DATA_DIR / "app.db"
AREAS_PATH = Path(os.environ.get("AREAS_FILE", BASE_DIR / "areas.json"))
TZ = ZoneInfo(os.environ.get("TZ_NAME", "Asia/Bangkok"))
_h, _m = os.environ.get("CUTOFF", "07:00").split(":")
CUTOFF = time(int(_h), int(_m))
ADMIN_PIN = os.environ.get("ADMIN_PIN", "").strip()  # ถ้าตั้งไว้ ต้องใส่ PIN ก่อนลบรายการ
EDIT_PIN = os.environ.get("EDIT_PIN", "2780").strip()
SLIDE_PIN = os.environ.get("SLIDE_PIN", "0000").strip()  # รหัสก่อน Generate / ดาวน์โหลดสไลด์  # รหัสสำหรับแก้ไขหัวข้อ Checklist ผ่านหน้าเว็บ
MAX_IMAGES = int(os.environ.get("MAX_IMAGES", "8"))      # รูปสูงสุดต่อ 1 สไลด์
MAX_PAGES = int(os.environ.get("MAX_PAGES", "15"))       # สไลด์สูงสุดต่อ 1 การตรวจ
MAX_VIDEO_SEC = float(os.environ.get("MAX_VIDEO_SEC", "10"))   # ความยาววิดีโอสูงสุด (วินาที)
MAX_VIDEO_MB = float(os.environ.get("MAX_VIDEO_MB", "80"))     # ขนาดไฟล์วิดีโอสูงสุด
VIDEO_EXT = {".mp4", ".mov", ".m4v", ".webm", ".3gp"}
BACKDATE_DAYS = int(os.environ.get("BACKDATE_DAYS", "7"))      # ส่ง/แก้ไขย้อนหลังได้กี่วัน
ACCESS_CODE = os.environ.get("ACCESS_CODE", "").strip()  # ว่าง = ไม่ต้องใส่รหัสเข้าเว็บ          # รหัสเข้าเว็บ (แนะนำให้ตั้งเมื่อเปิดเป็นสาธารณะ)
ACCESS_TOKEN = hashlib.sha256(("pc:" + ACCESS_CODE).encode()).hexdigest() if ACCESS_CODE else ""
MAX_SIDE = int(os.environ.get("IMAGE_MAX_SIDE", "1600"))
THUMB_SIDE = int(os.environ.get("THUMB_SIDE", "360"))           # ขนาดรูปย่อที่ใช้แสดงในรายการ
ROUNDS_LIMIT = int(os.environ.get("ROUNDS_LIMIT", "60"))         # แท็บสไลด์แสดงย้อนหลังกี่วัน
ARCHIVE_DIR = Path(os.environ.get("ARCHIVE_DIR", "").strip() or (DATA_DIR / "archive_uploads"))

mimetypes.add_type("application/manifest+json", ".webmanifest")
mimetypes.add_type("text/javascript", ".js")
mimetypes.add_type("font/woff", ".woff")
mimetypes.add_type("font/woff2", ".woff2")

for d in (UPLOAD_DIR, SLIDE_DIR):
    d.mkdir(parents=True, exist_ok=True)


CUSTOMERS = ("Weruva", "Smucker")  # ลูกค้าที่เลือกให้แต่ละ Checklist ได้
BOTH_CUST = "both"  # Checklist ที่ใช้กับทั้งสองลูกค้า (ไม่ระบุ = ไม่มีลูกค้า)
CUST_FILTERS = CUSTOMERS + ("any", "none")  # ตัวเลือกลูกค้าใน Dashboard: any = ลูกค้าทั้งหมด · none = ไม่มีลูกค้า
NO_ACT = "ไม่มีกิจกรรม"  # ผลการตรวจแบบที่ 3: ไปตรวจแล้วแต่จุดนั้นไม่มีกิจกรรม → ไม่นับในกราฟ/การวิเคราะห์
STATUSES = ("ปกติ", "พบปัญหา", NO_ACT)


def load_areas():
    with open(AREAS_PATH, encoding="utf-8") as f:
        cfg = json.load(f)
    item_map = {}
    for z in cfg["zones"]:
        for it in z["items"]:
            item_map[it["id"]] = {**it, "zone_id": z["id"], "zone": z["name"], "owner": z["owner"]}
    return cfg, item_map


AREAS, ITEM_MAP = load_areas()


# ป้ายบอกว่ารายการตรวจอยู่ในกระบวนการไหน (เลือกได้ในหน้าแก้ไขหัวข้อ) — แก้รายการได้ที่ "process_tags" ใน areas.json
DEFAULT_PROCESS_TAGS = ["Rawmat", "จุดเตรียม", "บรรจุ", "Retort&Store", "ห้องเคมี", "ห้องอิ้ง", "Warehouse"]


def process_tags():
    t = AREAS.get("process_tags")
    return t if isinstance(t, list) and t else DEFAULT_PROCESS_TAGS


def visible_zones():
    """พื้นที่/รายการที่ใช้งานอยู่ (ไม่รวมที่ถูกลบ = hidden) — ของที่ลบยังเก็บไว้ให้ข้อมูลย้อนหลังอ่านชื่อได้"""
    out = []
    for z in AREAS["zones"]:
        if z.get("hidden"):
            continue
        out.append({**z, "deck": zone_deck(z), "items": [it for it in z["items"] if not it.get("hidden")]})
    return out


def item_active(it, rd: str, checked: bool) -> bool:
    """รายการนี้นับอยู่ใน Checklist ของรอบ rd หรือไม่ (ใช้ % ย้อนหลังตามรายการที่มีจริงในวันนั้น)
    - added   = รอบที่เพิ่มรายการ: ก่อนหน้านั้นไม่นับ / รอบที่เพิ่ม นับเฉพาะถ้ามีคนตรวจ
    - removed = รอบที่ลบรายการ: หลังจากนั้นไม่นับ / รอบที่ลบ นับเฉพาะถ้ามีคนตรวจ"""
    a, r = it.get("added"), it.get("removed")
    if it.get("hidden") is not True and a and rd == a and rd >= round_of(now()).isoformat():
        return True  # รอบปัจจุบัน: รายการที่เพิ่งเพิ่มต้องแสดงในฟอร์มทันที
    if a and (rd < a or (rd == a and not checked)):
        return False
    if r and (rd > r or (rd == r and not checked)):
        return False
    if it.get("hidden") and not r:
        return checked
    return True


def zones_for_round(rd: str, checked_ids: set):
    """โครงสร้าง Checklist ตามที่เป็นในรอบ rd"""
    out = []
    for z in AREAS["zones"]:
        items = [it for it in z["items"] if item_active(it, rd, it["id"] in checked_ids)]
        if not items or (z.get("hidden") and not any(it["id"] in checked_ids for it in items)):
            continue
        out.append({**z, "items": items})
    return out


GMP_DEFAULT_AREAS = ["PF1", "PF2", "CS", "โรงเนื้อ", "WH"]


def _migrate_gmp_areas():
    """ครั้งเดียว: เปลี่ยน GMP จากพื้นที่เดียว เป็นกลุ่มที่มีหลายพื้นที่ (PF1, PF2, CS, โรงเนื้อ, WH) แต่ละพื้นที่มี checklist ของตัวเอง
    พื้นที่ GMP เดิมถูกซ่อนไว้ (ข้อมูลย้อนหลังยังอยู่ครบ)"""
    global AREAS, ITEM_MAP
    if AREAS.get("gmp_groups"):
        return
    cur = round_of(now()).isoformat()
    zones = AREAS["zones"]
    old = [z for z in zones if not z.get("hidden") and zone_deck(z) == "gmp"]
    if any(sum(1 for i in z["items"] if not i.get("hidden") and not i.get("extra")) > 1 for z in old):
        AREAS["gmp_groups"] = True  # มีพื้นที่ GMP แบบหลายรายการอยู่แล้ว ไม่ต้องย้าย
    else:
        used = {z["id"] for z in zones} | {i["id"] for z in zones for i in z["items"]}

        def nid(prefix):
            while True:
                i = prefix + uuid.uuid4().hex[:6]
                if i not in used:
                    used.add(i)
                    return i
        first = next((i["name"] for z in old for i in z["items"] if not i.get("extra")),
                     "GMP - ตรวจทั่วไป (ความสะอาด, พฤติกรรมพนักงาน, สุขลักษณะ)")
        owner = old[0].get("owner", "") if old else ""
        for z in old:
            z["hidden"] = True
            z["items"] = [{**i, "removed": i.get("removed") or cur} if not i.get("hidden") else i for i in z["items"]]
        for name in GMP_DEFAULT_AREAS:
            zones.append({"id": nid("Z"), "name": name, "owner": owner, "deck": "gmp", "items": [
                {"id": nid("A"), "no": 1, "name": first, "added": cur},
                {"id": nid("X"), "no": "+", "name": "หัวข้อเพิ่มเติม (นอกเหนือ Checklist)", "extra": True, "free": True,
                 "added": cur}]})
        AREAS["gmp_groups"] = True
    _save_areas(AREAS, "before_gmp_groups")
    AREAS, ITEM_MAP = load_areas()
    print("[areas] GMP แยกเป็นหลายพื้นที่แล้ว", flush=True)


WH_POINTS = ["WH1", "WH5", "WH6", "WH7", "WH15", "WH16"]


def _migrate_wh_points():
    """ครั้งเดียว: แยกพื้นที่ GMP "WH" เป็นจุดย่อย WH1, WH5, WH6, WH7, WH15, WH16 (กลุ่มย่อย "WH")
    แต่ละจุดได้รายการตรวจชุดเดียวกับ WH เดิม · พื้นที่ WH เดิมถูกซ่อน (ข้อมูลย้อนหลังยังอยู่ครบ)"""
    global AREAS, ITEM_MAP
    if AREAS.get("wh_points"):
        return
    zones = AREAS["zones"]
    old = next((z for z in zones if not z.get("hidden") and zone_deck(z) == "gmp"
                and str(z.get("name", "")).strip().upper() == "WH"), None)
    AREAS["wh_points"] = True
    if old is not None and not any(z.get("sub") == "WH" for z in zones):
        cur = round_of(now()).isoformat()
        used = {z["id"] for z in zones} | {i["id"] for z in zones for i in z["items"]}

        def nid(prefix):
            while True:
                i = prefix + uuid.uuid4().hex[:6]
                if i not in used:
                    used.add(i)
                    return i
        base = [i for i in old["items"] if not i.get("hidden") and not i.get("extra")]
        new_zones = []
        for name in WH_POINTS:
            items = [{k: v for k, v in i.items() if k not in ("id", "added", "removed", "hidden")} | {"id": nid("A"), "added": cur}
                     for i in base]
            items.append({"id": nid("X"), "no": "+", "name": "หัวข้อเพิ่มเติม (นอกเหนือ Checklist)", "extra": True, "free": True,
                          "added": cur})
            new_zones.append({"id": nid("Z"), "name": name, "owner": old.get("owner", ""), "deck": "gmp", "sub": "WH",
                              "items": items})
        old["hidden"] = True
        old["items"] = [{**i, "removed": i.get("removed") or cur} if not i.get("hidden") else i for i in old["items"]]
        pos = zones.index(old)
        zones[pos + 1:pos + 1] = new_zones
        print("[areas] แยก GMP WH เป็น " + ", ".join(WH_POINTS), flush=True)
    _save_areas(AREAS, "wh_points")
    AREAS, ITEM_MAP = load_areas()


DELAY_ROOMS = ["Chill3", "Chill4", "Chill5", "Chill2", "Chill6", "Ante", "CSR3"]


def _migrate_delay_qty():
    """ครั้งเดียว: หัวข้อ "วัตถุดิบ Delay ห้องเย็น ..." ให้มีช่องกรอกจำนวนคันแยกตามห้อง (ตั้งชื่อห้องได้ในหน้าแก้ไขหัวข้อ)"""
    global AREAS, ITEM_MAP
    if AREAS.get("delay_qty") and AREAS.get("delay_rooms"):
        return
    AREAS["delay_qty"] = AREAS["delay_rooms"] = True
    n = 0
    for z in AREAS["zones"]:
        for it in z["items"]:
            nm = str(it.get("name", "")).lower()
            if it.get("extra") or "delay" not in nm or "ห้องเย็น" not in nm:
                continue
            q = it.get("qty") or {"unit": "คัน", "label": "วัตถุดิบ Delay", "rooms": []}
            if not q.get("rooms"):  # ใส่ห้องเริ่มต้น (แก้ไขได้ในหน้าแก้ไขหัวข้อ)
                q["rooms"] = list(DELAY_ROOMS)
            it["qty"] = q
            n += 1
    print(f"[areas] ช่องนับจำนวนคัน (Delay) {n} หัวข้อ · ห้อง: " + ", ".join(DELAY_ROOMS), flush=True)
    _save_areas(AREAS, "delay_qty")
    AREAS, ITEM_MAP = load_areas()


# ลูกค้าของแต่ละหัวข้อ (รายการจากหน้างาน 6 พ.ย. 2569) · หัวข้ออื่นใน Process Control / GMP = ไม่มีลูกค้า
# (deck, ชื่อพื้นที่, เลขที่, ชื่อหัวข้อขึ้นต้นด้วย, ลูกค้า)  both = Weruva, Smucker
CUST_SEED = [
    ("pc", "PF2", 5, "สุ่มตรวจสอบการปั่นผสมสารละลาย", "Weruva"),
    ("pc", "PF1 - จุดเตรียม", 37, "ลานเขียว จุด Tomra", "both"),
    ("pc", "PF1 - จุดเตรียม", 66, "จุดล้างถาด", "both"),
    ("pc", "PF1 - บรรจุ", 40, "สุ่มตรวจสอบการปั่นผสมสารละลาย", "both"),
    ("pc", "PF1 - บรรจุ", 47, "การชั่งน้ำหนัก/เคาะ/กวาดของบรรจุ", "both"),
    ("pc", "PF1 - บรรจุ", 48, "การคลุกผสม", "both"),
    ("pc", "PF1 - บรรจุ", 55, "ห้องเคมี PF1", "both"),
    ("pc", "PF1 - บรรจุ", 73, "รายงาน Delay Time", "both"),
    ("pc", "ห้องเย็น PF1", 74, "สุ่มดูเปลือกฟักทอง", "Weruva"),
    ("pc", "ห้องเย็น PF1", 82, "จุดเตรียมผัก", "Weruva"),
    ("pc", "ห้องเย็น PF1", 84, "ตรวจสอบวัตถุดิบห้อง CSR3", "Weruva"),
    ("gmp", "PF1", 2, "พื้นที่เตรียมการผลิต - ตรวจสอบความสะอาด", "both"),
    ("gmp", "PF1", 3, "พื้นที่เตรียมการผลิต - พนักงานที่ปฏิบัติงาน", "both"),
    ("gmp", "PF1", 6, "ไลน์บรรจุ - พนักงานที่ปฏิบัติงาน", "both"),
    ("gmp", "PF1", 8, "ไลน์บรรจุ - ตรวจสอบความสะอาด", "both"),
    ("gmp", "PF1", 9, "ไลน์บรรจุ - เศษตกพื้น", "Weruva"),
    ("gmp", "PF1", 12, "ห้องเคมี - ตรวจสอบความสะอาด", "both"),
    ("gmp", "PF1", 14, "ห้องเคมี - การตัดกระสอบ", "both"),
    ("gmp", "PF2", 2, "พื้นที่เตรียมการผลิต - ตรวจสอบความสะอาด", "Smucker"),
    ("gmp", "PF2", 8, "ไลน์บรรจุ - ตรวจสอบความสะอาด", "Weruva"),
    ("gmp", "PF2", 9, "ไลน์บรรจุ - เศษตกพื้น", "Weruva"),
]


def _migrate_cust_list():
    """ครั้งเดียว: ตั้งลูกค้าของหัวข้อ Process Control / GMP ตาม CUST_SEED · หัวข้ออื่น = ไม่มีลูกค้า
    จับคู่ด้วย ชื่อพื้นที่ + เลขที่ และตรวจชื่อหัวข้อ (ถ้าเลขที่ไม่ตรงชื่อ หาจากชื่อในพื้นที่เดียวกัน) · หลังจากนี้แก้ได้ที่หน้าแก้ไขหัวข้อ"""
    global AREAS, ITEM_MAP
    if AREAS.get("cust_seed_v1"):
        return
    norm = lambda t: " ".join(str(t or "").split())
    live = [z for z in AREAS["zones"] if not z.get("hidden") and (z.get("deck") or zone_deck(z)) in ("pc", "gmp")]
    want, missing = {}, []
    for deck, zname, no, prefix, cust in CUST_SEED:
        zs = [z for z in live if (z.get("deck") or zone_deck(z)) == deck and norm(z.get("name")) == zname]
        items = [it for z in zs for it in z["items"] if not it.get("extra") and not it.get("hidden")]
        one = lambda c: c[0] if len(c) == 1 else None
        hit = next((it for it in items if str(it.get("no")) == str(no) and norm(it.get("name")).startswith(prefix)), None) \
            or one([it for it in items if norm(it.get("name")).startswith(prefix)]) \
            or one([it for it in items if str(it.get("no")) == str(no)])
        if hit:
            want[id(hit)] = cust
        else:
            missing.append(f"{deck.upper()} {zname} #{no}")
    n = 0
    for z in live:
        for it in z["items"]:
            if it.get("extra"):
                continue
            c = want.get(id(it))
            if c:
                it["cust"] = c
                n += 1
            else:
                it.pop("cust", None)
    AREAS["cust_seed_v1"] = True
    print(f"[areas] ตั้งลูกค้าตามรายการ {n}/{len(CUST_SEED)} หัวข้อ · ที่เหลือ = ไม่มีลูกค้า"
          + (f" · หาไม่พบ: {', '.join(missing)}" if missing else ""), flush=True)
    _save_areas(AREAS, "cust_seed")
    AREAS, ITEM_MAP = load_areas()


# หัวข้อ Checklist ใหม่ (★ ทดลองใช้ · ยังไม่นับใน % ของ Dashboard = pilot) — เพิ่มครั้งเดียวตอนเปิดเซิร์ฟเวอร์
# (deck, ชื่อพื้นที่, เลขที่, ชื่อหัวข้อ, ลูกค้า)
NEW_ITEMS_V1 = [
    ("pc", "PF2", 97, "ตรวจสภาพอุปกรณ์สัมผัสอาหาร จุด Block cutter / Steam tunnel (พลั่ว ถาด ปะเก็น) และนับเครื่องมือหลังซ่อม", "Smucker"),
    ("pc", "PF2", 98, "ทวนสอบบันทึก: ค่าที่บันทึกตรงกับเครื่อง (steamer), หน่วยตาชั่ง, อุณหภูมิ gravy / chunk ครบทุกจุด", "both"),
    ("pc", "PF1 - จุดเตรียม", 96, "สภาพอุปกรณ์ไลน์ red meat: guide Metal detector, safety holder เครื่อง flaker, ปะเก็นฝา mixer", "Weruva"),
    ("pc", "PF1 - บรรจุ", 95, "ตรวจรอยเชื่อม/สภาพอุปกรณ์สัมผัสอาหาร (กรวยบรรจุ ตะแกรงเครื่องผสม gravy) ก่อน-หลังผลิต และนับเครื่องมือหลังซ่อม", "Smucker"),
    ("gmp", "PF1", 16, "พื้นที่เตรียมการผลิต - วัสดุลอก/หลุดเหนือโซนผลิตภัณฑ์ (สติกเกอร์ tag เทป ซิลิโคน สนิม สีลอก)", "both"),
    ("gmp", "PF1", 17, "ไลน์บรรจุ - วัสดุลอก/หลุดเหนือโซนผลิตภัณฑ์ (สติกเกอร์ tag เทป ซิลิโคน สนิม สีลอก)", "both"),
    ("gmp", "PF1", 18, "ห้องเคมี - วัสดุลอก/หลุด และสภาพถุง/ภาชนะเคมี (ไม่ขาด ไม่รั่ว)", "both"),
    ("gmp", "PF1", 19, "น้ำ condensate / น้ำหยดเหนือผลิตภัณฑ์ (ฝ้า ท่อแอร์ fan coil) และรถเข็น RM มีหลังคาคลุม", "both"),
    ("gmp", "PF1", 20, "ผนัง / ฝ้า / ประตู ไม่มีคราบดำหรือรา ไม่ชำรุด", "both"),
    ("gmp", "PF1", 21, "แมลง/สัตว์พาหะในจุดเสี่ยง และที่ดักแมลงมีบันทึกการตรวจ", "both"),
    ("gmp", "PF2", 15, "พื้นที่เตรียมการผลิต - วัสดุลอก/หลุดเหนือโซนผลิตภัณฑ์ (สติกเกอร์ tag เทป ซิลิโคน สนิม สีลอก)", "Smucker"),
    ("gmp", "PF2", 16, "ไลน์บรรจุ - วัสดุลอก/หลุดเหนือโซนผลิตภัณฑ์ (สติกเกอร์ tag เทป ซิลิโคน สนิม สีลอก)", "Weruva"),
    ("gmp", "PF2", 17, "น้ำ condensate / น้ำหยดเหนือผลิตภัณฑ์ (ฝ้า ท่อ hopper)", "Smucker"),
    ("gmp", "PF2", 18, "ผนัง / ฝ้า / ประตู ไม่มีคราบดำหรือรา ไม่ชำรุด", "Smucker"),
]


def _migrate_new_items_v1():
    """ครั้งเดียว: เพิ่มหัวข้อ Checklist ใหม่ตาม NEW_ITEMS_V1 (ไม่เพิ่มซ้ำถ้ามีหัวข้อชื่อเดียวกันในพื้นที่นั้นแล้ว)
    หัวข้อใหม่ติดป้าย pilot = แสดงใน Dashboard แต่ยังไม่นับใน % (เปิดให้นับได้ที่หน้าแก้ไขหัวข้อ)"""
    global AREAS, ITEM_MAP
    if AREAS.get("new_items_v1"):
        return
    norm = lambda t: " ".join(str(t or "").split())
    used = {z["id"] for z in AREAS["zones"]} | {it["id"] for z in AREAS["zones"] for it in z["items"]}
    rd = round_of(now()).isoformat()
    added, missing = 0, []
    for deck, zname, no, name, cust in NEW_ITEMS_V1:
        z = next((z for z in AREAS["zones"] if not z.get("hidden") and (z.get("deck") or zone_deck(z)) == deck and norm(z.get("name")) == zname), None)
        if not z:
            missing.append(f"{deck.upper()} {zname}")
            continue
        if any(not it.get("hidden") and norm(it.get("name")) == norm(name) for it in z["items"]):
            continue  # มีแล้ว
        iid = "A" + uuid.uuid4().hex[:6]
        while iid in used:
            iid = "A" + uuid.uuid4().hex[:6]
        used.add(iid)
        it = {"id": iid, "no": no, "name": name, "added": rd, "pilot": True}
        if cust in CUSTOMERS or cust == BOTH_CUST:
            it["cust"] = cust
        pos = next((i for i, x in enumerate(z["items"]) if x.get("extra")), len(z["items"]))  # ก่อน "หัวข้อเพิ่มเติม"
        z["items"].insert(pos, it)
        added += 1
    AREAS["new_items_v1"] = True
    print(f"[areas] เพิ่มหัวข้อใหม่ {added}/{len(NEW_ITEMS_V1)} หัวข้อ (★ ยังไม่นับใน % ของ Dashboard)"
          + (f" · ไม่พบพื้นที่: {', '.join(missing)}" if missing else ""), flush=True)
    _save_areas(AREAS, "new_items_v1")
    AREAS, ITEM_MAP = load_areas()


def clean_qty_cfg(q):
    """ตั้งค่าช่องนับจำนวนของหัวข้อ: {unit, label, rooms[]}"""
    if not isinstance(q, dict):
        return None
    rooms = []
    for r in q.get("rooms") or []:
        r = str(r).strip()[:40]
        if r and r not in rooms:
            rooms.append(r)
    return {"unit": (str(q.get("unit") or "คัน").strip()[:10] or "คัน"),
            "label": (str(q.get("label") or "จำนวน").strip()[:60] or "จำนวน"), "rooms": rooms[:30]}


def parse_qty(raw, cfg):
    """ค่าที่พนักงานกรอก → {"rooms": {ห้อง: จำนวน}, "total": รวม} หรือ None"""
    if not cfg or raw in (None, ""):
        return None
    try:
        d = json.loads(raw) if isinstance(raw, str) else raw
    except ValueError:
        raise HTTPException(400, "ข้อมูลจำนวนไม่ถูกต้อง")
    if not isinstance(d, dict):
        return None

    def num(v):
        if v in (None, ""):
            return None
        try:
            v = int(float(v))
        except (TypeError, ValueError):
            raise HTTPException(400, "จำนวนต้องเป็นตัวเลข")
        if v < 0 or v > 99999:
            raise HTTPException(400, "จำนวนไม่ถูกต้อง")
        return v
    rooms = {}
    for k, v in (d.get("rooms") or {}).items():
        k = str(k).strip()[:40]
        v = num(v)
        if k and v is not None:
            rooms[k] = v
    total = sum(rooms.values()) if rooms else num(d.get("total"))
    if total is None:
        return None
    return {"rooms": rooms, "total": total}


def _save_areas(cfg, tag="manual"):
    backup_dir = DATA_DIR / "areas_backup"
    backup_dir.mkdir(exist_ok=True)
    stamp = now().strftime("%Y%m%d_%H%M%S")
    (backup_dir / f"areas_{stamp}.json").write_text(AREAS_PATH.read_text(encoding="utf-8"), encoding="utf-8")
    tmp = AREAS_PATH.with_suffix(".tmp")
    tmp.write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(AREAS_PATH)


def _backfill_item_dates():
    """ใส่วันที่เพิ่ม/ลบรายการ (added/removed) ให้รายการเดิม โดยดูจากไฟล์สำรอง areas_backup
    (ไฟล์สำรองถูกสร้างทุกครั้งที่กดบันทึกในหน้าแก้ไขหัวข้อ = สถานะก่อนแก้)"""
    global AREAS, ITEM_MAP
    bdir = DATA_DIR / "areas_backup"
    files = sorted(bdir.glob("areas_*.json")) if bdir.exists() else []
    if not files:
        return
    snaps = []
    for f in files:
        try:
            t = datetime.strptime(f.stem.split("_", 1)[1], "%Y%m%d_%H%M%S").replace(tzinfo=TZ)
            cfg = json.loads(f.read_text(encoding="utf-8"))
            st = {it["id"]: bool(it.get("hidden")) for z in cfg["zones"] for it in z["items"]}
            snaps.append((t, st))
        except Exception:
            continue
    changed = False
    for z in AREAS["zones"]:
        for it in z["items"]:
            iid = it["id"]
            if "added" not in it:
                missing = [t for t, st in snaps if iid not in st]
                if missing:  # บันทึกครั้งหลังสุดที่ยังไม่มีรายการนี้ = ครั้งที่เพิ่ม
                    it["added"] = round_of(max(missing)).isoformat()
                    changed = True
            if it.get("hidden") and "removed" not in it:
                alive = [t for t, st in snaps if iid in st and not st[iid]]
                if alive:  # บันทึกครั้งหลังสุดที่รายการยังไม่ถูกลบ = ครั้งที่ลบ
                    it["removed"] = round_of(max(alive)).isoformat()
                    changed = True
    if changed:
        AREAS_PATH.write_text(json.dumps(AREAS, ensure_ascii=False, indent=1), encoding="utf-8")
        AREAS, ITEM_MAP = load_areas()


# ---------------------------------------------------------------- time helpers
def now() -> datetime:
    return datetime.now(TZ)


def round_of(dt: datetime) -> date:
    """รอบของเวลา dt: ก่อน 07:00 นับเป็นรอบของเมื่อวาน"""
    return dt.date() - timedelta(days=1) if dt.time() < CUTOFF else dt.date()


def round_window(d: date):
    start = datetime.combine(d, CUTOFF, TZ)
    return start, start + timedelta(days=1)


def parse_round(s: Optional[str]) -> date:
    if not s:
        return round_of(now())
    try:
        return date.fromisoformat(s)
    except ValueError:
        raise HTTPException(400, "รูปแบบวันที่ไม่ถูกต้อง (YYYY-MM-DD)")


# ---------------------------------------------------------------- database
@contextmanager
def db():
    con = sqlite3.connect(DB_PATH, timeout=30)
    con.row_factory = sqlite3.Row
    try:
        yield con
        con.commit()
    finally:
        con.close()


with db() as con:
    con.executescript(
        """
        PRAGMA journal_mode=WAL;
        CREATE TABLE IF NOT EXISTS submissions(
            id TEXT PRIMARY KEY,
            round_date TEXT NOT NULL,
            area_id TEXT NOT NULL,
            inspector TEXT NOT NULL,
            status TEXT NOT NULL,
            note TEXT,
            created_at TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS ix_sub_round ON submissions(round_date, area_id);
        CREATE TABLE IF NOT EXISTS retort_problems(
            id TEXT PRIMARY KEY,
            date TEXT NOT NULL,
            data TEXT NOT NULL,
            updated_at INTEGER NOT NULL
        );
        CREATE INDEX IF NOT EXISTS ix_rp_date ON retort_problems(date);
        CREATE TABLE IF NOT EXISTS images(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            submission_id TEXT NOT NULL,
            filename TEXT NOT NULL,
            seq INTEGER NOT NULL
        );
        CREATE INDEX IF NOT EXISTS ix_img_sub ON images(submission_id);
        CREATE TABLE IF NOT EXISTS followups(
            id TEXT PRIMARY KEY,
            data TEXT NOT NULL,
            status TEXT NOT NULL,
            created_at INTEGER NOT NULL,
            updated_at INTEGER NOT NULL
        );
        """
    )
    # เพิ่มคอลัมน์สำหรับหลายสไลด์ต่อ 1 การตรวจ (รองรับฐานข้อมูลเดิม)
    cols = {r[1] for r in con.execute("PRAGMA table_info(submissions)")}
    if "pages" not in cols:
        con.execute("ALTER TABLE submissions ADD COLUMN pages TEXT")
    if "topic" not in cols:
        con.execute("ALTER TABLE submissions ADD COLUMN topic TEXT")
    if "edited_at" not in cols:
        con.execute("ALTER TABLE submissions ADD COLUMN edited_at TEXT")
    if "late" not in cols:  # ส่งย้อนหลัง (ส่งหลังรอบนั้นปิดแล้ว)
        con.execute("ALTER TABLE submissions ADD COLUMN late INTEGER NOT NULL DEFAULT 0")
    if "qty" not in cols:  # จำนวนที่นับ เช่น วัตถุดิบ Delay แยกห้อง (JSON)
        con.execute("ALTER TABLE submissions ADD COLUMN qty TEXT")
    for c in ("remark", "action", "cause"):  # หมายเหตุ / การแก้ไข / สาเหตุ (เมื่อพบปัญหา)
        if c not in cols:
            con.execute(f"ALTER TABLE submissions ADD COLUMN {c} TEXT")
    cols = {r[1] for r in con.execute("PRAGMA table_info(images)")}
    if "page" not in cols:
        con.execute("ALTER TABLE images ADD COLUMN page INTEGER NOT NULL DEFAULT 0")


def is_video(fn: str) -> bool:
    return Path(fn).suffix.lower() in VIDEO_EXT


def media_info(fn: str) -> dict:
    v = is_video(fn)
    return {"file": fn, "url": f"/uploads/{fn}", "kind": "video" if v else "image",
            "poster": f"/uploads/{fn}.jpg" if v else f"/uploads/{fn}",
            "thumb": f"/t/{fn}"}


def thumb_path(fn: str) -> Path:
    return UPLOAD_DIR / (fn + ".t.jpg")


def make_thumb(fn: str, src_im: Optional[Image.Image] = None) -> Optional[Path]:
    """สร้างรูปย่อ (~25 KB) ไว้แสดงในรายการ รูปจริงยังอยู่ครบสำหรับดูเต็มจอและใช้ในสไลด์"""
    tp = thumb_path(fn)
    try:
        if src_im is None:
            src = UPLOAD_DIR / (fn + ".jpg" if is_video(fn) else fn)
            if not src.exists():
                return None
            with Image.open(src) as im0:
                im = im0.convert("RGB")
        else:
            im = src_im.copy()
        im.thumbnail((THUMB_SIDE, THUMB_SIDE))
        tmp = tp.with_name(tp.name + ".tmp")
        im.save(tmp, "JPEG", quality=72, optimize=True)
        tmp.replace(tp)
        return tp
    except Exception as e:  # noqa
        print("[thumb] error:", fn, e, flush=True)
        return None


def fetch_submissions(rd: date, area_id: Optional[str] = None):
    q = "SELECT * FROM submissions WHERE round_date=?"
    args = [rd.isoformat()]
    if area_id:
        q += " AND area_id=?"
        args.append(area_id)
    q += " ORDER BY created_at"
    with db() as con:
        subs = [dict(r) for r in con.execute(q, args)]
        for s in subs:
            s["_img"] = [(r["page"], r["filename"]) for r in con.execute(
                "SELECT page, filename FROM images WHERE submission_id=? ORDER BY page, seq", (s["id"],))]
    for s in subs:
        it = ITEM_MAP.get(s["area_id"], {})
        s["area_name"] = it.get("name", s["area_id"])
        s["zone"] = it.get("zone", "")
        s["zone_id"] = it.get("zone_id", "")
        s["proc"] = it.get("proc", "")  # ป้ายกระบวนการ (Rawmat / จุดเตรียม / ...)
        s["cust"] = it.get("cust", "")  # ลูกค้า: Weruva / Smucker / both (ว่าง = ไม่มีลูกค้า)
        try:
            s["qty"] = json.loads(s.get("qty") or "null")
        except ValueError:
            s["qty"] = None
        s["qty_cfg"] = it.get("qty")
        try:
            notes = json.loads(s.get("pages") or "null")
        except ValueError:
            notes = None
        if not isinstance(notes, list) or not notes:
            notes = [s.get("note") or ""]
        n = max(len(notes), 1 + max([pg for pg, _ in s["_img"]], default=0))
        notes += [""] * (n - len(notes))
        s["pages"] = [{
            "note": notes[i],
            "images": [f for pg, f in s["_img"] if pg == i],
            "image_urls": [f"/uploads/{f}" for pg, f in s["_img"] if pg == i],
            "media": [media_info(f) for pg, f in s["_img"] if pg == i],
        } for i in range(n)]
        s["images"] = [f for _, f in s["_img"]]
        s["image_urls"] = [f"/uploads/{f}" for f in s["images"]]
        del s["_img"]
    return subs


# ---------------------------------------------------------------- แยกชุดสไลด์: Process Control / GMP
DECKS = {
    "pc": {"file": "ProcessControl", "name": "Process Control"},
    "gmp": {"file": "GMP", "name": "GMP"},
    "retort": {"file": "RetortAnalysis", "name": "Retort Analysis"},
}


def live_decks():
    """ชุดสไลด์ที่มีพื้นที่ใช้งานอยู่ (Process Control มีเสมอ) — ไม่สร้างไฟล์ว่างให้หัวข้อที่ยังไม่มีพื้นที่"""
    used = {zone_deck(z) for z in AREAS["zones"] if not z.get("hidden")}
    return [d for d in DECKS if d == "pc" or d in used]


def zone_deck(z) -> str:
    """พื้นที่ไหนไปอยู่สไลด์ชุดไหน: กำหนดด้วย "deck" ใน areas.json หรือชื่อพื้นที่ขึ้นต้นด้วย GMP"""
    d = z.get("deck")
    if d in DECKS:
        return d
    return "gmp" if str(z.get("name", "")).strip().upper().startswith("GMP") else "pc"


def status_for(rd: date, deck: Optional[str] = None):
    subs = fetch_submissions(rd)
    by_area = {}
    for s in subs:
        by_area.setdefault(s["area_id"], []).append(s)
    zones = []
    total = done = issues = 0
    for z in zones_for_round(rd.isoformat(), set(by_area)):
        if deck and zone_deck(z) != deck:
            continue
        items = []
        zdone = 0
        for it in z["items"]:
            ss = by_area.get(it["id"], [])
            last = ss[-1] if ss else None
            has_issue = any(x["status"] == "พบปัญหา" for x in ss)
            items.append({
                **it,
                "checked": bool(ss),
                "noact": bool(ss) and all(x["status"] == NO_ACT for x in ss),  # ส่งแล้วแต่ไม่มีกิจกรรมทุกครั้ง
                "count": len(ss),
                "issue": has_issue,
                "last_inspector": last["inspector"] if last else None,
                "last_time": last["created_at"] if last else None,
                "images": sum(len(x["images"]) for x in ss),
            })
            issues += has_issue
            if not it.get("extra"):  # หัวข้อเพิ่มเติมไม่นับรวมในยอด checklist
                zdone += bool(ss)
        ztotal = sum(1 for it in items if not it.get("extra"))
        total += ztotal
        done += zdone
        zones.append({"id": z["id"], "name": z["name"], "owner": z["owner"], "deck": zone_deck(z), "sub": z.get("sub", ""),
                      "done": zdone, "total": ztotal, "items": items})
    start, end = round_window(rd)
    return {
        "round": rd.isoformat(),
        "window": [start.isoformat(), end.isoformat()],
        "closed": now() >= end,
        "total": total, "done": done, "issues": issues,
        "zones": zones,
    }


# ---------------------------------------------------------------- slides
def slide_path(rd: date, deck: str = "pc") -> Path:
    return SLIDE_DIR / f"{DECKS[deck]['file']}_{rd.isoformat()}.pptx"


def _day_data(rd: date, deck: str):
    data = status_for(rd, deck)
    zone_ids = {z["id"] for z in data["zones"]}
    subs = [x for x in fetch_submissions(rd) if x["zone_id"] in zone_ids]
    return data, subs


def generate_slides(rd: date, deck: str = "pc", order: str = "zone") -> bytes:
    data, subs = _day_data(rd, deck)
    return build_pptx(AREAS["title"], data, subs, UPLOAD_DIR, CUTOFF, report_name=DECKS[deck]["name"], order=order)


def generate_slides_range(start: date, end: date, deck: str = "pc", order: str = "zone") -> bytes:
    """รวมหลายวันเป็นไฟล์เดียว: สิ่งที่พบปัญหาของทุกวันขึ้นก่อน แล้วตามด้วยรายละเอียดของแต่ละวัน"""
    days = [_day_data(start + timedelta(days=i), deck) for i in range((end - start).days + 1)]
    return build_pptx_multi(AREAS["title"], days, UPLOAD_DIR, CUTOFF, report_name=DECKS[deck]["name"], order=order)


_LIVE = {}  # (round, deck) -> (signature, bytes) : ไม่มีข้อมูลเปลี่ยน = ไม่ต้องสร้างสไลด์ใหม่
_LIVE_LOCK = threading.Lock()


def _round_signature(rd: date) -> tuple:
    with db() as con:
        r = con.execute("SELECT COUNT(*), MAX(created_at), MAX(COALESCE(edited_at,'')) FROM submissions WHERE round_date=?",
                        (rd.isoformat(),)).fetchone()
    return tuple(r) + (AREAS_PATH.stat().st_mtime,)


def live_slides(rd: date, deck: str, order: str = "zone") -> bytes:
    sig = _round_signature(rd)
    key = (rd, deck, order)
    hit = _LIVE.get(key)
    if hit and hit[0] == sig:
        return hit[1]
    with _LIVE_LOCK:  # สร้างทีละชุด กันเครื่องทำงานหนักพร้อมกันหลายคำขอ
        hit = _LIVE.get(key)
        if hit and hit[0] == sig:
            return hit[1]
        blob = generate_slides(rd, deck, order)
        if len(_LIVE) > 6:
            _LIVE.clear()
        _LIVE[key] = (sig, blob)
        return blob


def _order(v: str) -> str:
    return "proc" if v == "proc" else "zone"


def archive_round(rd: date):
    for deck in live_decks():
        blob = generate_slides(rd, deck)
        tmp = slide_path(rd, deck).with_suffix(".tmp")
        tmp.write_bytes(blob)
        tmp.replace(slide_path(rd, deck))
    print(f"[cutoff] archived slides for round {rd}", flush=True)


def scheduler_loop():
    """ทุก 30 วินาที: ถ้ารอบก่อนหน้าปิดแล้วแต่ยังไม่มีไฟล์สไลด์ → สร้างเก็บไว้"""
    while True:
        try:
            prev = round_of(now()) - timedelta(days=1)
            if not all(slide_path(prev, d).exists() for d in live_decks()):
                archive_round(prev)
        except Exception as e:  # noqa
            print("[cutoff] error:", e, flush=True)
        _time.sleep(30)


# ---------------------------------------------------------------- app
app = FastAPI(title="Process Control Form")
# บีบอัดหน้าเว็บและข้อมูลก่อนส่ง (เล็กลง ~4-5 เท่า ช่วยมากเมื่อเน็ตช้า)
app.add_middleware(GZipMiddleware, minimum_size=800)


@app.on_event("startup")
def _start():
    try:
        _backfill_item_dates()
        _migrate_gmp_areas()
        _migrate_wh_points()
        _migrate_delay_qty()
        _migrate_cust_list()
        _migrate_new_items_v1()
    except Exception as e:  # noqa
        print("[areas] backfill error:", e, flush=True)
    try:
        _migrate_pf1_retorts()
    except Exception as e:  # noqa
        print("[retort] PF1 retorts migrate error:", e, flush=True)
    threading.Thread(target=scheduler_loop, daemon=True).start()
    print(f"[slides] รูปแบบสไลด์เวอร์ชัน {SLIDES_VERSION}", flush=True)


@app.middleware("http")
async def no_stale_pages(request: Request, call_next):
    """ให้มือถือโหลดหน้าเว็บเวอร์ชันล่าสุดเสมอ (กันเบราว์เซอร์จำหน้าเก่าไว้)"""
    resp = await call_next(request)
    path = request.url.path
    if path.startswith(("/uploads/", "/t/", "/fonts/")) and resp.status_code == 200:
        # ชื่อไฟล์รูปไม่ซ้ำและไม่ถูกเขียนทับ → ให้มือถือจำไว้ได้เลย ไม่ต้องโหลดซ้ำ
        resp.headers["Cache-Control"] = "public, max-age=31536000, immutable"
    elif (path == "/" or path.endswith((".html", ".js", ".webmanifest")) or path.startswith("/api/")) \
            and "immutable" not in resp.headers.get("Cache-Control", ""):  # รูป Retort ตั้งให้จำได้ถาวรเอง ไม่ต้องเขียนทับ
        resp.headers["Cache-Control"] = "no-cache"
    return resp


@app.middleware("http")
async def access_gate(request: Request, call_next):
    path = request.url.path
    if ACCESS_CODE and path.startswith(("/api/", "/uploads/", "/t/")) \
            and path not in ("/api/login",) and request.cookies.get("pc_access") != ACCESS_TOKEN:
        return JSONResponse({"detail": "ACCESS_REQUIRED"}, status_code=401)
    return await call_next(request)


@app.post("/api/login")
async def api_login(request: Request):
    form = await request.form()
    if not ACCESS_CODE or str(form.get("code") or "") == ACCESS_CODE:
        resp = JSONResponse({"ok": True})
        resp.set_cookie("pc_access", ACCESS_TOKEN, max_age=365 * 24 * 3600, httponly=True, samesite="lax")
        return resp
    raise HTTPException(403, "รหัสไม่ถูกต้อง")


@app.get("/api/config")
def api_config():
    n = now()
    rd = round_of(n)
    start, end = round_window(rd)
    return {
        "title": AREAS["title"],
        "zones": visible_zones(),
        "cutoff": CUTOFF.strftime("%H:%M"),
        "server_time": n.isoformat(),
        "current_round": rd.isoformat(),
        "round_end": end.isoformat(),
        "pin_required": bool(ADMIN_PIN),
        "max_images": MAX_IMAGES,
        "max_pages": MAX_PAGES,
        "max_video_sec": MAX_VIDEO_SEC,
        "max_video_mb": MAX_VIDEO_MB,
        "backdate_days": BACKDATE_DAYS,
        "process_tags": process_tags(),
    }


@app.get("/api/status")
def api_status(round: Optional[str] = Query(None)):
    return status_for(parse_round(round))


@app.get("/api/submissions")
def api_submissions(round: Optional[str] = None, area: Optional[str] = None):
    return fetch_submissions(parse_round(round), area)


def _clean(t: str) -> str:
    return (t or "").replace("\r\n", "\n").replace("\r", "\n").strip()


def pick_round(value) -> date:
    """รอบที่จะบันทึก: ไม่ระบุ = รอบปัจจุบัน / ระบุได้ย้อนหลังไม่เกิน BACKDATE_DAYS วัน"""
    cur = round_of(now())
    if not value:
        return cur
    try:
        rd = date.fromisoformat(str(value))
    except ValueError:
        raise HTTPException(400, "รูปแบบวันที่ไม่ถูกต้อง")
    if rd > cur:
        raise HTTPException(400, "ส่งล่วงหน้าไม่ได้")
    if (cur - rd).days > BACKDATE_DAYS:
        raise HTTPException(400, f"ส่ง/แก้ไขย้อนหลังได้ไม่เกิน {BACKDATE_DAYS} วัน")
    return rd


def regen_if_closed(rd: date):
    """ถ้ารอบนั้นปิดไปแล้ว (มีไฟล์สไลด์) ให้สร้างสไลด์ใหม่เบื้องหลัง เพื่อให้มีข้อมูลที่ส่งย้อนหลัง"""
    if rd < round_of(now()) and any(slide_path(rd, dk).exists() for dk in DECKS):
        threading.Thread(target=lambda: _safe_archive(rd), daemon=True).start()


def _safe_archive(rd: date):
    try:
        archive_round(rd)
    except Exception as e:  # noqa
        print("[regen] error:", e, flush=True)


async def save_media(form, i: int, rd: date, prefix: str, start_seq: int = 0):
    """บันทึกรูป (images_i) และวิดีโอ (videos_i + vposters_i) ของสไลด์ที่ i คืนรายชื่อไฟล์ตามลำดับ"""
    out = []
    day = rd.isoformat()
    (UPLOAD_DIR / day).mkdir(parents=True, exist_ok=True)
    seq = start_seq
    for f in [f for f in form.getlist(f"images_{i}") if getattr(f, "filename", None)]:
        raw = await f.read()
        seq += 1
        rel = f"{day}/{prefix}_{i + 1}_{seq}.jpg"
        # ย่อ/บันทึกรูปในเธรดแยก เพื่อไม่ให้คนอื่นที่ใช้งานอยู่ค้าง
        if not await run_in_threadpool(_store_image, raw, rel):
            raise HTTPException(400, f"ไฟล์ {f.filename} ไม่ใช่รูปภาพที่รองรับ")
        out.append(rel)
    videos = [f for f in form.getlist(f"videos_{i}") if getattr(f, "filename", None)]
    posters = [f for f in form.getlist(f"vposters_{i}") if getattr(f, "filename", None)]
    for k, f in enumerate(videos):
        ext = Path(f.filename).suffix.lower()
        if ext not in VIDEO_EXT:
            ext = ".mp4" if "mp4" in (f.content_type or "") else (".webm" if "webm" in (f.content_type or "") else ".mov")
        raw = await f.read()
        if len(raw) > MAX_VIDEO_MB * 1024 * 1024:
            raise HTTPException(400, f"วิดีโอ {f.filename} ใหญ่เกิน {int(MAX_VIDEO_MB)} MB")
        seq += 1
        rel = f"{day}/{prefix}_{i + 1}_{seq}{ext}"
        praw = await posters[k].read() if k < len(posters) else b""
        await run_in_threadpool(_store_video, raw, praw, rel)
        out.append(rel)
    return out


def _store_image(raw: bytes, rel: str) -> bool:
    try:
        im = Image.open(io.BytesIO(raw))
        im = ImageOps.exif_transpose(im).convert("RGB")
        im.thumbnail((MAX_SIDE, MAX_SIDE))
    except Exception:
        return False
    im.save(UPLOAD_DIR / rel, "JPEG", quality=82, optimize=True)
    make_thumb(rel, im)
    return True


def _store_video(raw: bytes, praw: bytes, rel: str):
    (UPLOAD_DIR / rel).write_bytes(raw)
    # ภาพหน้าปกวิดีโอ (ใช้แสดงในเว็บและสไลด์)
    try:
        pim = Image.open(io.BytesIO(praw)).convert("RGB")
        pim.thumbnail((MAX_SIDE, MAX_SIDE))
    except Exception:
        pim = Image.new("RGB", (1280, 720), (30, 41, 59))
    pim.save(UPLOAD_DIR / (rel + ".jpg"), "JPEG", quality=82)
    make_thumb(rel, pim)


def _remove_media(fn: str):
    for f in ((fn, fn + ".jpg") if is_video(fn) else (fn,)) + (fn + ".t.jpg",):
        try:
            (UPLOAD_DIR / f).unlink()
        except FileNotFoundError:
            pass


@app.post("/api/submissions")
async def api_create(request: Request):
    """รับข้อมูลแบบหลายสไลด์: page_count, note_i, images_i, videos_i, vposters_i
    + status, topic, remark (หมายเหตุ), action (การแก้ไขเมื่อพบปัญหา)"""
    form = await request.form()
    area_id = str(form.get("area_id") or "")
    inspector = str(form.get("inspector") or "").strip()[:60]
    status = str(form.get("status") or "ปกติ")
    topic = _clean(str(form.get("topic") or ""))[:120]
    remark = _clean(str(form.get("remark") or ""))[:1000]
    action = _clean(str(form.get("action") or ""))[:2000]
    cause = _clean(str(form.get("cause") or ""))[:2000]
    if area_id not in ITEM_MAP or ITEM_MAP[area_id].get("hidden"):
        raise HTTPException(400, "ไม่พบพื้นที่นี้")
    if status not in STATUSES:
        status = "ปกติ"
    qcfg = ITEM_MAP[area_id].get("qty")
    qty = None if status == NO_ACT else parse_qty(form.get("qty"), qcfg)  # ไม่มีกิจกรรม = ไม่ต้องนับจำนวน
    if qcfg and qty is None and status != NO_ACT:
        raise HTTPException(400, f"กรุณากรอกจำนวน{qcfg.get('label', '')}")
    if not inspector:
        raise HTTPException(400, "กรุณาใส่ชื่อผู้ตรวจ")
    if status != "พบปัญหา":
        action = cause = ""

    try:
        page_count = int(form.get("page_count") or 0)
    except ValueError:
        page_count = 0
    if page_count <= 0:
        page_count = 1
    if page_count > MAX_PAGES:
        raise HTTPException(400, f"ทำได้ไม่เกิน {MAX_PAGES} สไลด์ต่อครั้ง")

    n = now()
    rd = pick_round(form.get("round"))  # ค่าเริ่มต้น = รอบปัจจุบัน (ตัดรอบที่ฝั่ง server)
    late = 1 if rd < round_of(n) else 0
    sid = uuid.uuid4().hex[:12]
    pages, saved = [], []  # saved: (page, seq, filename)
    for i in range(page_count):
        note = _clean(str(form.get(f"note_{i}") or (form.get("note") if i == 0 else "") or ""))[:4000]
        nfiles = len([f for f in form.getlist(f"images_{i}") + form.getlist(f"videos_{i}") if getattr(f, "filename", None)])
        if nfiles > MAX_IMAGES:
            raise HTTPException(400, f"แนบรูป/วิดีโอได้ไม่เกิน {MAX_IMAGES} ไฟล์ต่อสไลด์")
        if not note and not nfiles:
            continue  # ข้ามสไลด์ว่าง
        pi = len(pages)
        files = await save_media(form, i, rd, sid)
        saved += [(pi, k, fn) for k, fn in enumerate(files)]
        pages.append(note)
    if not pages:
        pages = [""]

    with db() as con:
        con.execute(
            "INSERT INTO submissions(id, round_date, area_id, inspector, status, note, created_at, pages, topic, remark, action, late, cause, qty) "
            "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (sid, rd.isoformat(), area_id, inspector, status, "\n\n".join(x for x in pages if x),
             n.isoformat(timespec="seconds"), json.dumps(pages, ensure_ascii=False), topic, remark, action, late, cause,
             json.dumps(qty, ensure_ascii=False) if qty else None),
        )
        con.executemany(
            "INSERT INTO images(submission_id, filename, seq, page) VALUES(?,?,?,?)",
            [(sid, fn, seq, pg) for pg, seq, fn in saved],
        )
    regen_if_closed(rd)
    return {"ok": True, "id": sid, "round": rd.isoformat(), "pages": len(pages), "late": bool(late)}


@app.put("/api/submissions/{sid}")
async def api_update(sid: str, request: Request):
    """แก้ไขรายการที่บันทึกแล้ว (เฉพาะรอบปัจจุบัน): ข้อความ, ผลตรวจ, หัวข้อ, รูป (ลบรูปเดิม/เพิ่มรูปใหม่), จำนวนสไลด์
    ฟิลด์: status, topic, inspector, page_count, note_i, keep_i (JSON รายชื่อรูปเดิมที่เก็บไว้), images_i (รูปใหม่)"""
    with db() as con:
        row = con.execute("SELECT * FROM submissions WHERE id=?", (sid,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบรายการ")
        old_files = {r["filename"] for r in con.execute("SELECT filename FROM images WHERE submission_id=?", (sid,))}
    rd = date.fromisoformat(row["round_date"])
    pick_round(rd.isoformat())  # แก้ไขย้อนหลังได้ไม่เกิน BACKDATE_DAYS วัน

    form = await request.form()
    status = str(form.get("status") or row["status"])
    if status not in STATUSES:
        status = row["status"]
    topic = _clean(str(form.get("topic") if form.get("topic") is not None else (row["topic"] or "")))[:120]
    inspector = str(form.get("inspector") or row["inspector"]).strip()[:60] or row["inspector"]
    try:
        page_count = int(form.get("page_count") or 0)
    except ValueError:
        page_count = 0
    if page_count < 1 or page_count > MAX_PAGES:
        raise HTTPException(400, "จำนวนสไลด์ไม่ถูกต้อง")

    remark = _clean(str(form.get("remark") if form.get("remark") is not None else (row["remark"] or "")))[:1000]
    action = _clean(str(form.get("action") if form.get("action") is not None else (row["action"] or "")))[:2000]
    cause = _clean(str(form.get("cause") if form.get("cause") is not None else (row["cause"] or "")))[:2000]
    if status != "พบปัญหา":
        action = cause = ""
    qcfg = ITEM_MAP.get(row["area_id"], {}).get("qty")
    if form.get("qty") is not None:
        qty = parse_qty(form.get("qty"), qcfg)
        qty_s = json.dumps(qty, ensure_ascii=False) if qty else None
    else:
        qty_s = row["qty"]
    if status == NO_ACT:
        qty_s = None

    pages, saved = [], []
    stamp = uuid.uuid4().hex[:4]
    for i in range(page_count):
        note = _clean(str(form.get(f"note_{i}") or ""))[:4000]
        try:
            keep = [f for f in json.loads(str(form.get(f"keep_{i}") or "[]")) if f in old_files]
        except ValueError:
            keep = []
        nnew = len([f for f in form.getlist(f"images_{i}") + form.getlist(f"videos_{i}") if getattr(f, "filename", None)])
        if len(keep) + nnew > MAX_IMAGES:
            raise HTTPException(400, f"แนบรูป/วิดีโอได้ไม่เกิน {MAX_IMAGES} ไฟล์ต่อสไลด์")
        if not (note or keep or nnew):
            continue
        pi = len(pages)
        new = await save_media(form, i, rd, f"{sid}_e{stamp}")
        saved += [(pi, k, fn) for k, fn in enumerate(keep + new)]
        pages.append(note)
    if not pages:
        pages = [""]

    with db() as con:
        con.execute("UPDATE submissions SET status=?, topic=?, inspector=?, note=?, pages=?, edited_at=?, remark=?, action=?, cause=?, qty=? WHERE id=?",
                    (status, topic, inspector, "\n\n".join(x for x in pages if x),
                     json.dumps(pages, ensure_ascii=False), now().isoformat(timespec="seconds"), remark, action, cause, qty_s, sid))
        con.execute("DELETE FROM images WHERE submission_id=?", (sid,))
        con.executemany("INSERT INTO images(submission_id, filename, seq, page) VALUES(?,?,?,?)",
                        [(sid, fn, seq, pg) for pg, seq, fn in saved])
    kept = {fn for _, _, fn in saved}
    for fn in old_files - kept:  # ลบไฟล์ที่เอาออก
        _remove_media(fn)
    regen_if_closed(rd)
    return {"ok": True, "id": sid}


@app.delete("/api/submissions/{sid}")
def api_delete(sid: str, pin: str = ""):
    if ADMIN_PIN and pin != ADMIN_PIN:
        raise HTTPException(403, "PIN ไม่ถูกต้อง")
    with db() as con:
        row = con.execute("SELECT round_date FROM submissions WHERE id=?", (sid,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบรายการ")
        files = [r["filename"] for r in con.execute("SELECT filename FROM images WHERE submission_id=?", (sid,))]
        con.execute("DELETE FROM images WHERE submission_id=?", (sid,))
        con.execute("DELETE FROM submissions WHERE id=?", (sid,))
    for fn in files:
        _remove_media(fn)
    # ถ้าเป็นรอบที่ปิดไปแล้ว ให้สร้างสไลด์ใหม่
    rd = date.fromisoformat(row["round_date"])
    regen_if_closed(rd)
    return {"ok": True}


@app.get("/api/heatmap")
def api_heatmap(days: int = Query(14, ge=3, le=62), start: Optional[str] = None, end: Optional[str] = None, cust: str = ""):
    """ข้อมูล Dashboard heat map: สถานะรายการ × วัน ย้อนหลัง N รอบ หรือช่วงที่เลือก (?start=&end=)
    code: 0 = ยังไม่ตรวจ, 1 = ตรวจแล้วปกติ, 2 = พบปัญหา, 3 = ไม่มีกิจกรรม (ไม่นับใน % และจำนวนครั้ง)"""
    cur = round_of(now())
    if start and end:
        a, b = parse_round(start), min(parse_round(end), cur)
        if a > b:
            a, b = b, a
        if (b - a).days > 200:  # หน้าเว็บจำกัดช่วงที่เลือกไว้ 93 วัน + ดึงช่วงก่อนหน้าไว้เทียบ
            raise HTTPException(400, "ช่วงวันที่ยาวเกินไป")
        rounds = [(a + timedelta(days=i)).isoformat() for i in range((b - a).days + 1)]
    else:
        rounds = [(cur - timedelta(days=i)).isoformat() for i in range(days - 1, -1, -1)]
    with db() as con:
        rows = con.execute(
            "SELECT round_date, area_id, SUM(status<>?) n, SUM(status=?) na, MAX(status='พบปัญหา') ng, SUM(status='พบปัญหา') ngn "
            "FROM submissions WHERE round_date BETWEEN ? AND ? GROUP BY round_date, area_id", (NO_ACT, NO_ACT, rounds[0], rounds[-1])).fetchall()
        qrows = con.execute("SELECT round_date, area_id, qty FROM submissions WHERE qty IS NOT NULL AND round_date BETWEEN ? AND ? "
                            "ORDER BY created_at", (rounds[0], rounds[-1])).fetchall()
    # code 3 = ส่งแล้วแต่ไม่มีกิจกรรมทั้งวัน · จำนวนครั้ง (n) ไม่รวมรายการที่ไม่มีกิจกรรม
    cell = {(r["round_date"], r["area_id"]): (2 if r["ng"] else 1 if r["n"] else 3, r["n"] or 0, r["ngn"] or 0) for r in rows}
    qcell = {}  # จำนวนที่นับ (ใช้ค่าล่าสุดของวันนั้น)
    for r in qrows:
        try:
            qcell[(r["round_date"], r["area_id"])] = json.loads(r["qty"])
        except ValueError:
            pass
    # เลือกลูกค้า: นับเฉพาะหัวข้อของลูกค้านั้น (รวมหัวข้อที่เป็นของทั้งสองลูกค้า)
    # any = ลูกค้าทั้งหมด (ไม่รวมหัวข้อที่ไม่มีลูกค้า) · none = เฉพาะหัวข้อที่ไม่มีลูกค้า
    has_cust = lambda it: it.get("cust") in CUSTOMERS + (BOTH_CUST,)
    cust_ok = (lambda it: True) if cust not in CUST_FILTERS else \
        has_cust if cust == "any" else (lambda it: not has_cust(it)) if cust == "none" else \
        (lambda it: it.get("cust") in (cust, BOTH_CUST))
    zones = []
    for z in AREAS["zones"]:
        visible = not z.get("hidden")
        items, done, totals = [], [0] * len(rounds), [0] * len(rounds)
        for it in z["items"]:
            if not it.get("extra") and not cust_ok(it):
                continue
            codes, counts, ngs = [], [], []
            for di, rd in enumerate(rounds):
                c, n, ngn = cell.get((rd, it["id"]), (0, 0, 0))
                codes.append(c)
                counts.append(n)
                ngs.append(ngn)  # จำนวนครั้งที่พบปัญหาในวันนั้น
                # % ย้อนหลังคิดจากรายการที่มีอยู่จริงในวันนั้น (วันที่ไม่มีกิจกรรมไม่นับ)
                if c != 3 and not it.get("extra") and not it.get("pilot") and item_active(it, rd, bool(c)):
                    totals[di] += 1
                    done[di] += bool(c)
            if visible and not it.get("hidden"):
                row = {"id": it["id"], "no": it["no"], "name": it["name"], "extra": bool(it.get("extra")), "pilot": bool(it.get("pilot")),
                       "proc": it.get("proc", ""), "cust": it.get("cust", ""),
                       "codes": codes, "counts": counts, "ngs": ngs}
                if it.get("qty"):
                    row["qty_cfg"] = it["qty"]
                    row["qty"] = [qcell.get((rd, it["id"])) for rd in rounds]
                items.append(row)
        if not visible:  # พื้นที่ที่ลบ/แยกไปแล้ว (เช่น WH เดิม, GMP เดิม) ไม่ต้องแสดงใน Dashboard
            continue
        total = sum(1 for it in z["items"] if not it.get("extra") and not it.get("hidden") and not it.get("pilot") and cust_ok(it))
        if cust and not any(not r["extra"] for r in items):  # ไม่มีหัวข้อของลูกค้านี้ในพื้นที่นี้ (หัวข้อ ★ ยังแสดง)
            continue
        zones.append({"id": z["id"], "name": z["name"], "owner": z["owner"], "deck": zone_deck(z),
                      "total": total, "totals": totals,
                      "done": done, "items": items})
    return {"rounds": rounds, "current": cur.isoformat(), "zones": zones, "cust": cust if cust in CUST_FILTERS else ""}


@app.get("/api/inspectors")
def api_inspectors():
    """รายชื่อผู้ตรวจที่เคยส่ง (ล่าสุดก่อน) ไว้ให้เลือก ไม่ต้องพิมพ์ใหม่"""
    since = (now() - timedelta(days=90)).isoformat()
    with db() as con:
        rows = con.execute(
            "SELECT inspector, MAX(created_at) last FROM submissions WHERE created_at >= ? "
            "GROUP BY inspector ORDER BY last DESC LIMIT 60", (since,)).fetchall()
    return [r["inspector"] for r in rows if r["inspector"]]


@app.get("/api/rounds")
def api_rounds(days: int = Query(ROUNDS_LIMIT, ge=1, le=3660)):
    """รายการรอบย้อนหลัง (ค่าเริ่มต้น 60 วันล่าสุด — วันเก่ากว่านั้นเลือกจากปฏิทินในแท็บสไลด์)"""
    since = (round_of(now()) - timedelta(days=days - 1)).isoformat()
    with db() as con:
        rows = con.execute(
            "SELECT round_date, COUNT(DISTINCT area_id) areas, COUNT(*) subs FROM submissions "
            "WHERE round_date >= ? GROUP BY round_date", (since,)
        ).fetchall()
    extras = [k for k, v in ITEM_MAP.items() if v.get("extra")]
    with db() as con:
        ex = {r[0]: r[1] for r in con.execute(
            "SELECT round_date, COUNT(DISTINCT area_id) FROM submissions WHERE round_date >= ? AND area_id IN (%s) "
            "GROUP BY round_date" % ",".join("?" * len(extras)), [since] + extras)} if extras else {}
    info = {r["round_date"]: {"areas": r["areas"] - ex.get(r["round_date"], 0), "subs": r["subs"]} for r in rows}
    for p in SLIDE_DIR.glob("*_*.pptx"):
        d = p.stem.rsplit("_", 1)[1]
        if d >= since:
            info.setdefault(d, {"areas": 0, "subs": 0})
    cur = round_of(now()).isoformat()
    info.setdefault(cur, {"areas": 0, "subs": 0})
    total = sum(1 for z in visible_zones() for it in z["items"] if not it.get("extra"))
    out = []
    for rd in sorted(info, reverse=True):
        out.append({
            "round": rd, **info[rd], "total": total,
            "current": rd == cur,
            "archived": slide_path(date.fromisoformat(rd)).exists(),
            "archived_decks": [d for d in DECKS if slide_path(date.fromisoformat(rd), d).exists()],
        })
    return out


@app.post("/api/slides/check")
async def api_slides_check(request: Request):
    data = await request.json()
    if str(data.get("pin") or "") != SLIDE_PIN:
        raise HTTPException(403, "รหัสไม่ถูกต้อง")
    return {"ok": True}


MAX_RANGE_DAYS = int(os.environ.get("MAX_RANGE_DAYS", "31"))


@app.get("/api/slides/range.pptx")
def api_slides_range(start: str, end: str, pin: str = "", deck: str = "pc", order: str = "zone"):
    """สไลด์รวมหลายวัน ?start=YYYY-MM-DD&end=YYYY-MM-DD"""
    if pin != SLIDE_PIN:
        raise HTTPException(403, "ต้องใส่รหัสก่อน Generate สไลด์")
    if deck not in DECKS:
        raise HTTPException(400, "ชุดสไลด์ไม่ถูกต้อง")
    a, b = parse_round(start), parse_round(end)
    if a > b:
        a, b = b, a
    b = min(b, round_of(now()))
    if a > b:
        raise HTTPException(400, "ช่วงวันที่ไม่ถูกต้อง")
    n = (b - a).days + 1
    if n > MAX_RANGE_DAYS:
        raise HTTPException(400, f"รวมได้ไม่เกิน {MAX_RANGE_DAYS} วันต่อไฟล์")
    order = _order(order)
    if n == 1:
        return api_slides(a.isoformat(), live=int(a == round_of(now())), pin=pin, deck=deck, order=order)
    key = ("range", a, b, deck, order)
    sig = tuple(_round_signature(a + timedelta(days=i)) for i in range(n))
    hit = _LIVE.get(key)
    if hit and hit[0] == sig:
        blob = hit[1]
    else:
        with _LIVE_LOCK:
            blob = generate_slides_range(a, b, deck, order)
            if len(_LIVE) > 6:
                _LIVE.clear()
            _LIVE[key] = (sig, blob)
    fname = f"{DECKS[deck]['file']}_{a.isoformat()}_to_{b.isoformat()}" + ("_by_process" if order == "proc" else "") + ".pptx"
    return Response(blob, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    headers={"Content-Disposition": f'attachment; filename="{fname}"'})


@app.get("/api/slides/{rd}.pptx")
def api_slides(rd: str, live: int = 0, pin: str = "", deck: str = "pc", order: str = "zone"):
    """order: zone = เรียงตามพื้นที่ (ค่าเดิม) / proc = เรียงตามกระบวนการ (Rawmat → จุดเตรียม → ห้องเคมี → ห้องอิ้ง → บรรจุ → ...)"""
    order = _order(order)
    if pin != SLIDE_PIN:
        raise HTTPException(403, "ต้องใส่รหัสก่อน Generate สไลด์")
    if deck not in DECKS:
        raise HTTPException(400, "ชุดสไลด์ไม่ถูกต้อง")
    d = parse_round(rd)
    p = slide_path(d, deck)
    fname = f"{DECKS[deck]['file']}_{d.isoformat()}.pptx"
    if order != "zone":
        live = 1  # ไฟล์ที่เก็บไว้ตอนตัดรอบเรียงตามพื้นที่ → แบบตามกระบวนการสร้างใหม่ทุกครั้ง
    if p.exists() and not live and p.stat().st_mtime < SLIDES_CODE_MTIME:
        _safe_archive(d)  # ไฟล์ถูกสร้างด้วยโค้ดเวอร์ชันเก่า → สร้างใหม่ด้วยรูปแบบล่าสุด
    if p.exists() and not live:
        return FileResponse(p, filename=fname,
                            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    blob = live_slides(d, deck, order)
    if order == "proc":
        fname = fname.replace(".pptx", "_by_process.pptx")
    return Response(blob, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    headers={"Content-Disposition": f'attachment; filename="{fname}"'})


# ---------------------------------------------------------------- แก้ไขหัวข้อ Checklist (ต้องใช้รหัส EDIT_PIN)
def _check_edit_pin(pin: str):
    if str(pin or "") != EDIT_PIN:
        raise HTTPException(403, "รหัสแก้ไขไม่ถูกต้อง")


@app.post("/api/admin/check")
async def api_admin_check(request: Request):
    data = await request.json()
    _check_edit_pin(data.get("pin"))
    return {"ok": True, "zones": visible_zones()}


@app.post("/api/admin/areas")
async def api_admin_save(request: Request):
    global AREAS, ITEM_MAP
    data = await request.json()
    _check_edit_pin(data.get("pin"))
    new_zones = data.get("zones")
    if not isinstance(new_zones, list) or not new_zones:
        raise HTTPException(400, "ข้อมูลไม่ถูกต้อง")

    old_zones = {z["id"]: z for z in AREAS["zones"]}
    old_items = {it["id"]: it for z in AREAS["zones"] for it in z["items"]}
    used_ids = set(old_zones) | set(old_items)
    result, seen_zone = [], set()

    def new_id(prefix):
        while True:
            i = prefix + uuid.uuid4().hex[:6]
            if i not in used_ids:
                used_ids.add(i)
                return i

    for z in new_zones:
        name = str(z.get("name") or "").strip()[:60]
        if not name:
            raise HTTPException(400, "ชื่อพื้นที่ห้ามว่าง")
        zid = z.get("id") if z.get("id") in old_zones else new_id("Z")
        seen_zone.add(zid)
        items, seen_item = [], set()
        for it in z.get("items") or []:
            iname = str(it.get("name") or "").strip()[:200]
            if not iname:
                continue
            is_new = it.get("id") not in old_items
            iid = it.get("id") if not is_new else new_id("A")
            base = dict(old_items.get(iid, {}))
            base.pop("hidden", None)
            no = str(it.get("no") if it.get("no") is not None else "").strip()[:8]
            base.update({"id": iid, "no": int(no) if no.isdigit() else (no or "-"), "name": iname})
            proc = str(it.get("proc") or "").strip()
            if proc in process_tags() and not base.get("extra"):
                base["proc"] = proc  # ป้ายกระบวนการ เช่น จุดเตรียม / บรรจุ
            else:
                base.pop("proc", None)
            cust = str(it.get("cust") or "").strip()
            if (cust in CUSTOMERS or cust == BOTH_CUST) and not base.get("extra"):
                base["cust"] = cust  # Checklist ของลูกค้า · both = ทั้งสองลูกค้า (ไม่ระบุ = ไม่มีลูกค้า)
            else:
                base.pop("cust", None)
            if "pilot" in it:  # ★ หัวข้อทดลอง: แสดงใน Dashboard แต่ยังไม่นับใน %
                if it.get("pilot") and not base.get("extra"):
                    base["pilot"] = True
                else:
                    base.pop("pilot", None)
            if "qty" in it:  # ช่องนับจำนวน (เช่น วัตถุดิบ Delay แยกห้อง)
                q = clean_qty_cfg(it.get("qty"))
                if q and not base.get("extra"):
                    base["qty"] = q
                else:
                    base.pop("qty", None)
            if is_new:
                base["added"] = round_of(now()).isoformat()
            items.append(base)
            seen_item.add(iid)
        # รายการเดิมที่ถูกลบ → ซ่อนไว้ (ไม่หายจากข้อมูลย้อนหลัง)
        if zid in old_zones:
            for it in old_zones[zid]["items"]:
                if it["id"] not in seen_item:
                    items.append({**it, "hidden": True, "removed": it.get("removed") or round_of(now()).isoformat()})
        # ทุกพื้นที่มีหัวข้อเพิ่มเติมท้ายสุด
        active = [i for i in items if not i.get("hidden")]
        if any(not i.get("extra") for i in active) and not any(i.get("extra") for i in active):
            items.append({"id": new_id("X"), "no": "+", "name": "หัวข้อเพิ่มเติม (นอกเหนือ Checklist)",
                          "extra": True, "free": True, "added": round_of(now()).isoformat()})
        # ให้หัวข้อเพิ่มเติมอยู่ท้ายสุดเสมอ
        items.sort(key=lambda i: bool(i.get("extra")))
        zbase = {k: v for k, v in old_zones.get(zid, {}).items() if k not in ("hidden", "items")}
        zbase.update({"id": zid, "name": name, "owner": str(z.get("owner") or "").strip()[:80], "items": items})
        if zid not in old_zones and z.get("deck") in DECKS and z.get("deck") != "pc":
            zbase["deck"] = z["deck"]  # พื้นที่ใหม่ที่เพิ่มในกลุ่ม GMP
        sub = str(z.get("sub") or "").strip()[:30]  # กลุ่มย่อย เช่น WH → WH1, WH5 ... รวมเป็นการ์ดเดียว
        if sub:
            zbase["sub"] = sub
        else:
            zbase.pop("sub", None)
        result.append(zbase)
    for zid, z in old_zones.items():
        if zid not in seen_zone:
            result.append({**z, "hidden": True, "items": [
                {**it, "removed": it.get("removed") or round_of(now()).isoformat()} if not it.get("hidden") else it
                for it in z["items"]]})

    cfg = {**AREAS, "zones": result}
    backup_dir = DATA_DIR / "areas_backup"
    backup_dir.mkdir(exist_ok=True)
    stamp = now().strftime("%Y%m%d_%H%M%S")
    (backup_dir / f"areas_{stamp}.json").write_text(json.dumps(AREAS, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp = AREAS_PATH.with_suffix(".tmp")
    tmp.write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(AREAS_PATH)
    AREAS, ITEM_MAP = load_areas()
    return {"ok": True, "zones": visible_zones()}


@app.get("/t/{fn:path}")
def api_thumb(fn: str):
    """รูปย่อ: ถ้ายังไม่มี (รูปเก่าก่อนอัปเดต) จะสร้างให้ครั้งแรกแล้วเก็บไว้ใช้ต่อ"""
    base = UPLOAD_DIR.resolve()
    tp = thumb_path(fn).resolve()
    if base not in tp.parents:
        raise HTTPException(404)
    if not tp.exists() and not make_thumb(fn):
        raise HTTPException(404)
    return FileResponse(tp, media_type="image/jpeg")


def _dir_size(p: Path) -> int:
    return sum(f.stat().st_size for f in p.rglob("*") if f.is_file()) if p.exists() else 0


def _gb(n: int) -> float:
    return round(n / 1024 ** 3, 2)


@app.post("/api/admin/disk")
async def api_admin_disk(request: Request):
    """พื้นที่ดิสก์: ใช้ไปเท่าไหร่ เหลือเท่าไหร่ รูปเก่าสุดวันไหน (ต้องใช้รหัส EDIT_PIN)"""
    data = await request.json()
    _check_edit_pin(data.get("pin"))

    def calc():
        du = shutil.disk_usage(DATA_DIR)
        days = sorted(d.name for d in UPLOAD_DIR.iterdir() if d.is_dir() and re.match(r"^\d{4}-\d{2}-\d{2}$", d.name))
        return {
            "disk_total_gb": _gb(du.total), "disk_free_gb": _gb(du.free),
            "disk_used_pct": round(100 * (du.total - du.free) / du.total, 1),
            "uploads_gb": _gb(_dir_size(UPLOAD_DIR)), "slides_gb": _gb(_dir_size(SLIDE_DIR)),
            "upload_days": len(days), "oldest": days[0] if days else None,
            "archive_dir": str(ARCHIVE_DIR), "archived_gb": _gb(_dir_size(ARCHIVE_DIR)),
        }
    return await run_in_threadpool(calc)


@app.post("/api/admin/archive")
async def api_admin_archive(request: Request):
    """ย้ายรูป/วิดีโอที่เก่ากว่า N วันไปเก็บที่ ARCHIVE_DIR (สไลด์ของวันนั้นยังอยู่ครบ เพราะรูปฝังในไฟล์ pptx แล้ว)"""
    data = await request.json()
    _check_edit_pin(data.get("pin"))
    try:
        keep_days = int(data.get("days") or 0)
    except ValueError:
        keep_days = 0
    if keep_days < 30:
        raise HTTPException(400, "ต้องเก็บรูปในเครื่องไว้อย่างน้อย 30 วัน")
    limit = (round_of(now()) - timedelta(days=keep_days)).isoformat()

    def run():
        moved, freed = [], 0
        ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
        for d in sorted(UPLOAD_DIR.iterdir()):
            if not d.is_dir() or d.name >= limit:
                continue
            try:
                rd = date.fromisoformat(d.name)
            except ValueError:
                continue
            if not all(slide_path(rd, dk).exists() for dk in live_decks()):
                archive_round(rd)  # ทำสไลด์เก็บไว้ก่อนย้ายรูปออก
            size = _dir_size(d)
            dest = ARCHIVE_DIR / d.name
            if dest.exists():
                for f in d.rglob("*"):
                    if f.is_file():
                        t = dest / f.relative_to(d)
                        t.parent.mkdir(parents=True, exist_ok=True)
                        shutil.move(str(f), str(t))
                shutil.rmtree(d, ignore_errors=True)
            else:
                shutil.move(str(d), str(dest))
            moved.append(d.name)
            freed += size
        return {"ok": True, "moved_days": len(moved), "from": moved[0] if moved else None,
                "to": moved[-1] if moved else None, "freed_gb": _gb(freed), "archive_dir": str(ARCHIVE_DIR)}
    return await run_in_threadpool(run)


# ---------------------------------------------------------------- Retort Analysis (บันทึกปัญหา Retort + วิเคราะห์)
RETORT_CFG_PATH = DATA_DIR / "retort_config.json"
RETORT_PIN = os.environ.get("RETORT_PIN", "1111").strip()  # รหัสเข้าหน้า Retort Analysis
RT_TOKEN = hashlib.sha256(("retort:" + RETORT_PIN).encode()).hexdigest()[:32]
APPROVE_PIN = os.environ.get("APPROVE_PIN", "").strip()  # รหัสผู้อนุมัติ (ไม่ตั้ง = ไม่ต้องใช้รหัส)
_RP_APPROVAL = ("approvedBy", "approvedAt")  # ข้อมูลการอนุมัติ: เปลี่ยนได้เฉพาะ /approve เท่านั้น
RT_PHOTO_DIR = DATA_DIR / "retort_photos"      # รูปถ่ายปัญหา Retort (มีเวลาประทับจากหน้าเว็บ)
_RT_PHOTO_RE = re.compile(r"^[0-9]{8}_[0-9a-f]{16}\.jpg$")
_RT_NAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9 _./-]{0,19}$")


# บทบาท: หัวหน้า (Supervisor) เข้าด้วย EDIT_PIN → เห็นทุกแท็บ · พนักงานเข้าด้วย RETORT_PIN → บันทึก + รายการ
RT_SUP_TOKEN = hashlib.sha256(("retort-sup:" + EDIT_PIN).encode()).hexdigest()[:32]
RT_ROLE_NAME = {"sup": "หัวหน้า (Supervisor)", "staff": "พนักงาน"}


def _rt_role(request: Request) -> str:
    """sup / staff / '' (ยังไม่ได้เข้าระบบ) — จำไว้ใน cookie 400 วัน (เข้าครั้งเดียว)"""
    t = request.cookies.get("rt_access") or ""
    if EDIT_PIN and t == RT_SUP_TOKEN:
        return "sup"
    if t == RT_TOKEN or not RETORT_PIN:
        return "staff"
    return ""


def _rt_auth(request: Request):
    if not _rt_role(request):
        raise HTTPException(401, "RETORT_LOCKED")


def _rt_sup(request: Request, pin=None):
    """เฉพาะหัวหน้า: เข้าระบบด้วยรหัสหัวหน้าแล้ว หรือส่งรหัสแก้ไข (EDIT_PIN) มาด้วย"""
    _rt_auth(request)
    if _rt_role(request) != "sup":
        if pin is None:
            raise HTTPException(403, "เฉพาะหัวหน้า (Supervisor)")
        _check_edit_pin(pin)


@app.post("/api/retort/login")
async def api_retort_login(request: Request):
    data = await request.json()
    pin = str(data.get("pin") or "").strip()
    if pin and EDIT_PIN and pin == EDIT_PIN:
        role, tok = "sup", RT_SUP_TOKEN
    elif pin and pin == RETORT_PIN:
        role, tok = "staff", RT_TOKEN
    else:
        raise HTTPException(403, "รหัสไม่ถูกต้อง")
    resp = JSONResponse({"ok": True, "role": role, "role_name": RT_ROLE_NAME[role]})
    resp.set_cookie("rt_access", tok, max_age=60 * 60 * 24 * 400, httponly=True, samesite="lax")
    return resp


@app.post("/api/retort/logout")
def api_retort_logout():
    resp = JSONResponse({"ok": True})
    resp.delete_cookie("rt_access", samesite="lax")
    return resp
_RP_TEXT = {"time": 5, "shift": 4, "plant": 10, "program": 20, "programDoc": 20, "retort": 20, "product": 120, "batch": 60, "code": 60, "category": 40, "cause": 200,
            "causeOther": 200, "duration": 12, "detail": 3000, "action": 3000, "impact": 8, "by": 60, "updatedBy": 60}


def _clean_rp(d: dict) -> dict:
    """ตรวจ/ตัดข้อมูลปัญหา Retort ที่ส่งมาจากหน้าเว็บ"""
    if not isinstance(d, dict):
        raise HTTPException(400, "ข้อมูลไม่ถูกต้อง")
    rid = str(d.get("id") or "").strip()[:40]
    if not rid or not all(c.isalnum() or c in "-_" for c in rid):
        raise HTTPException(400, "รหัสรายการไม่ถูกต้อง")
    try:
        dt = date.fromisoformat(str(d.get("date") or ""))
    except ValueError:
        raise HTTPException(400, "วันที่ไม่ถูกต้อง")
    out = {"id": rid, "date": dt.isoformat()}
    for k, n in _RP_TEXT.items():
        out[k] = str(d.get(k) if d.get(k) is not None else "").strip()[:n]
    if not out["category"]:
        raise HTTPException(400, "กรุณาเลือกหัวข้อปัญหา")
    vals = d.get("values") if isinstance(d.get("values"), dict) else {}
    # ค่าที่วัดของหัวข้อปัญหา + ช่องรายละเอียดที่ผู้ดูแลเพิ่มเอง (key ขึ้นต้น x_)
    out["values"] = {str(k)[:40]: str(v)[:200] for k, v in list(vals.items())[:40]}
    ph = d.get("photos") if isinstance(d.get("photos"), list) else []
    out["photos"] = [str(f) for f in ph if isinstance(f, str) and _RT_PHOTO_RE.match(f)][:8]
    out["impact"] = "hold" if out["impact"] == "hold" else "none"
    out["shift"] = out["shift"] if out["shift"] in ("DS", "NS") else "DS"
    for k in ("createdAt", "updatedAt"):
        try:
            out[k] = int(d.get(k) or 0)
        except (TypeError, ValueError):
            out[k] = 0
    return out


@app.get("/api/retort")
def api_retort(request: Request, since: int = 0):
    _rt_auth(request)
    """ค่าตั้งต้น + รายการปัญหาทั้งหมด (หรือเฉพาะที่แก้ไขหลังเวลา since มิลลิวินาที)"""
    cfg = None
    if RETORT_CFG_PATH.exists():
        try:
            cfg = json.loads(RETORT_CFG_PATH.read_text(encoding="utf-8"))
        except ValueError:
            cfg = None
    with db() as con:
        rows = con.execute("SELECT data FROM retort_problems WHERE updated_at > ? ORDER BY date DESC", (since,)).fetchall()
        ids = [r["id"] for r in con.execute("SELECT id FROM retort_problems")]
    return {"config": cfg, "items": [json.loads(r["data"]) for r in rows], "ids": ids,
            "now": int(now().timestamp() * 1000), "role": _rt_role(request)}


@app.post("/api/retort/problems")
async def api_retort_save(request: Request):
    _rt_auth(request)
    x = _clean_rp(await request.json())
    x["updatedAt"] = int(now().timestamp() * 1000)
    x["createdAt"] = x["createdAt"] or x["updatedAt"]
    with db() as con:
        old = con.execute("SELECT data FROM retort_problems WHERE id=?", (x["id"],)).fetchone()
    if old:  # หน้าเว็บส่งชื่อผู้อนุมัติมาเองไม่ได้ → ใช้ค่าเดิมจากเซิร์ฟเวอร์
        od = json.loads(old["data"])
        for k in _RP_APPROVAL:
            if od.get(k):
                x[k] = od[k]
    with db() as con:
        con.execute("INSERT INTO retort_problems(id, date, data, updated_at) VALUES(?,?,?,?) "
                    "ON CONFLICT(id) DO UPDATE SET date=excluded.date, data=excluded.data, updated_at=excluded.updated_at",
                    (x["id"], x["date"], json.dumps(x, ensure_ascii=False), x["updatedAt"]))
    return {"ok": True, "item": x}


@app.delete("/api/retort/problems/{rid}")
def api_retort_delete(rid: str, request: Request, pin: str = ""):
    _rt_auth(request)
    if _rt_role(request) != "sup":  # พนักงานลบรายการไม่ได้ (แก้ไขได้)
        raise HTTPException(403, "ลบรายการได้เฉพาะหัวหน้า (Supervisor)")
    if ADMIN_PIN and pin != ADMIN_PIN:
        raise HTTPException(403, "PIN ไม่ถูกต้อง")
    with db() as con:
        con.execute("DELETE FROM retort_problems WHERE id=?", (rid,))
    return {"ok": True}


@app.post("/api/retort/problems/{rid}/approve")
async def api_retort_approve(rid: str, request: Request):
    """อนุมัติ / ยกเลิกอนุมัติรายการปัญหา Retort: {pin, name, undo} — เฉพาะหัวหน้า"""
    _rt_auth(request)
    data = await request.json()
    if _rt_role(request) != "sup":
        raise HTTPException(403, "อนุมัติได้เฉพาะหัวหน้า (Supervisor)")
    if APPROVE_PIN and str(data.get("pin") or "").strip() != APPROVE_PIN:
        raise HTTPException(403, "รหัสผู้อนุมัติไม่ถูกต้อง")
    undo = bool(data.get("undo"))
    name = str(data.get("name") or "").strip()[:60]
    if not undo and not name:
        raise HTTPException(400, "กรุณาใส่ชื่อผู้อนุมัติ")
    ts = int(now().timestamp() * 1000)
    with db() as con:
        row = con.execute("SELECT data FROM retort_problems WHERE id=?", (rid,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบรายการ")
        x = json.loads(row["data"])
        if undo:
            x.pop("approvedBy", None)
            x.pop("approvedAt", None)
        else:
            x["approvedBy"], x["approvedAt"] = name, ts
        x["updatedAt"] = ts
        con.execute("UPDATE retort_problems SET data=?, updated_at=? WHERE id=?",
                    (json.dumps(x, ensure_ascii=False), ts, rid))
    return {"ok": True, "item": x}


@app.post("/api/retort/photo")
async def api_retort_photo(request: Request, file: UploadFile = File(...)):
    """รับรูปถ่ายปัญหา (JPEG ที่ประทับเวลาแล้วจากหน้าเว็บ)"""
    _rt_auth(request)
    data = await file.read()
    if len(data) > 8 * 1024 * 1024:
        raise HTTPException(413, "รูปใหญ่เกินไป")
    if data[:3] != b"\xff\xd8\xff":
        raise HTTPException(400, "รองรับเฉพาะรูป JPEG")
    RT_PHOTO_DIR.mkdir(parents=True, exist_ok=True)
    fn = f"{now().strftime('%Y%m%d')}_{secrets.token_hex(8)}.jpg"
    (RT_PHOTO_DIR / fn).write_bytes(data)
    return {"ok": True, "file": fn}


RT_THUMB_W = 480  # รูปย่อสำหรับตาราง/รายการ (รูปเต็ม 1600px ใช้ตอนกดดูรูปใหญ่เท่านั้น)


def _rt_thumb(src: Path) -> Path:
    tp = RT_PHOTO_DIR / "t" / src.name
    if not tp.exists():
        tp.parent.mkdir(exist_ok=True)
        try:
            with Image.open(src) as im:
                im = ImageOps.exif_transpose(im).convert("RGB")
                im.thumbnail((RT_THUMB_W, RT_THUMB_W))
                tmp = tp.with_suffix(f".{secrets.token_hex(4)}.tmp")
                im.save(tmp, "JPEG", quality=78, optimize=True)
                tmp.replace(tp)
        except Exception:  # รูปเสีย/อ่านไม่ได้ → ส่งรูปเต็มแทน
            return src
    return tp


@app.get("/api/retort/photo/{fn}")
def api_retort_photo_get(fn: str, request: Request, w: int = 0):
    _rt_auth(request)
    if not _RT_PHOTO_RE.match(fn) or not (RT_PHOTO_DIR / fn).exists():
        raise HTTPException(404, "ไม่พบรูป")
    src = RT_PHOTO_DIR / fn
    return FileResponse(_rt_thumb(src) if w else src, media_type="image/jpeg",
                        headers={"Cache-Control": "private, max-age=31536000, immutable"})


def _read_rt_cfg():
    if RETORT_CFG_PATH.exists():
        try:
            c = json.loads(RETORT_CFG_PATH.read_text(encoding="utf-8"))
            if isinstance(c, dict):
                return c
        except ValueError:
            pass
    return {}


def _clean_cats(base) -> list:
    """ตรวจรายการหัวข้อปัญหาที่หน้าเว็บส่งมา (ใช้เมื่อไฟล์ตั้งค่าบนเซิร์ฟเวอร์ยังเป็นรุ่นเก่า)"""
    out = []
    for c in (base if isinstance(base, list) else [])[:200]:
        if not isinstance(c, dict):
            continue
        cid, name = str(c.get("id") or "").strip()[:40], str(c.get("name") or "").strip()[:120]
        if not cid or not name:
            continue
        fields = [f for f in (c.get("fields") or []) if isinstance(f, dict)][:10]
        out.append({"id": cid, "grp": str(c.get("grp") or "other")[:12], "name": name, "active": c.get("active") is not False,
                    "causes": [str(x).strip()[:200] for x in (c.get("causes") or []) if str(x).strip()][:60],
                    "fields": [{k: str(f.get(k) or "")[:40] for k in ("key", "label", "unit")} for f in fields]})
    return out


@app.post("/api/retort/add")
async def api_retort_add(request: Request):
    """หน้างานเพิ่มตัวเลือกเองได้ (ไม่ต้องใช้รหัสแก้ไข): พื้นที่ / หม้อ / หัวข้อปัญหา / สาเหตุ"""
    _rt_auth(request)
    data = await request.json()
    kind = str(data.get("kind") or "")
    name = " ".join(str(data.get("name") or "").split())
    cfg = _read_rt_cfg()
    plants = cfg.get("plants") if isinstance(cfg.get("plants"), list) and cfg.get("plants") else ["PF1", "PF2"]
    if kind in ("plant", "retort"):
        name = name.upper()
        if not _RT_NAME_RE.match(name):
            raise HTTPException(400, "ชื่อไม่ถูกต้อง (ใช้ตัวอักษร/ตัวเลข ไม่เกิน 20 ตัว)")
    elif not 0 < len(name) <= 120:
        raise HTTPException(400, "กรอกชื่อไม่เกิน 120 ตัวอักษร")
    if kind == "plant":
        if name not in plants:
            plants.append(name)
        cfg["plants"] = plants[:20]
    elif kind == "retort":
        plant = str(data.get("plant") or "").strip()
        if plant not in plants:
            raise HTTPException(400, "พื้นที่ไม่ถูกต้อง")
        pr = cfg.get("plantRetorts") if isinstance(cfg.get("plantRetorts"), dict) else {}
        cur = pr.get(plant) if isinstance(pr.get(plant), list) and pr.get(plant) else \
            [str(x).strip().upper() for x in (data.get("base") or []) if _RT_NAME_RE.match(str(x).strip())][:80]
        if name not in cur:
            cur.append(name)
        pr[plant] = cur
        cfg["plantRetorts"] = pr
    elif kind in ("category", "cause"):
        cats = cfg.get("categories")
        if (cfg.get("version") or 1) < 3 or not isinstance(cats, list) or not cats:
            cats = _clean_cats(data.get("base"))
            cfg["version"] = 3
            if not cats:
                raise HTTPException(400, "ข้อมูลหัวข้อไม่ถูกต้อง")
        if kind == "category":
            hit = next((c for c in cats if str(c.get("name", "")).strip().lower() == name.lower()), None)
            if hit:
                hit["active"] = True
                hit.pop("deleted", None)  # เพิ่มชื่อเดิมที่เคยลบ = กู้คืน
            else:
                cats.append({"id": "c" + secrets.token_hex(5), "grp": "other", "name": name, "active": True, "causes": [], "fields": []})
        else:
            cat = next((c for c in cats if c.get("id") == str(data.get("category") or "")), None)
            if not cat:
                raise HTTPException(400, "ไม่พบหัวข้อปัญหา")
            causes = cat.setdefault("causes", [])
            if name.lower() not in [str(x).lower() for x in causes]:
                causes.append(name)
        cfg["categories"] = cats
    else:
        raise HTTPException(400, "ประเภทไม่ถูกต้อง")
    cfg["updatedAt"] = int(now().timestamp() * 1000)
    tmp = RETORT_CFG_PATH.with_suffix(".tmp")
    tmp.write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(RETORT_CFG_PATH)
    return {"ok": True, "config": cfg}


# หม้อ PF1 ชุดใหม่: RP01–RP23, RA01–RA04, RS01–RS02 + RS04–RS09, RT01–RT07 + RT09–RT10
PF1_RETORTS = ([f"RP{i:02d}" for i in range(1, 24)] + [f"RA{i:02d}" for i in range(1, 5)]
               + [f"RS{i:02d}" for i in range(1, 10) if i != 3] + [f"RT{i:02d}" for i in range(1, 11) if i != 8])


def _migrate_pf1_retorts():
    """ครั้งเดียว: ค่าตั้ง Retort ที่บันทึกหม้อ PF1 ไว้แล้ว → ใช้หม้อชุดใหม่ (รวมหม้อที่เคยเพิ่มเองไว้ด้วย)
    ถ้ายังไม่เคยบันทึกรายการหม้อ PF1 หน้าเว็บจะใช้ชุดใหม่เป็นค่าเริ่มต้นอยู่แล้ว"""
    cfg = _read_rt_cfg()
    if not cfg or cfg.get("pf1Retorts2"):
        return
    pr = cfg.get("plantRetorts") if isinstance(cfg.get("plantRetorts"), dict) else {}
    old = pr.get("PF1") if isinstance(pr.get("PF1"), list) else []
    if old:
        pr["PF1"] = list(dict.fromkeys(PF1_RETORTS + [str(x) for x in old]))
        cfg["plantRetorts"] = pr
    cfg["pf1Retorts2"] = True
    _write_rt_cfg(cfg)
    print(f"[retort] หม้อ PF1 ชุดใหม่ (RP/RA/RS/RT) · {len(pr.get('PF1') or PF1_RETORTS)} หม้อ", flush=True)


def _write_rt_cfg(cfg: dict):
    cfg["updatedAt"] = int(now().timestamp() * 1000)
    if RETORT_CFG_PATH.exists():  # สำรองไฟล์เดิมก่อนเขียนทับ
        bdir = DATA_DIR / "areas_backup"
        bdir.mkdir(exist_ok=True)
        (bdir / f"retort_config_{now().strftime('%Y%m%d_%H%M%S')}.json").write_text(
            RETORT_CFG_PATH.read_text(encoding="utf-8"), encoding="utf-8")
    tmp = RETORT_CFG_PATH.with_suffix(".tmp")
    tmp.write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(RETORT_CFG_PATH)


@app.post("/api/retort/remove")
async def api_retort_remove(request: Request):
    """ลบพื้นที่ / หม้อ / หัวข้อปัญหา ออกจากรายการ (ต้องใส่รหัสแก้ไข EDIT_PIN)
    ข้อมูลที่บันทึกไว้แล้วยังอยู่ครบ: หัวข้อปัญหาถูกซ่อน (deleted) เพื่อให้ประวัติเก่ายังแสดงชื่อได้"""
    data = await request.json()
    _rt_sup(request, data.get("pin"))
    kind = str(data.get("kind") or "")
    name = " ".join(str(data.get("name") or "").split())
    cfg = _read_rt_cfg()
    plants = cfg.get("plants") if isinstance(cfg.get("plants"), list) and cfg.get("plants") else ["PF1", "PF2"]
    pr = cfg.get("plantRetorts") if isinstance(cfg.get("plantRetorts"), dict) else {}
    if kind == "plant":
        if name not in plants:
            raise HTTPException(404, "ไม่พบพื้นที่นี้")
        if len(plants) <= 1:
            raise HTTPException(400, "ต้องมีอย่างน้อย 1 พื้นที่")
        cfg["plants"] = [x for x in plants if x != name]
        pr.pop(name, None)
        cfg["plantRetorts"] = pr
    elif kind == "retort":
        plant = str(data.get("plant") or "").strip()
        if plant not in plants:
            raise HTTPException(400, "พื้นที่ไม่ถูกต้อง")
        cur = pr.get(plant) if isinstance(pr.get(plant), list) else \
            [str(x).strip().upper() for x in (data.get("base") or []) if _RT_NAME_RE.match(str(x).strip())][:80]
        if name not in cur:
            raise HTTPException(404, "ไม่พบหม้อนี้")
        if len(cur) <= 1:
            raise HTTPException(400, "ต้องมีอย่างน้อย 1 หม้อในพื้นที่นี้")
        pr[plant] = [x for x in cur if x != name]
        cfg["plantRetorts"] = pr
    elif kind == "category":
        cats = cfg.get("categories")
        if (cfg.get("version") or 1) < 3 or not isinstance(cats, list) or not cats:
            cats = _clean_cats(data.get("base"))
            cfg["version"] = 3
        cat = next((c for c in cats if c.get("id") == str(data.get("id") or "")), None)
        if not cat:
            raise HTTPException(404, "ไม่พบหัวข้อปัญหา")
        cat["active"] = False
        cat["deleted"] = True
        cfg["categories"] = cats
    else:
        raise HTTPException(400, "ประเภทไม่ถูกต้อง")
    _write_rt_cfg(cfg)
    return {"ok": True, "config": cfg}


# ---------------------------------------------------------------- Retort: อ่านค่าโปรแกรมฆ่าเชื้อจาก PDF (TP002-*)
# หัวหน้าอัปโหลด PDF จาก ISO Master → อ่านตาราง Come up / Cooking / Cooling → หน้าเว็บให้ตรวจ/แก้ก่อนกดยืนยัน
# อ่านได้เฉพาะ PDF ที่เป็นตัวอักษร (ไม่ใช่ภาพสแกน) · ต้องติดตั้ง pdfplumber (อยู่ใน requirements.txt)
_TP_DOC_RE = re.compile(r"\b(TP\d{3}-[A-Z0-9]{1,4})\b")
_TP_PROG_RE = re.compile(r"(?:program|prog\.?|p/?g\.?|โปรแกรม)\s*(?:no\.?|number|ที่)?\s*[:#.]?\s*(\d{1,3})\b", re.I)
_TP_REV_RE = re.compile(r"\bRev(?:ision)?\.?\s*(?:No\.?)?\s*[:.]?\s*(\d{1,3})\b", re.I)
_TP_EFF_RE = re.compile(r"(?:effective|eff\.|มีผล|บังคับใช้|ประกาศใช้)[^0-9]{0,30}(\d{1,2}\s*[/.-]\s*\d{1,2}\s*[/.-]\s*\d{2,4})", re.I)


def _tp_phase(t: str) -> str:
    t = (t or "").lower()
    if re.search(r"come\s*-?\s*up|heat(?:ing)?\s*-?\s*up|venting|ไล่อากาศ|ขึ้นอุณหภูมิ", t):
        return "come"
    if re.search(r"cool|หล่อเย็น|ลดอุณหภูมิ", t):
        return "cool"
    if re.search(r"cook|steril|ฆ่าเชื้อ", t):
        return "cook"
    return ""


def _tp_num(c):
    c = (c or "").strip().replace(",", ".")
    if re.fullmatch(r"[xX]{1,3}", c):
        return "XX"
    return c if re.fullmatch(r"-?\d+(?:\.\d+)?", c) else None


def _tp_rows(table):
    """แตกแถวที่ pdfplumber รวมไว้ในช่องเดียว (หลายบรรทัด) ให้เป็นทีละแถว"""
    out = []
    for row in table:
        cells = [str(c or "").strip() for c in row]
        parts = [c.split("\n") for c in cells]
        k = max([len(p) for p, c in zip(parts, cells) if _tp_num(c.split("\n")[0]) is not None] or [1])
        if k > 1 and sum(1 for p in parts if len(p) == k) >= 2:
            for i in range(k):
                out.append([p[i] if len(p) == k else (p[0] if i == 0 else "") for p in parts])
        else:
            out.append([c.replace("\n", " ") for c in cells])
    return out


def _tp_table_steps(table):
    cols = {}
    steps, phase = [], ""
    for cells in _tp_rows(table):
        low = [c.lower() for c in cells]
        hdr = {}
        for i, c in enumerate(low):
            if re.search(r"temp|อุณหภูมิ|°\s*c", c) and "temp" not in hdr:
                hdr["temp"] = i
            elif re.search(r"hold", c):
                hdr["hold"] = i
            elif re.search(r"press|bar|ความดัน", c) and "press" not in hdr:
                hdr["press"] = i
            elif re.search(r"time|เวลา|min|นาที", c) and "time" not in hdr:
                hdr["time"] = i
        if len(hdr) >= 2 and not any(_tp_num(c) for c in cells):
            cols = hdr
            continue
        ph = next((_tp_phase(c) for c in cells if _tp_num(c) is None and _tp_phase(c)), "")
        phase = ph or phase
        if cols and {"temp", "time", "press"} <= set(cols):
            g = lambda k: _tp_num(cells[cols[k]]) if k in cols and cols[k] < len(cells) else None
            vals = [g("temp"), g("time"), g("press")]
            hold = g("hold")
        else:
            nums = [n for n in (_tp_num(c) for c in cells) if n is not None]
            if len(nums) >= 2 and re.fullmatch(r"\d{1,2}", nums[0]) and len(nums) >= 4 and not ph:
                nums = nums[1:]  # คอลัมน์ลำดับขั้น
            vals, hold = nums[:3], (nums[3] if len(nums) > 3 else None)
            vals += [None] * (3 - len(vals))
        t, tm, pr = vals
        try:
            ok = t not in (None, "XX") and 20 <= float(t) <= 150 and pr not in (None, "XX") and 0 <= float(pr) <= 6 and tm is not None
        except ValueError:
            ok = False
        if not ok:
            continue
        st = [phase or ("cook" if steps and steps[-1][0] == "cook" else "come"), t, tm, pr]
        if hold is not None:
            st.append(hold)
        steps.append(st)
    return steps


def _tp_parse_pdf(raw: bytes, filename: str) -> dict:
    try:
        import pdfplumber
    except ImportError:
        raise HTTPException(501, "เซิร์ฟเวอร์ยังไม่มีตัวอ่าน PDF · รันคำสั่ง pip install pdfplumber (หรือ start_windows.bat ใหม่) แล้วรีสตาร์ท")
    progs, order, text_all, n_tables = {}, [], "", 0
    last_no = None
    try:
        pdf = pdfplumber.open(io.BytesIO(raw))
    except Exception:
        raise HTTPException(400, "เปิดไฟล์ PDF นี้ไม่ได้")
    with pdf:
        for page in pdf.pages[:60]:
            text_all += (page.extract_text() or "") + "\n"
            for tb in page.find_tables():
                n_tables += 1
                table = tb.extract()
                steps = _tp_table_steps(table)
                if not steps:
                    continue
                top = max(0, tb.bbox[1] - 140)
                above = (page.crop((0, top, page.width, tb.bbox[1])).extract_text() or "") if tb.bbox[1] > 1 else ""
                inside = " ".join(str(c or "") for r in table[:3] for c in r)
                m = list(_TP_PROG_RE.finditer(inside)) or list(_TP_PROG_RE.finditer(above))
                no = m[-1].group(1) if m else last_no
                no = str(int(no)) if no and no.isdigit() else (no or f"?{len(order) + 1}")
                note = ""
                if above:
                    ln = [x.strip() for x in above.split("\n") if x.strip()]
                    note = next((x for x in reversed(ln) if _TP_PROG_RE.search(x)), "")[:120]
                if no not in progs:
                    progs[no] = {"no": no, "note": note, "steps": []}
                    order.append(no)
                progs[no]["steps"] += steps
                last_no = no if m else last_no
    if len(text_all.strip()) < 40:
        raise HTTPException(422, "PDF นี้ไม่มีตัวอักษรให้อ่าน (น่าจะเป็นไฟล์สแกนเป็นรูป) · ใส่ค่าเองได้ที่ปุ่ม แก้ไขค่าเอง")
    fname = filename or ""
    doc = (_TP_DOC_RE.search(fname.upper()) or _TP_DOC_RE.search(text_all) or [None, ""])[1]
    rev = (re.search(r"\bRe(?:v)?\.?\s*(\d{1,3})", fname, re.I) or _TP_REV_RE.search(text_all) or [None, ""])[1]
    eff = re.sub(r"\s", "", (_TP_EFF_RE.search(text_all) or [None, ""])[1])
    return {"doc": doc, "rev": rev, "eff": eff, "progs": [progs[k] for k in order],
            "pages": len(text_all.split("\f")), "tables": n_tables, "chars": len(text_all)}


@app.post("/api/retort/tp/parse")
async def api_retort_tp_parse(request: Request, file: UploadFile = File(...), pin: str = Form("")):
    _rt_sup(request, pin or None)
    raw = await file.read()
    if len(raw) > 30 * 1024 * 1024:
        raise HTTPException(400, "ไฟล์ใหญ่เกิน 30 MB")
    return await run_in_threadpool(_tp_parse_pdf, raw, file.filename or "")


@app.post("/api/retort/config")
async def api_retort_config(request: Request):
    data = await request.json()
    _rt_sup(request, data.get("pin"))
    cfg = data.get("config")
    if not isinstance(cfg, dict) or not isinstance(cfg.get("categories"), list):
        raise HTTPException(400, "ข้อมูลไม่ถูกต้อง")
    cfg["updatedAt"] = int(now().timestamp() * 1000)
    if RETORT_CFG_PATH.exists():  # สำรองไฟล์เดิมก่อนเขียนทับ
        bdir = DATA_DIR / "areas_backup"
        bdir.mkdir(exist_ok=True)
        (bdir / f"retort_config_{now().strftime('%Y%m%d_%H%M%S')}.json").write_text(
            RETORT_CFG_PATH.read_text(encoding="utf-8"), encoding="utf-8")
    tmp = RETORT_CFG_PATH.with_suffix(".tmp")
    tmp.write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(RETORT_CFG_PATH)
    return {"ok": True, "config": cfg}


# ---------------------------------------------------------------- Retort Analysis: ผู้ช่วย AI (Google Gemini API)
GEMINI_KEY_PATH = DATA_DIR / "gemini_key.txt"
GEMINI_MODELS = [m for m in [os.environ.get("GEMINI_MODEL", "").strip(),
                             "gemini-3.8-flash", "gemini-3.5-flash-lite", "gemini-2.5-flash", "gemini-flash-latest"] if m]
_gemini_ok_model = {"m": ""}  # จำรุ่นที่เรียกสำเร็จล่าสุด


def gemini_key() -> str:
    k = os.environ.get("GEMINI_API_KEY", "").strip()
    if not k and GEMINI_KEY_PATH.exists():
        k = GEMINI_KEY_PATH.read_text(encoding="utf-8").strip()
    return k


RETORT_AI_SYSTEM = (
    "คุณคือผู้ช่วยวิเคราะห์ปัญหากระบวนการฆ่าเชื้อ (Retort / Thermal process) ของโรงงานอาหารสัตว์เลี้ยงบรรจุกระป๋อง/เพาช์ "
    "ทำงานให้ทีม QC/Process Control ตอบเป็นภาษาไทย กระชับ เป็นข้อ ๆ อ่านง่าย ใช้หัวข้อสั้น ๆ และตัวหนาเฉพาะจุดสำคัญ "
    "วิเคราะห์จากข้อมูลที่ให้มาเท่านั้น อ้างอิงตัวเลข (จำนวนครั้ง, หม้อ, โปรแกรม, กะ, ช่วงวันที่) ทุกครั้งที่สรุป "
    "ถ้าข้อมูลไม่พอให้บอกตรง ๆ ว่าขาดข้อมูลอะไร ห้ามแต่งตัวเลขหรือเหตุการณ์ขึ้นเอง "
    "เมื่อพูดถึงค่าควบคุมของโปรแกรมฆ่าเชื้อ ให้ยึดตามตารางโปรแกรมที่แนบมา และเตือนให้ยืนยันกับเอกสารต้นฉบับ/ผู้รับผิดชอบ Thermal process "
    "เรื่องความปลอดภัยอาหาร (under-process, deviation, การ Hold สินค้า) ให้แนะนำตามหลัก GMP/HACCP และขั้นตอนของโรงงาน ไม่ตัดสินปล่อยสินค้าแทนคน"
)


def _gemini_call(key: str, body: dict) -> dict:
    import urllib.request
    import urllib.error
    models = ([_gemini_ok_model["m"]] if _gemini_ok_model["m"] else []) + [m for m in GEMINI_MODELS if m != _gemini_ok_model["m"]]
    busy = False  # มีรุ่นที่ตอบว่าคนใช้เยอะ
    quota = False  # มีรุ่นที่ตอบว่าใช้เกินโควตา (429)
    for m in models:
        for attempt in range(2):  # รุ่นเดิมลองซ้ำ 1 ครั้งถ้าระบบ AI ยุ่งชั่วคราว
            req = urllib.request.Request(
                f"https://generativelanguage.googleapis.com/v1beta/models/{m}:generateContent",
                data=json.dumps(body).encode("utf-8"),
                headers={"Content-Type": "application/json", "x-goog-api-key": key}, method="POST")
            try:
                with urllib.request.urlopen(req, timeout=90) as r:
                    data = json.loads(r.read().decode("utf-8"))
                _gemini_ok_model["m"] = m
                return {"model": m, "data": data}
            except urllib.error.HTTPError as e:
                try:
                    msg = json.loads(e.read().decode("utf-8")).get("error", {}).get("message", "")
                except Exception:  # noqa
                    msg = ""
                if e.code == 404:  # ไม่มีรุ่นนี้ → ลองรุ่นถัดไป
                    break
                if e.code in (400, 401, 403) and "API key" in msg:
                    raise HTTPException(400, "API Key ของ AI ไม่ถูกต้อง หรือยังไม่เปิดใช้งาน")
                print(f"[ai] {m} -> HTTP {e.code} {msg[:160]}", flush=True)
                if e.code in (429, 500, 502, 503, 504):  # ยุ่ง/เกินโควตาชั่วคราว → รอแล้วลองซ้ำ แล้วค่อยเปลี่ยนรุ่น
                    if e.code == 429:
                        quota = True
                    else:
                        busy = True
                    if attempt == 0 and e.code != 429:
                        _time.sleep(2)
                        continue
                    break
                raise HTTPException(502, f"AI ตอบกลับผิดพลาด (รหัส {e.code}) ลองใหม่อีกครั้ง")
            except urllib.error.URLError as e:
                print(f"[ai] connect error: {e.reason}", flush=True)
                raise HTTPException(502, "เชื่อมต่อ AI ไม่ได้ ตรวจสอบอินเทอร์เน็ตของเครื่อง server")
    if quota and not busy:
        raise HTTPException(429, "ใช้ AI ครบโควตาแล้ว (โควตาฟรีรีเซ็ตเป็นรายนาที/รายวัน) ลองใหม่ภายหลัง")
    if busy or quota:
        raise HTTPException(503, "ตอนนี้ระบบ AI มีผู้ใช้งานมาก ลองใหม่อีกสักครู่")
    raise HTTPException(502, "ไม่พบรุ่น AI ที่ใช้ได้ กรุณาแจ้งผู้ดูแลระบบ")


@app.get("/api/retort/ai/status")
def api_retort_ai_status(request: Request):
    _rt_auth(request)
    return {"configured": bool(gemini_key()),
            "from_env": bool(os.environ.get("GEMINI_API_KEY", "").strip())}


@app.post("/api/retort/ai/key")
async def api_retort_ai_key(request: Request):
    data = await request.json()
    _rt_sup(request, data.get("pin"))
    key = str(data.get("key") or "").strip()
    if key:
        if len(key) < 20 or any(c.isspace() for c in key):
            raise HTTPException(400, "รูปแบบ API Key ไม่ถูกต้อง")
        GEMINI_KEY_PATH.write_text(key, encoding="utf-8")
    elif GEMINI_KEY_PATH.exists():
        GEMINI_KEY_PATH.unlink()
    return {"ok": True, "configured": bool(gemini_key())}


@app.post("/api/retort/ai")
async def api_retort_ai(request: Request):
    """ถาม AI: ส่งข้อมูลปัญหา Retort ที่สรุปจากหน้าเว็บ (context) + ประวัติแชท (messages) ไปให้ Gemini — เฉพาะหัวหน้า"""
    _rt_sup(request)
    key = gemini_key()
    if not key:
        raise HTTPException(400, "ยังไม่ได้ตั้งค่า API Key ของ AI (ตั้งได้ที่หน้า ตั้งค่า)")
    data = await request.json()
    ctx = str(data.get("context") or "")[:150000]
    msgs = [m for m in (data.get("messages") or []) if isinstance(m, dict) and str(m.get("text") or "").strip()][-12:]
    if not msgs:
        raise HTTPException(400, "กรุณาพิมพ์คำถาม")
    contents = [{"role": "user", "parts": [{"text": "ข้อมูลสำหรับวิเคราะห์ (สรุปจากระบบ Retort Analysis):\n\n" + ctx}]},
                {"role": "model", "parts": [{"text": "รับทราบข้อมูลแล้ว พร้อมวิเคราะห์ครับ"}]}]
    for m in msgs:
        contents.append({"role": "model" if m.get("role") == "ai" else "user", "parts": [{"text": str(m["text"])[:4000]}]})
    body = {"systemInstruction": {"parts": [{"text": RETORT_AI_SYSTEM}]}, "contents": contents,
            "generationConfig": {"temperature": 0.4, "maxOutputTokens": 4096}}
    r = await run_in_threadpool(_gemini_call, key, body)
    cands = r["data"].get("candidates") or []
    text = "".join(p.get("text", "") for c in cands[:1] for p in (c.get("content") or {}).get("parts", []) if not p.get("thought"))
    if not text.strip():
        reason = (cands[0].get("finishReason") if cands else "") or (r["data"].get("promptFeedback") or {}).get("blockReason", "")
        raise HTTPException(502, f"AI ไม่ได้ตอบกลับ ({reason or 'ไม่ทราบสาเหตุ'})")
    return {"text": text.strip(), "model": "AI"}  # ไม่ส่งชื่อรุ่นไปที่หน้าเว็บ


# ---------------------------------------------------------------- เรื่องที่ต้องติดตาม (Follow-up)
# เพิ่ม/แก้ไข/ปิดเรื่องได้จากหน้าเว็บ · ลบต้องใส่รหัสแก้ไข (EDIT_PIN) · รูปเก็บที่ uploads/followups/
FU_STATUS = {"open": "รอดำเนินการ", "doing": "กำลังดำเนินการ", "done": "เสร็จแล้ว"}
FU_PHOTO_RE = re.compile(r"^followups/\d{4}-\d{2}/[A-Za-z0-9_-]+\.jpg$")
FU_MAX_PHOTOS = 12
FU_ISSUE_DAYS = 14


def _fu_ms() -> int:
    return int(now().timestamp() * 1000)


def _fu_txt(v, n: int) -> str:
    return _clean(str(v or ""))[:n]


def _fu_sub_info(con, sid: str):
    """ผลตรวจที่เป็นต้นเรื่อง: วันที่ ผู้ตรวจ หัวข้อ สิ่งที่พบ และรูปย่อ"""
    r = con.execute("SELECT * FROM submissions WHERE id=?", (sid,)).fetchone()
    if not r:
        return None
    r = dict(r)
    it = ITEM_MAP.get(r["area_id"], {})
    imgs = [x["filename"] for x in con.execute(
        "SELECT filename FROM images WHERE submission_id=? ORDER BY page, seq LIMIT 8", (sid,))]
    return {"sid": sid, "round": r["round_date"], "area_id": r["area_id"], "no": it.get("no", ""),
            "name": it.get("name", r["area_id"]), "zone": it.get("zone", ""), "zone_id": it.get("zone_id", ""),
            "inspector": r["inspector"], "at": r["created_at"], "note": r.get("note") or "",
            "cause": r.get("cause") or "", "action": r.get("action") or "",
            "media": [media_info(f) for f in imgs]}


@app.get("/api/followups")
def api_followups():
    """รายการเรื่องที่ต้องติดตาม + ปัญหาที่พบล่าสุดที่ยังไม่ได้เปิดเรื่องติดตาม"""
    with db() as con:
        items = [json.loads(r["data"]) for r in con.execute("SELECT data FROM followups ORDER BY created_at DESC")]
        linked = {(x.get("src") or {}).get("sid") for x in items} - {None, ""}
        for x in items:
            sid = (x.get("src") or {}).get("sid")
            x["src_info"] = _fu_sub_info(con, sid) if sid else None
        since = (round_of(now()) - timedelta(days=FU_ISSUE_DAYS)).isoformat()
        rows = [dict(r) for r in con.execute(
            "SELECT * FROM submissions WHERE status='พบปัญหา' AND round_date>=? ORDER BY created_at DESC LIMIT 300", (since,))]
        thumbs = {}
        for r in rows:
            f = con.execute("SELECT filename FROM images WHERE submission_id=? ORDER BY page, seq LIMIT 1", (r["id"],)).fetchone()
            thumbs[r["id"]] = media_info(f["filename"]) if f else None
    zdeck = {z["id"]: zone_deck(z) for z in AREAS["zones"]}
    issues = []
    for r in rows:
        if r["id"] in linked:
            continue
        it = ITEM_MAP.get(r["area_id"], {})
        issues.append({"sid": r["id"], "round": r["round_date"], "area_id": r["area_id"], "no": it.get("no", ""),
                       "name": it.get("name", r["area_id"]), "zone": it.get("zone", ""), "zone_id": it.get("zone_id", ""),
                       "deck": zdeck.get(it.get("zone_id"), "pc"), "inspector": r["inspector"], "at": r["created_at"],
                       "note": r.get("note") or "", "cause": r.get("cause") or "", "action": r.get("action") or "",
                       "thumb": thumbs.get(r["id"])})
    return {"items": items, "issues": issues[:120], "issue_days": FU_ISSUE_DAYS, "status": FU_STATUS}


@app.post("/api/followups")
async def api_followup_save(request: Request):
    """เพิ่มเรื่องใหม่ (ไม่มี id) หรือแก้ไขเรื่องเดิม · เปลี่ยนสถานะจะบันทึกลงประวัติให้อัตโนมัติ"""
    data = await request.json()
    by = _fu_txt(data.get("by"), 60)
    if not by:
        raise HTTPException(400, "ตั้งชื่อผู้ใช้ก่อน")
    title = _fu_txt(data.get("title"), 200)
    if not title:
        raise HTTPException(400, "กรอกเรื่องที่ต้องติดตาม")
    status = data.get("status") if data.get("status") in FU_STATUS else "open"
    due = str(data.get("due") or "")[:10]
    if due:
        try:
            date.fromisoformat(due)
        except ValueError:
            raise HTTPException(400, "วันที่กำหนดเสร็จไม่ถูกต้อง")
    zone_id = str(data.get("zone_id") or "")
    zone = next((z for z in AREAS["zones"] if z["id"] == zone_id), None)
    photos = []
    for f in data.get("photos") or []:
        f = str(f)
        if FU_PHOTO_RE.match(f) and (UPLOAD_DIR / f).exists() and f not in photos:
            photos.append(f)
    ts = _fu_ms()
    fid = str(data.get("id") or "")
    with db() as con:
        old = con.execute("SELECT data FROM followups WHERE id=?", (fid,)).fetchone() if fid else None
        old = json.loads(old["data"]) if old else None
        if fid and not old:
            raise HTTPException(404, "ไม่พบเรื่องนี้ (อาจถูกลบไปแล้ว)")
        x = old or {"id": "f" + uuid.uuid4().hex[:10], "created_by": by, "created_at": ts, "notes": []}
        src = data.get("src") if not old else old.get("src")
        if src and isinstance(src, dict) and src.get("sid"):
            ok = con.execute("SELECT 1 FROM submissions WHERE id=?", (str(src["sid"]),)).fetchone()
            src = {"sid": str(src["sid"])} if ok else None
        else:
            src = None
        x.update({"title": title, "detail": _fu_txt(data.get("detail"), 3000),
                  "zone_id": zone["id"] if zone else "", "zone": zone["name"] if zone else _fu_txt(data.get("zone"), 80),
                  "owner": _fu_txt(data.get("owner"), 80), "due": due, "urgent": bool(data.get("urgent")),
                  "photos": photos[:FU_MAX_PHOTOS], "src": src, "updated_at": ts, "updated_by": by})
        prev = (old or {}).get("status")
        if prev != status:
            x["status"] = status
            if old:
                x["notes"].append({"at": ts, "by": by, "kind": "status", "text": f"เปลี่ยนสถานะเป็น {FU_STATUS[status]}"})
            if status == "done":
                x["done_at"], x["done_by"] = ts, by
            else:
                x.pop("done_at", None), x.pop("done_by", None)
        removed = [f for f in (old or {}).get("photos", []) if f not in x["photos"]]
        con.execute("INSERT INTO followups(id, data, status, created_at, updated_at) VALUES(?,?,?,?,?) "
                    "ON CONFLICT(id) DO UPDATE SET data=excluded.data, status=excluded.status, updated_at=excluded.updated_at",
                    (x["id"], json.dumps(x, ensure_ascii=False), x["status"], x["created_at"], ts))
    for f in removed:
        _remove_media(f)
    return {"ok": True, "item": x}


@app.post("/api/followups/{fid}/note")
async def api_followup_note(fid: str, request: Request):
    """บันทึกความคืบหน้า"""
    data = await request.json()
    by, text = _fu_txt(data.get("by"), 60), _fu_txt(data.get("text"), 1000)
    if not by:
        raise HTTPException(400, "ตั้งชื่อผู้ใช้ก่อน")
    if not text:
        raise HTTPException(400, "พิมพ์ความคืบหน้าก่อน")
    with db() as con:
        r = con.execute("SELECT data FROM followups WHERE id=?", (fid,)).fetchone()
        if not r:
            raise HTTPException(404, "ไม่พบเรื่องนี้")
        x = json.loads(r["data"])
        ts = _fu_ms()
        x.setdefault("notes", []).append({"at": ts, "by": by, "text": text})
        x["notes"] = x["notes"][-200:]
        x["updated_at"], x["updated_by"] = ts, by
        con.execute("UPDATE followups SET data=?, updated_at=? WHERE id=?", (json.dumps(x, ensure_ascii=False), ts, fid))
    return {"ok": True, "item": x}


@app.post("/api/followups/photo")
async def api_followup_photo(file: UploadFile = File(...)):
    raw = await file.read()
    if len(raw) > 25 * 1024 * 1024:
        raise HTTPException(400, "ไฟล์รูปใหญ่เกินไป")
    month = now().strftime("%Y-%m")
    (UPLOAD_DIR / "followups" / month).mkdir(parents=True, exist_ok=True)
    rel = f"followups/{month}/{uuid.uuid4().hex[:12]}.jpg"
    if not await run_in_threadpool(_store_image, raw, rel):
        raise HTTPException(400, "ไฟล์นี้ไม่ใช่รูปภาพที่รองรับ")
    return {"ok": True, "file": rel, **media_info(rel)}


@app.delete("/api/followups/{fid}")
def api_followup_delete(fid: str, pin: str = ""):
    _check_edit_pin(pin)
    with db() as con:
        r = con.execute("SELECT data FROM followups WHERE id=?", (fid,)).fetchone()
        if not r:
            raise HTTPException(404, "ไม่พบเรื่องนี้")
        con.execute("DELETE FROM followups WHERE id=?", (fid,))
    for f in json.loads(r["data"]).get("photos") or []:
        _remove_media(f)
    return {"ok": True}


@app.get("/healthz")
def health():
    return {"ok": True, "time": now().isoformat()}


_FONT_RE = re.compile(r"^[A-Za-z0-9._-]+\.(woff2?|ttf)$")
_FONT_TYPES = {".woff2": "font/woff2", ".woff": "font/woff", ".ttf": "font/ttf"}


@app.get("/fonts/{fn}")
def web_font(fn: str):
    """ฟอนต์หน้าเว็บ: หาใน app/static/fonts ก่อน ถ้าไม่มีใช้ app/fonts (ที่เก็บเดิม)
    กันกรณีอัปเดตแค่ไฟล์หน้าเว็บแต่ยังไม่ได้ย้ายโฟลเดอร์ฟอนต์ → ไม่ขึ้น 404 ทุกครั้งที่เปิดหน้า"""
    if not _FONT_RE.match(fn):
        raise HTTPException(404, "ไม่พบไฟล์")
    for d in (BASE_DIR / "static" / "fonts", BASE_DIR / "fonts"):
        if (d / fn).is_file():
            return FileResponse(d / fn, media_type=_FONT_TYPES[Path(fn).suffix.lower()])
    raise HTTPException(404, "ไม่พบไฟล์")


app.mount("/uploads", StaticFiles(directory=UPLOAD_DIR), name="uploads")
app.mount("/", StaticFiles(directory=BASE_DIR / "static", html=True), name="static")
