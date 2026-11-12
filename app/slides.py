"""สร้างสไลด์ .pptx จากข้อมูลการตรวจของแต่ละรอบ"""
import io
import os
from datetime import date, datetime, timedelta
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Emu, Inches, Pt

SLIDES_VERSION = "2026-10-26"  # แสดงในหน้าต่าง server ตอนเปิด ใช้เช็คว่าใช้ไฟล์ล่าสุดอยู่

FONT = os.environ.get("SLIDE_FONT", "TH Sarabun New")
# TH Sarabun New ตัวเล็กกว่าฟอนต์ทั่วไป → ขยายทุกขนาดอัตโนมัติ และบีบระยะบรรทัดเล็กน้อย
SARABUN = "sarabun" in FONT.lower()
SCALE = float(os.environ.get("SLIDE_FONT_SCALE", "1.4" if SARABUN else "1.0"))
LINE_H = 1.22  # ระยะบรรทัด = 1.22 เท่าของขนาดตัวอักษร
_FONT_DIR = Path(__file__).resolve().parent / "fonts"
try:
    from PIL import ImageFont
    _MEASURE = ImageFont.truetype(str(_FONT_DIR / "THSarabunNew.ttf"), 100) if SARABUN else None
except Exception:
    _MEASURE = None

NAVY = RGBColor(0x1B, 0x2A, 0x41)
INK = RGBColor(0x0F, 0x17, 0x2A)
AMBER = RGBColor(0xF5, 0x9E, 0x0B)
GREEN = RGBColor(0x16, 0xA3, 0x4A)
RED = RGBColor(0xDC, 0x26, 0x26)
GREY = RGBColor(0x94, 0xA3, 0xB8)
YELLOW = RGBColor(0xEA, 0xB3, 0x08)
YELLOW_INK = RGBColor(0x85, 0x4D, 0x0E)
YELLOW_SOFT = RGBColor(0xFE, 0xF9, 0xC3)
NO_ACT = "ไม่มีกิจกรรม"
LIGHT = RGBColor(0xF1, 0xF5, 0xF9)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT = RGBColor(0x1E, 0x29, 0x3B)
MUTED = RGBColor(0x64, 0x74, 0x8B)

TH_MONTH = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
SW, SH = Inches(13.333), Inches(7.5)


def th_date(d: date) -> str:
    return f"{d.day} {TH_MONTH[d.month - 1]} {d.year + 543}"


def hm(iso: str) -> str:
    return datetime.fromisoformat(iso).strftime("%H:%M") if iso else ""


def _box(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(1)
    s.shadow.inherit = False
    return s


def _text(slide, x, y, w, h, text, size=14, bold=False, color=TEXT, align=PP_ALIGN.LEFT,
          anchor=MSO_ANCHOR.TOP, fill=None, shape=None):
    if fill is None and shape is None:
        s = slide.shapes.add_textbox(x, y, w, h)
    else:
        s = _box(slide, x, y, w, h, fill=fill, shape=shape or MSO_SHAPE.RECTANGLE)
    tf = s.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.08)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    lines = text if isinstance(text, list) else text.split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if isinstance(ln, tuple):
            runs = ln
        else:
            runs = ((ln, size, bold, color),)
        if SARABUN:  # ระยะบรรทัดแบบกำหนดตายตัว → ผลเหมือนกันทั้ง PowerPoint / LibreOffice / Google Slides
            p.line_spacing = Pt(round(max(r[1] for r in runs) * SCALE * LINE_H, 1))
        for (t, sz, b, c) in runs:
            r = p.add_run()
            r.text = t
            r.font.size = Pt(round(sz * SCALE, 1))
            r.font.bold = b
            r.font.color.rgb = c
            r.font.name = FONT
            _force_font(r)
    return s


def _force_font(run):
    """บังคับฟอนต์ให้ครบทุกชุดอักษร: ภาษาไทยใน PowerPoint ใช้ฟอนต์ complex script (a:cs)
    ถ้าไม่กำหนด PowerPoint จะแสดงภาษาไทยด้วยฟอนต์อื่น (เช่น Cordia/Tahoma)"""
    from pptx.oxml.ns import qn
    rPr = run._r.get_or_add_rPr()
    rPr.set("lang", "th-TH")
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", FONT)


def _header(slide, title, subtitle, round_label, page=""):
    _box(slide, 0, 0, SW, Inches(1.0), fill=NAVY)
    _box(slide, 0, Inches(1.0), SW, Inches(0.06), fill=AMBER)
    size = _fit_one_line(title, 10.0, (24, 22, 20, 18, 16))
    while _text_w_pt(title, size) > 10.0 * 72 - 16 and len(title) > 10:
        title = title[:-2].rstrip("… ") + "…"
    _text(slide, Inches(0.4), Inches(0.1), Inches(10.0), Inches(0.52), title, size, True, WHITE,
          anchor=MSO_ANCHOR.MIDDLE)
    _text(slide, Inches(0.4), Inches(0.6), Inches(10.0), Inches(0.34), subtitle, 13, False, RGBColor(0xCB, 0xD5, 0xE1))
    _text(slide, Inches(10.4), Inches(0.14), Inches(2.6), Inches(0.4), round_label, 13, True, AMBER, PP_ALIGN.RIGHT)
    if page:
        _text(slide, Inches(10.4), Inches(0.55), Inches(2.6), Inches(0.35), page, 11, False,
              RGBColor(0xCB, 0xD5, 0xE1), PP_ALIGN.RIGHT)


def _footer(slide, title):
    _text(slide, Inches(0.4), Inches(7.1), Inches(8), Inches(0.3), title, 9, False, MUTED)


def _progress(slide, x, y, w, h, frac, color=GREEN):
    _box(slide, x, y, w, h, fill=RGBColor(0xE2, 0xE8, 0xF0), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    if frac > 0:
        _box(slide, x, y, max(int(w * frac), int(h)), h, fill=color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)


VIDEO_EXT = {".mp4", ".mov", ".m4v", ".webm", ".3gp"}
VIDEO_MIME = {".mp4": "video/mp4", ".m4v": "video/mp4", ".mov": "video/quicktime",
              ".webm": "video/webm", ".3gp": "video/3gpp"}


def _is_video(path) -> bool:
    return Path(path).suffix.lower() in VIDEO_EXT


def _visual(path) -> Path:
    """ไฟล์ภาพที่ใช้วัดขนาด/แสดง: วิดีโอใช้ภาพหน้าปก (.jpg)"""
    return Path(str(path) + ".jpg") if _is_video(path) else Path(path)


def _add_picture_fit(slide, path: Path, x, y, w, h):
    try:
        with Image.open(_visual(path)) as im:
            iw, ih = im.size
    except Exception:
        _text(slide, x, y, w, h, "(เปิดรูปไม่ได้)", 12, color=MUTED, align=PP_ALIGN.CENTER,
              anchor=MSO_ANCHOR.MIDDLE, fill=LIGHT)
        return
    _box(slide, x, y, w, h, fill=LIGHT)
    scale = min(w / iw, h / ih)
    pw, ph = int(iw * scale), int(ih * scale)
    px, py = x + (w - pw) // 2, y + (h - ph) // 2
    if _is_video(path):
        # ฝังวิดีโอในสไลด์ (กดเล่นได้ใน PowerPoint) + ป้ายบอกว่าเป็นวิดีโอ
        slide.shapes.add_movie(str(path), px, py, pw, ph, poster_frame_image=str(_visual(path)),
                               mime_type=VIDEO_MIME.get(Path(path).suffix.lower(), "video/mp4"))
        _text(slide, px + Inches(0.08), py + Inches(0.08), Inches(1.1), Inches(0.32), "วิดีโอ", 10, True, WHITE,
              PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, fill=RGBColor(0x0F, 0x17, 0x2A), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    else:
        slide.shapes.add_picture(str(path), px, py, pw, ph)


PROC_ORDER = ["Rawmat", "จุดเตรียม", "ห้องเย็น", "ห้องเคมี", "บรรจุ", "ห้องอิ้ง", "Retort&Store", "Warehouse"]


def _proc_rank(p: str) -> int:
    if not p:
        return 999  # ยังไม่ระบุกระบวนการ → ท้ายสุด
    return PROC_ORDER.index(p) if p in PROC_ORDER else len(PROC_ORDER)


def build_pptx(title: str, data: dict, subs: list, upload_dir: Path, cutoff, report_name: str = "Process Control",
               order: str = "zone") -> bytes:
    """สไลด์ของ 1 รอบ"""
    return build_pptx_multi(title, [(data, subs)], upload_dir, cutoff, report_name, order=order)


def _range_label(days):
    a, b = days[0], days[-1]
    if a == b:
        return f"รอบ {th_date(a)}"
    if (a.year, a.month) == (b.year, b.month):
        return f"รอบ {a.day}–{b.day} {TH_MONTH[b.month - 1]} {b.year + 543}"
    return f"รอบ {th_date(a)} – {th_date(b)}"


def build_pptx_multi(title: str, days: list, upload_dir: Path, cutoff, report_name: str = "Process Control",
                     order: str = "zone") -> bytes:
    """สไลด์รวมหลายรอบ (days = [(data, subs), ...] เรียงตามวันที่)
    ลำดับ (หลายวัน): ปก → สิ่งที่พบปัญหาของทุกวัน (สรุป + รายละเอียด) → เรื่องที่ปกติของทุกวัน
      (ทั้งสองส่วนเรียงตามหัวข้อ ไม่แยกวัน) → แต่ละวัน: Checklist / พื้นที่ที่ยังไม่ตรวจ"""
    days = sorted(days, key=lambda d: d[0]["round"])
    cut = cutoff.strftime("%H:%M")
    rds = [date.fromisoformat(d["round"]) for d, _ in days]
    multi = len(days) > 1
    label_all = _range_label(rds)
    window_all = f"{th_date(rds[0])} {cut} น. – {th_date(rds[-1] + timedelta(days=1))} {cut} น."
    foot_all = f"{title} · {window_all}"

    prs = Presentation()
    prs.slide_width, prs.slide_height = SW, SH
    blank = prs.slide_layouts[6]

    # ---------- 1) Cover / summary
    s = prs.slides.add_slide(blank)
    _box(s, 0, 0, SW, SH, fill=INK)
    _box(s, 0, Inches(2.55), Inches(0.25), Inches(1.3), fill=AMBER)
    _text(s, Inches(0.6), Inches(0.6), Inches(12), Inches(0.5), title, 16, True, AMBER)
    ct = f"รายงานการตรวจพื้นที่ {report_name}"
    _text(s, Inches(0.6), Inches(1.1), Inches(12), Inches(1.2), ct, _fit_one_line(ct, 12, (38, 34, 30, 26)), True, WHITE)
    _text(s, Inches(0.6), Inches(2.5), Inches(10), Inches(0.6), label_all + (f"  ({len(days)} วัน)" if multi else ""),
          28, True, WHITE)
    _text(s, Inches(0.6), Inches(3.1), Inches(10), Inches(0.5), window_all, 16, False, GREY)

    done = sum(d["done"] for d, _ in days)
    total = sum(d["total"] for d, _ in days)
    n_issue = sum(d["issues"] for d, _ in days)
    pct = done / total if total else 0
    kpis = [
        (f"{done}/{total}", "พื้นที่ตรวจแล้ว" + (" (รวมทุกวัน)" if multi else ""), GREEN),
        (f"{total - done}", "ยังไม่ตรวจ", GREY if done == total else AMBER),
        (f"{n_issue}", "พื้นที่พบปัญหา", RED if n_issue else GREY),
        (f"{round(pct * 100)}%", "ความครบถ้วน", AMBER),
    ]
    for i, (v, lbl, c) in enumerate(kpis):
        x = Inches(0.6 + i * 3.1)
        _box(s, x, Inches(4.0), Inches(2.9), Inches(1.35), fill=RGBColor(0x1E, 0x29, 0x3B),
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        _text(s, x, Inches(4.05), Inches(2.9), Inches(0.8), v, 34, True, c, PP_ALIGN.CENTER)
        _text(s, x, Inches(4.85), Inches(2.9), Inches(0.4), lbl, 13, False, GREY, PP_ALIGN.CENTER)

    # แถวความคืบหน้า: 1 วัน = รายพื้นที่ / หลายวัน = รายวัน
    rows = ([(f"{th_date(rd)}", d["done"], d["total"]) for rd, (d, _) in zip(rds, days)] if multi
            else [(z["name"], z["done"], z["total"]) for z in days[0][0]["zones"]])
    y = Inches(5.65)
    colw = Inches(12.1 / max(1, len(rows)))
    for i, (nm, dn, tt) in enumerate(rows):
        x = Inches(0.6) + colw * i
        _text(s, x, y, colw - Inches(0.15), Inches(0.35), nm, 12, True, WHITE)
        _text(s, x, y + Inches(0.32), colw - Inches(0.15), Inches(0.3), f"{dn}/{tt}", 11, False, GREY)
        _progress(s, x + Inches(0.08), y + Inches(0.7), colw - Inches(0.35), Inches(0.12), dn / tt if tt else 0)
    gen = datetime.now().astimezone()
    _text(s, Inches(0.6), Inches(7.05), Inches(12), Inches(0.3),
          f"สร้างเมื่อ {th_date(gen.date())} {gen.strftime('%H:%M')} น.", 9, False, MUTED)

    # ---------- เตรียมข้อมูลแต่ละวัน
    per_day = []
    for rd, (data, subs) in zip(rds, days):
        round_label = f"รอบ {th_date(rd)}"
        window = f"{th_date(rd)} {cut} น. – {th_date(rd + timedelta(days=1))} {cut} น."
        foot = f"{title} · {window}"
        zone_order = {z["name"]: i for i, z in enumerate(data["zones"])}
        item_no = {it["id"]: it["no"] for z in data["zones"] for it in z["items"]}
        owners = {z["name"]: z["owner"] for z in data["zones"]}
        extra_ids = {it["id"] for z in data["zones"] for it in z["items"] if it.get("extra")}

        def _key(x, zone_order=zone_order, item_no=item_no, extra_ids=extra_ids):
            no = item_no.get(x["area_id"], 0)
            return (zone_order.get(x["zone"], 99), x["area_id"] in extra_ids,
                    no if isinstance(no, int) else 0, x["created_at"])
        ss = sorted(subs, key=_key)
        per_day.append(dict(rd=rd, data=data, round_label=round_label, foot=foot, item_no=item_no, owners=owners,
                            issues=[x for x in ss if x["status"] == "พบปัญหา"],
                            normal=[x for x in ss if x["status"] != "พบปัญหา"]))

    # ลำดับตามหัวข้อ (พื้นที่ → ลำดับรายการ) รวมทุกวัน ไม่แยกตามวัน; หัวข้อเดียวกันเรียงตามเวลา
    zrank, irank = {}, {}
    for d in reversed(per_day):  # ใช้โครงสร้างของวันล่าสุดเป็นหลัก
        for zi, z in enumerate(d["data"]["zones"]):
            zrank.setdefault(z["id"], zi)
            for ii, it in enumerate(z["items"]):
                irank.setdefault(it["id"], (bool(it.get("extra")), ii))

    def _topic_key(pair):
        x = pair[1]
        k = (zrank.get(x.get("zone_id"), 99), irank.get(x["area_id"], (True, 999)), x["round_date"], x["created_at"])
        if order == "proc":  # เรียงตามกระบวนการก่อน แล้วค่อยพื้นที่ / ลำดับรายการ (หัวข้อเพิ่มเติมไว้ท้าย)
            extra = irank.get(x["area_id"], (True, 999))[0]
            return (extra, _proc_rank(x.get("proc", "")),) + k
        return k

    # ---------- 2) รายละเอียด
    #   พบปัญหา: หน้าสรุป (ปัญหาเดียวกันรวมเป็นแถวเดียว) + สไลด์รายละเอียดเต็มหน้าแบบเดิม ทีละเรื่อง
    #   ปกติ: 1 หัวข้อ = 1 หน้า (รวมทุกครั้ง/ทุกวันของหัวข้อนั้น) + สรุปผลแบบย่อด้านบน
    #   (หัวข้อเพิ่มเติม แต่ละเรื่องแยกหน้า)
    hdr = label_all[4:] if multi else label_all  # หัวสไลด์มีที่จำกัด: "21–22 ต.ค. 2569"
    item_no_all = {k: v for d in per_day for k, v in d["item_no"].items()}
    every = sorted([(d, x) for d in per_day for x in d["issues"] + d["normal"]], key=_topic_key)
    groups, gidx = [], {}
    for d, x in every:
        key = ("x", x["id"]) if d["item_no"].get(x["area_id"]) == "+" else ("i", x["area_id"])
        if key not in gidx:
            gidx[key] = len(groups)
            groups.append([])
        groups[gidx[key]].append((d, x))
    is_ng = lambda p: p[1]["status"] == "พบปัญหา"
    ng_groups = [g for g in groups if any(is_ng(p) for p in g)]
    ok_groups = [g for g in groups if not any(is_ng(p) for p in g)]
    if ng_groups:
        _issue_summary(prs, blank, [[p for p in g if is_ng(p)] for g in ng_groups], item_no_all, hdr, foot_all,
                       show_date=multi)
    for d, x in [p for p in every if is_ng(p)]:
        _detail_slides(prs, blank, x, d["item_no"], d["owners"], d["round_label"], d["foot"], upload_dir)
    ok_rows = [[p for p in g if not is_ng(p)] for g in groups]
    ok_rows = [g for g in ok_rows if g]
    for g in ok_rows:
        _normal_topic_slide(prs, blank, g, hdr, foot_all, upload_dir, show_date=multi)

    # ---------- 3) แต่ละวัน: Checklist → พื้นที่ที่ยังไม่ตรวจ
    for d in per_day:
        data, round_label, foot = d["data"], d["round_label"], d["foot"]
        if multi:
            s = prs.slides.add_slide(blank)
            _box(s, 0, 0, SW, SH, fill=INK)
            _box(s, 0, Inches(3.0), Inches(0.25), Inches(1.3), fill=AMBER)
            _text(s, Inches(0.6), Inches(2.9), Inches(12), Inches(0.8), round_label, 40, True, WHITE)
            _text(s, Inches(0.6), Inches(3.95), Inches(12), Inches(0.5),
                  f"ตรวจแล้ว {data['done']}/{data['total']}  ·  พบปัญหา {data['issues']}  ·  ส่งทั้งหมด "
                  f"{len(d['issues']) + len(d['normal'])} รายการ", 16, False, GREY)
        # ---------- 2) Checklist per zone
        for z in data["zones"]:
            s = prs.slides.add_slide(blank)
            _header(s, f"Checklist · {z['name']}", f"ผู้รับผิดชอบ: {z['owner']}   |   ตรวจแล้ว {z['done']}/{z['total']} พื้นที่",
                    round_label)
            items = [it for it in z["items"] if not it.get("extra")]
            n_extra = sum(it["count"] for it in z["items"] if it.get("extra"))
            if n_extra:
                _text(s, Inches(8.4), Inches(0.55), Inches(4.6), Inches(0.35), f"+ หัวข้อเพิ่มเติม {n_extra} เรื่อง",
                      11, True, AMBER, PP_ALIGN.RIGHT)
            per_col = max(1, (len(items) + 1) // 2)
            cols = [items[:per_col], items[per_col:]]
            top = Inches(1.3)
            row_h = min(Inches(0.56), int((Inches(6.95) - top) / max(per_col, 1)))
            for ci, col in enumerate(cols):
                x0 = Inches(0.4 + ci * 6.3)
                for ri, it in enumerate(col):
                    yy = top + row_h * ri
                    ok = it["checked"]
                    bg = LIGHT if ri % 2 == 0 else WHITE
                    _box(s, x0, yy, Inches(6.15), row_h, fill=bg)
                    mark_c = RED if it["issue"] else (GREEN if ok else RGBColor(0xCB, 0xD5, 0xE1))
                    mark = "!" if it["issue"] else ("✓" if ok else "–")
                    _text(s, x0 + Inches(0.06), yy + (row_h - Inches(0.36)) // 2, Inches(0.36), Inches(0.36), mark,
                          13, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, fill=mark_c, shape=MSO_SHAPE.OVAL)
                    # ชื่อยาว → ลดขนาด/ตัดท้าย ให้อยู่ในแถวไม่เกิน 2 บรรทัด
                    rows_fit = max(1, int((row_h / 914400) / _line_h(11)))
                    nm, nsz = it["name"], 11
                    for nsz in (11, 10, 9):
                        if _wrapped(f"{it['no']}  " + nm, 4.3, nsz) <= min(2, int((row_h / 914400) / _line_h(nsz))):
                            break
                    while _wrapped(f"{it['no']}  " + nm, 4.3, nsz) > max(1, min(2, int((row_h / 914400) / _line_h(nsz)))) \
                            and len(nm) > 8:
                        nm = nm[:-2].rstrip("… ") + "…"
                    _text(s, x0 + Inches(0.48), yy, Inches(4.3), row_h, [(
                        (f"{it['no']}  ", nsz - 1, True, MUTED), (nm, nsz, False, TEXT if ok else MUTED))],
                        anchor=MSO_ANCHOR.MIDDLE)
                    who = f"{it['last_inspector']}\n{hm(it['last_time'])} น." if ok else "ยังไม่ตรวจ"
                    _text(s, x0 + Inches(4.75), yy, Inches(1.38), row_h, who, 9, ok, GREEN if ok else AMBER,
                          PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)
            _footer(s, foot)


        # ---------- 4) Not-yet-checked summary (แบ่งหลายหน้าอัตโนมัติ)
        pending = [(z, [it for it in z["items"] if not it["checked"] and not it.get("extra")]) for z in data["zones"]]
        left = sum(len(p) for _, p in pending)
        sub_t = f"ณ เวลาตัดรอบ {cut} น.  ·  ยังไม่ตรวจ {left} จาก {data['total']} พื้นที่"
        if left == 0:
            s = prs.slides.add_slide(blank)
            _header(s, "พื้นที่ที่ยังไม่ได้ตรวจ", sub_t, round_label)
            _text(s, Inches(0.4), Inches(2.8), Inches(12.5), Inches(1.2), "ตรวจครบทุกพื้นที่แล้ว ✓", 36, True, GREEN,
                  PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
            _footer(s, foot)
        else:
            lines = []
            for z, p in pending:
                if not p:
                    continue
                lines.append(((f"{z['name']}  ", 12, True, NAVY), (f"({len(p)})", 10, False, MUTED)))
                for it in p:
                    name = it["name"]
                    while _text_w_pt(f"{it['no']}  {name}", 10) > 4.0 * 72 - 16 and len(name) > 8:
                        name = name[:-2].rstrip("… ") + "…"  # ให้อยู่บรรทัดเดียว
                    lines.append(((f"{it['no']}  ", 9, True, MUTED), (name, 10, False, TEXT)))
            per_col, ncol = max(8, int(5.5 / _line_h(12))), 3
            cols = [[]]
            for ln in lines:
                is_head = len(ln) == 2 and ln[0][2] and ln[0][3] == NAVY
                if len(cols[-1]) >= per_col or (is_head and len(cols[-1]) >= per_col - 1):
                    cols.append([])
                cols[-1].append(ln)
            pages = [cols[i:i + ncol] for i in range(0, len(cols), ncol)]
            for pi, page in enumerate(pages):
                s = prs.slides.add_slide(blank)
                more = f"  ({pi + 1}/{len(pages)})" if len(pages) > 1 else ""
                _header(s, "พื้นที่ที่ยังไม่ได้ตรวจ" + more, sub_t, round_label)
                for ci, col in enumerate(page):
                    _text(s, Inches(0.4 + ci * 4.2), Inches(1.3), Inches(4.1), Inches(5.7), col)
                _footer(s, foot)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------- detail slide helpers
import math
import unicodedata

CONTENT_TOP = Inches(2.0)
CONTENT_H = Inches(4.95)
LEFT = Inches(0.4)
FULL_W = Inches(12.53)
GAP = Inches(0.25)


def clean_note(note: str) -> list:
    """แปลง \r\n → \n (ต้นเหตุของ _x000D_), ตัดช่องว่างท้ายบรรทัด และยุบบรรทัดว่างที่ติดกัน"""
    text = (note or "").replace("\r\n", "\n").replace("\r", "\n")
    out = []
    for ln in text.split("\n"):
        ln = ln.rstrip()
        if not ln and (not out or not out[-1]):
            continue
        out.append(ln)
    while out and not out[-1]:
        out.pop()
    return out


def _vis_len(s: str) -> int:
    # สระ/วรรณยุกต์บน-ล่างของไทยไม่กินความกว้าง
    return sum(1 for c in s if unicodedata.category(c) != "Mn")


def _fit_one_line(text, width_in, sizes):
    for sz in sizes:
        if _text_w_pt(text, sz) <= width_in * 72 - 16:
            return sz
    return sizes[-1]


def _text_w_pt(line: str, size: float) -> float:
    """ความกว้างข้อความ (pt) ที่ขนาดตัวอักษร size (ก่อนคูณ SCALE)"""
    pt = size * SCALE
    if _MEASURE is not None:
        return _MEASURE.getlength(line) / 100 * pt * 1.06
    return _vis_len(line) * pt * 0.55


def _wrapped(line: str, width_in: float, size: int) -> int:
    usable = max(40.0, width_in * 72 - 14)  # หักขอบในกล่องข้อความ
    return max(1, math.ceil(_text_w_pt(line, size) / usable))


def _line_h(size):  # ความสูงบรรทัด (นิ้ว) เผื่อสระบน-ล่างของภาษาไทย
    if SARABUN:
        return size * SCALE * LINE_H / 72
    return size * 1.42 / 72


def _blocks(lines):
    out, cur = [], []
    for ln in lines:
        if ln:
            cur.append(ln)
        elif cur:
            out.append(cur)
            cur = []
    if cur:
        out.append(cur)
    return out


def _pack(blocks, cw, size, cap, maxcols):
    """วางข้อความลงคอลัมน์ โดยพยายามไม่ตัดย่อหน้ากลางคอลัมน์ คืน list ของคอลัมน์ หรือ None"""
    cols, cur, used = [], [], 0.0
    for blk in blocks:
        hs = [_wrapped(ln, cw, size) for ln in blk]
        need = sum(hs) + (1 if cur else 0)
        if used + need <= cap:
            if cur:
                cur.append("")
            cur += blk
            used += need
            continue
        if sum(hs) <= cap:  # ย้ายทั้งย่อหน้าไปคอลัมน์ใหม่
            cols.append(cur)
            cur, used = list(blk), float(sum(hs))
            continue
        # ย่อหน้ายาวเกินคอลัมน์ → ตัดทีละบรรทัด
        if cur:
            cur.append("")
            used += 1
        for ln, h in zip(blk, hs):
            if h > cap:
                return None
            if used + h > cap:
                cols.append(cur)
                cur, used = [], 0.0
            cur.append(ln)
            used += h
    if cur:
        cols.append(cur)
    cols = [c for c in cols if c]
    if maxcols and len(cols) > maxcols:
        return None
    return cols


def _ncols_options(width_in, maxcols):
    return [n for n in range(1, maxcols + 1) if n == 1 or (width_in - 0.3 * (n - 1)) / n >= (2.2 if SARABUN else 2.8)]


def fit_note(lines, width_in, height_in, maxcols=1, sizes=(20, 18, 16, 15, 14, 13, 12)):
    """หาขนาดตัวอักษรใหญ่สุด (และจำนวนคอลัมน์) ที่ใส่ข้อความได้พอดีกล่อง คืน (size, cols) หรือ None"""
    blocks = _blocks(lines)
    for size in sizes:
        cap = int(height_in / _line_h(size))
        for nc in _ncols_options(width_in, maxcols):
            cw = (width_in - 0.3 * (nc - 1)) / nc
            cols = _pack(blocks, cw, size, cap, nc)
            if cols:
                return size, cols
    return None


def split_note(lines, width_in, height_in, size=13):
    """ข้อความยาวมาก: แบ่งเป็นหลายหน้า หน้าละหลายคอลัมน์"""
    nc = _ncols_options(width_in, 4)[-1]
    cw = (width_in - 0.3 * (nc - 1)) / nc
    cols = _pack(_blocks(lines), cw, size, int(height_in / _line_h(size)), 0) or [lines]
    return [(size, cols[i:i + nc]) for i in range(0, len(cols), nc)]


def _note_card(slide, x, y, w, h, size, cols, title="รายละเอียด / ข้อสังเกต"):
    _box(slide, x, y, w, h, fill=LIGHT, shape=MSO_SHAPE.RECTANGLE)
    _box(slide, x, y, Inches(0.07), h, fill=AMBER)
    _text(slide, x + Inches(0.2), y + Inches(0.12), w - Inches(0.3), Inches(0.35), title, 11, True, MUTED)
    inner_x, inner_y = x + Inches(0.2), y + Inches(0.5)
    inner_w, inner_h = w - Inches(0.35), h - Inches(0.6)
    nc = len(cols)
    colgap = Inches(0.3)
    cw = int((inner_w - colgap * (nc - 1)) / nc)
    for i, col in enumerate(cols):
        paras = []
        for j, ln in enumerate(col):
            nxt = col[j + 1] if j + 1 < len(col) else ""
            is_head = (j == 0 or not col[j - 1]) and ln and nxt and _vis_len(ln) <= 40 \
                and not ln.lstrip().startswith(("-", "•", "*"))
            paras.append(((ln or " ", size, bool(is_head), NAVY if is_head else TEXT),))  # หัวข้อย่อหน้า = ตัวหนา
        tb = _text(slide, inner_x + (cw + colgap) * i, inner_y, cw, inner_h, paras, size, False, TEXT)
        pass


def _image_sizes(paths):
    out = []
    for p in paths:
        try:
            with Image.open(_visual(p)) as im:
                out.append(im.size)
        except Exception:
            out.append((4, 3))
    return out


def _image_grid(slide, paths, x, y, w, h, gap=Inches(0.12)):
    """จัดรูปให้เต็มพื้นที่มากที่สุด: ลองทุกแบบ แถว×คอลัมน์ แล้วเลือกแบบที่รูปใหญ่สุด"""
    n = len(paths)
    if not n:
        return
    sizes = _image_sizes(paths)
    best = None
    for cols in range(1, n + 1):
        rows = math.ceil(n / cols)
        cw = (w - gap * (cols - 1)) / cols
        ch = (h - gap * (rows - 1)) / rows
        area = 0
        for iw, ih in sizes:
            sc = min(cw / iw, ch / ih)
            area += iw * sc * ih * sc
        if best is None or area > best[0]:
            best = (area, cols, rows, cw, ch)
    _, cols, rows, cw, ch = best
    for r in range(rows):
        row_items = list(range(r * cols, min(n, (r + 1) * cols)))
        offset = (cols - len(row_items)) * (cw + gap) / 2  # จัดแถวสุดท้ายให้อยู่กลาง
        for k, idx in enumerate(row_items):
            _add_picture_fit(slide, paths[idx], int(x + offset + (cw + gap) * k), int(y + (ch + gap) * r),
                             int(cw), int(ch))


def _info_strip(slide, sub, n_img):
    issue, noact = sub["status"] == "พบปัญหา", sub["status"] == NO_ACT
    y = Inches(1.3)
    _text(slide, LEFT, y, Inches(1.6), Inches(0.45), sub["status"], 14, True, YELLOW_INK if noact else WHITE, PP_ALIGN.CENTER,
          MSO_ANCHOR.MIDDLE, fill=RED if issue else YELLOW if noact else GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    dt = datetime.fromisoformat(sub["created_at"])
    _text(slide, LEFT + Inches(1.8), y, Inches(8.5), Inches(0.45), [(
        ("ผู้ตรวจ  ", 11, False, MUTED), (sub["inspector"], 15, True, TEXT),
        ("        เวลา  ", 11, False, MUTED), (f"{th_date(dt.date())}  {dt.strftime('%H:%M')} น.", 13, False, TEXT),
    ) + ((("   (ส่งย้อนหลัง)", 11, True, AMBER),) if sub.get("late") else ())], anchor=MSO_ANCHOR.MIDDLE)
    if n_img:
        _text(slide, Inches(10.4), y, Inches(2.53), Inches(0.45), f"รูป/วิดีโอ {n_img} ไฟล์", 11, False, MUTED,
              PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)
    _box(slide, LEFT, Inches(1.85), FULL_W, Inches(0.02), fill=RGBColor(0xE2, 0xE8, 0xF0))


def _truncate_to_fit(lines, width_in, height_in, size, ncols):
    """กรณีข้อความยาวมากจนย่อสุดแล้วยังไม่พอ: ตัดท้ายแล้วใส่ … (ยังอยู่ใน 1 สไลด์)"""
    cw = (width_in - 0.3 * (ncols - 1)) / ncols
    cap = int(height_in / _line_h(size))
    cols, cur, used = [], [], 0
    for ln in lines:
        h = _wrapped(ln, cw, size) if ln else 1
        if used + h > cap:
            cols.append(cur)
            cur, used = [], 0
            if len(cols) == ncols:
                break
            if not ln:
                continue
        cur.append(ln)
        used += h
    if len(cols) < ncols and cur:
        cols.append(cur)
    if cols and cols[-1]:
        cols[-1][-1] = cols[-1][-1][:60] + " …"
    return size, [c for c in cols if c]


def _plan_page(lines, imgs):
    """1 สไลด์ที่ผู้ส่งกำหนด = 1 หน้า PowerPoint เสมอ (ไม่ตัดขึ้นหน้าใหม่)
    ระบบย่อตัวอักษร/ขยายคอลัมน์/ย่อรูป ให้พอดีหน้าเดียว"""
    Hin = CONTENT_H / 914400
    Win = FULL_W / 914400
    PAD_W, PAD_H = 0.35, 0.6
    ALL = (22, 20, 18, 16, 15, 14, 13, 12, 11, 10, 9, 8)

    if not lines and not imgs:
        return dict(layout="empty")
    if not lines:
        return dict(layout="img", imgs=imgs)
    if not imgs:
        fit = fit_note(lines, Win - PAD_W, Hin - PAD_H, maxcols=4, sizes=ALL) \
            or _truncate_to_fit(lines, Win - PAD_W, Hin - PAD_H, 8, 4)
        return dict(layout="text", note=fit)

    # มีทั้งข้อความและรูป
    short = fit_note(lines, Win - PAD_W, 3 * _line_h(16), maxcols=1, sizes=(18, 16))
    if short and sum(_wrapped(l, Win - PAD_W, short[0]) for c in short[1] for l in c) <= 3 and len(imgs) <= 6:
        n_lines = sum(len(c) for c in short[1])
        band = Inches(PAD_H + 0.1 + n_lines * _line_h(short[0]))
        return dict(layout="band", note=short, note_h=band, imgs=imgs)
    # ยิ่งข้อความยาว กล่องข้อความยิ่งกว้าง (รูปแคบลง) และตัวอักษรเล็กลงตามลำดับ
    for wi, mc, sizes in ((4.3, 1, (18, 16, 15, 14)), (5.4, 2, (16, 15, 14, 13)),
                          (6.4, 2, (15, 14, 13, 12)), (7.6, 3, (14, 13, 12, 11, 10)),
                          (8.4, 3, (11, 10, 9, 8))):
        fit = fit_note(lines, wi - PAD_W, Hin - PAD_H, maxcols=mc, sizes=sizes)
        if fit:
            return dict(layout="side", note=fit, note_w=Inches(wi), imgs=imgs)
    fit = _truncate_to_fit(lines, 8.4 - PAD_W, Hin - PAD_H, 8, 3)
    return dict(layout="side", note=fit, note_w=Inches(8.4), imgs=imgs)


PASS_MARKS = ("✅", "✔", "☑", "✓")
FAIL_MARKS = ("❌", "✖", "✗", "✘", "❎")


def _short_summary(name: str, subs: list) -> str:
    """สรุปผลแบบย่อของเรื่องที่ปกติ"""
    return "ผ่าน · ไม่พบปัญหา"


def _first_line(x) -> str:
    for p in (x.get("pages") or [{"note": x.get("note")}]):
        for ln in clean_note(p.get("note") or ""):
            t = ln.strip()
            for m in PASS_MARKS + FAIL_MARKS:
                t = t.replace(m, "")
            t = t.strip(" -•·:")
            if t:
                return t
    return ""


def _thumb_stream(path: Path, side_px: int = 260):
    """รูปย่อครอปสี่เหลี่ยม (วิดีโอใช้ภาพหน้าปก)"""
    try:
        with Image.open(_visual(path)) as im:
            im = im.convert("RGB")
            w, h = im.size
            m = min(w, h)
            im = im.crop(((w - m) // 2, (h - m) // 2, (w - m) // 2 + m, (h - m) // 2 + m))
            im.thumbnail((side_px, side_px))
            buf = io.BytesIO()
            im.save(buf, "JPEG", quality=80)
            buf.seek(0)
            return buf
    except Exception:
        return None


def _normal_topic_slide(prs, blank, group, round_label, foot, upload_dir, show_date=False):
    """เรื่องที่ปกติ: 1 หัวข้อ = 1 หน้า
    แถบสรุปผลแบบย่อ ("ผ่าน · ไม่พบปัญหา") แล้วตามด้วยการ์ดของแต่ละครั้งที่ส่ง (รูป + รายละเอียด)
    เกิน 4 การ์ด (ครั้ง/สไลด์ย่อย) ต่อหน้า จะแบ่งหน้า โดยยังเป็นหัวข้อเดียวกัน"""
    d0, x0 = group[-1]
    no = d0["item_no"].get(x0["area_id"], "")
    topic = (x0.get("topic") or "").strip()
    title = (f"เพิ่มเติม: {topic}" if topic else "หัวข้อเพิ่มเติม (นอกเหนือ Checklist)") if no == "+" \
        else f"{no}. {x0['area_name']}"
    days = sorted({x["round_date"] for _, x in group})
    subtitle = f"{x0['zone']}" + (f"  ·  {x0['proc']}" if x0.get("proc") else "") + \
        f"  ·  {d0['owners'].get(x0['zone'], '')}  ·  ส่ง {len(group)} ครั้ง" + (
        f" ใน {len(days)} วัน" if len(days) > 1 else "")
    summ = _short_summary(title, [x for _, x in group])
    n_na = sum(1 for _, x in group if x["status"] == NO_ACT)
    all_na = n_na == len(group)
    if n_na and not all_na:
        summ += f"  ·  {NO_ACT} {n_na} ครั้ง"
    cards = []
    for d, x in group:
        pages = [p for p in (x.get("pages") or [{"note": x.get("note"), "images": x.get("images", [])}])
                 if (p.get("note") or "").strip() or p.get("images")] or [{"note": "", "images": []}]
        for pi, pg in enumerate(pages):
            cards.append((x, pg, pi, len(pages)))
    per = 4
    npg = math.ceil(len(cards) / per)
    per = math.ceil(len(cards) / npg)
    chunks = [cards[i:i + per] for i in range(0, len(cards), per)]
    for ci, chunk in enumerate(chunks):
        s = prs.slides.add_slide(blank)
        _header(s, title, subtitle, round_label, f"หน้า {ci + 1}/{len(chunks)}" if len(chunks) > 1 else "")
        # แถบสรุปผลแบบย่อ (ไม่มีกิจกรรมทุกครั้ง = แถบเหลือง)
        _box(s, Inches(0.4), Inches(1.25), Inches(12.53), Inches(0.5), fill=YELLOW_SOFT if all_na else RGBColor(0xDC, 0xFC, 0xE7))
        _box(s, Inches(0.4), Inches(1.25), Inches(0.08), Inches(0.5), fill=YELLOW if all_na else GREEN)
        _text(s, Inches(0.6), Inches(1.25), Inches(12.2), Inches(0.5), "– " + NO_ACT if all_na else "✓ " + summ, 15, True,
              YELLOW_INK if all_na else GREEN, anchor=MSO_ANCHOR.MIDDLE)
        n = len(chunk)
        cols = min(n, 2) if n <= 4 else 3
        rows = math.ceil(n / cols)
        gap, top, avail_h = 0.18, 1.9, 5.05
        cw = (12.53 - gap * (cols - 1)) / cols
        ch = (avail_h - gap * (rows - 1)) / rows
        for i, (x, pg, pi, npages) in enumerate(chunk):
            r, c = divmod(i, cols)
            cx, cy = 0.4 + c * (cw + gap), top + r * (ch + gap)
            _box(s, Inches(cx), Inches(cy), Inches(cw), Inches(ch), fill=LIGHT)
            dt = datetime.fromisoformat(x["created_at"])
            rd = date.fromisoformat(x["round_date"])
            strip = f"{rd.day} {TH_MONTH[rd.month - 1]} {rd.year + 543}  {dt.strftime('%H:%M')} น.  ·  {x['inspector']}" + \
                    (f"  ·  สไลด์ {pi + 1}/{npages}" if npages > 1 else "") + ("  (ส่งย้อนหลัง)" if x.get("late") else "") + \
                    (f"  ·  {NO_ACT}" if x["status"] == NO_ACT and not all_na else "")
            while _text_w_pt(strip, 11) > (cw - 0.2) * 72 and len(strip) > 10:
                strip = strip[:-2].rstrip("… ") + "…"
            _text(s, Inches(cx), Inches(cy), Inches(cw), Inches(0.34), strip, 11, True, WHITE,
                  anchor=MSO_ANCHOR.MIDDLE, fill=GREEN)
            iy, ih = cy + 0.4, ch - 0.46
            imgs = [upload_dir / f for f in pg.get("images", []) if (upload_dir / f).exists()]
            lines = clean_note(pg.get("note") or "")
            if pi == 0 and (x.get("remark") or "").strip():
                lines += ["", "หมายเหตุ: " + " ".join(clean_note(x["remark"]))]
            wide = cw > 5  # การ์ดกว้าง: ข้อความซ้าย รูปขวา / การ์ดแคบ: รูปบน ข้อความล่าง
            if imgs and lines:
                if wide:
                    tw = cw * 0.45
                    _image_grid(s, imgs[:6], Inches(cx + tw + 0.1), Inches(iy), Inches(cw - tw - 0.2), Inches(ih),
                                gap=Inches(0.06))
                    tx, ty, tth = cx + 0.1, iy, ih
                else:
                    hh = ih * 0.55
                    _image_grid(s, imgs[:4], Inches(cx + 0.1), Inches(iy), Inches(cw - 0.2), Inches(hh), gap=Inches(0.06))
                    tx, ty, tw, tth = cx + 0.1, iy + hh + 0.06, cw - 0.2, ih - hh - 0.06
            elif imgs:
                _image_grid(s, imgs[:6], Inches(cx + 0.1), Inches(iy), Inches(cw - 0.2), Inches(ih), gap=Inches(0.06))
                continue
            else:
                tx, ty, tw, tth = cx + 0.1, iy, cw - 0.2, ih
                if not lines:
                    _text(s, Inches(tx), Inches(ty), Inches(tw), Inches(tth), "ไม่มีรายละเอียดและรูปภาพ", 11, False,
                          MUTED, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
                    continue
            if imgs and wide:
                tw = cw * 0.45 - 0.1
            # ย่อขนาดตัวอักษรให้พอดีการ์ด ถ้ายังไม่พอตัดท้าย
            for size in (13, 12, 11, 10, 9):
                need = sum(max(1, _wrapped(ln, tw, size)) if ln else 0.5 for ln in lines) * _line_h(size)
                if need <= tth:
                    break
            maxl = int(tth / _line_h(size))
            out, used = [], 0
            for ln in lines:
                k = max(1, _wrapped(ln, tw, size)) if ln else 1
                if used + k > maxl:
                    if used < maxl:
                        t = ln
                        while t and _wrapped(t + "…", tw, size) > maxl - used and len(t) > 4:
                            t = t[:-2]
                        out.append(t.rstrip() + "…")
                    break
                out.append(ln)
                used += k
            _text(s, Inches(tx), Inches(ty), Inches(tw), Inches(tth), [((ln, size, ln.startswith("หมายเหตุ"), TEXT),) for ln in out])
        _footer(s, foot)


def _issue_summary(prs, blank, groups, item_no, round_label, foot, show_date=False):
    """หน้าสรุปสิ่งที่พบปัญหา: 1 แถว = 1 ปัญหา (ปัญหาเดียวกันที่พบหลายครั้งรวมเป็นแถวเดียว)
    สาเหตุ/การแก้ไข ใช้ข้อมูลล่าสุดที่กรอกไว้"""
    rows = []
    n_times = sum(len(g) for g in groups)
    for g in groups:
        x = g[-1][1]
        no = item_no.get(x["area_id"], "")
        topic = (x.get("topic") or "").strip()
        name = (f"เพิ่มเติม: {topic}" if no == "+" else f"{no}. {x['area_name']}" + (f" – {topic}" if topic else ""))
        latest = lambda f: next((" ".join(clean_note(p[1].get(f) or "")) for p in reversed(g) if (p[1].get(f) or "").strip()),
                                "— ยังไม่ระบุ —")
        if len(g) == 1:
            info = (f"{th_date(date.fromisoformat(x['round_date']))} · " if show_date else "") + \
                   f"{x['zone']} · {x['inspector']} · {hm(x['created_at'])} น."
        else:
            days = sorted({date.fromisoformat(p[1]["round_date"]) for p in g})
            dl = ", ".join(f"{d.day}" for d in days) + f" {TH_MONTH[days[-1].month - 1]}" if len(days) > 1 else th_date(days[0])
            info = f"{x['zone']} · พบ {len(g)} ครั้ง" + (f" ({dl})" if show_date or len(days) > 1 else "")
        rows.append((name, info, latest("cause"), latest("action"), len(g)))
    per = 6
    pages = [rows[i:i + per] for i in range(0, len(rows), per)]
    title = f"สิ่งที่พบปัญหา ({len(groups)} เรื่อง" + (f" · {n_times} ครั้ง)" if n_times > len(groups) else ")")
    for pi, page in enumerate(pages):
        s = prs.slides.add_slide(blank)
        more = f"หน้า {pi + 1}/{len(pages)}" if len(pages) > 1 else ""
        _header(s, title, "สรุปปัญหา สาเหตุ และการแก้ไข (ปัญหาเดียวกันรวมเป็นแถวเดียว) · รายละเอียดอยู่ในสไลด์ถัดไป",
                round_label, more)
        top, rh = Inches(1.3), Inches(0.92)
        for ri, (name, info, c, a, k) in enumerate(page):
            y = top + rh * ri
            _box(s, LEFT, y, FULL_W, rh - Inches(0.08), fill=RGBColor(0xFE, 0xF2, 0xF2))
            _box(s, LEFT, y, Inches(0.08), rh - Inches(0.08), fill=RED)
            nm = name
            while _text_w_pt(nm, 13) > 6.2 * 72 - 16 and len(nm) > 8:
                nm = nm[:-2].rstrip("… ") + "…"
            _text(s, LEFT + Inches(0.2), y + Inches(0.04), Inches(6.3), Inches(0.42), nm, 13, True, TEXT, anchor=MSO_ANCHOR.MIDDLE)
            while _text_w_pt(info, 10) > 6.2 * 72 - 16 and len(info) > 8:
                info = info[:-2].rstrip("… ") + "…"
            _text(s, LEFT + Inches(0.2), y + Inches(0.44), Inches(6.3), Inches(0.36), info, 10, k > 1, RED if k > 1 else MUTED,
                  anchor=MSO_ANCHOR.MIDDLE)
            while _wrapped("สาเหตุ: " + c, 5.9, 11) > 1 and len(c) > 8:
                c = c[:-2].rstrip("… ") + "…"
            while _wrapped("การแก้ไข: " + a, 5.9, 11) > 1 and len(a) > 8:
                a = a[:-2].rstrip("… ") + "…"
            _text(s, LEFT + Inches(6.6), y + Inches(0.04), Inches(5.9), rh - Inches(0.16),
                  [(("สาเหตุ: ", 11, True, RED), (c, 11, False, TEXT)),
                   (("การแก้ไข: ", 11, True, GREEN), (a, 11, False, TEXT))], anchor=MSO_ANCHOR.MIDDLE)
        _footer(s, foot)


def _detail_slides(prs, blank, sub, item_no, owners, round_label, foot, upload_dir):
    no = item_no.get(sub['area_id'], '')
    topic = (sub.get("topic") or "").strip()
    if no == "+":
        title = f"เพิ่มเติม: {topic}" if topic else "หัวข้อเพิ่มเติม (นอกเหนือ Checklist)"
    else:
        title = f"{no}. {sub['area_name']}" + (f" – {topic}" if topic else "")
    subtitle = f"{sub['zone']}" + (f"  ·  {sub['proc']}" if sub.get("proc") else "") + f"  ·  {owners.get(sub['zone'], '')}"
    user_pages = sub.get("pages") or [{"note": sub.get("note"), "images": sub.get("images", [])}]
    n_img_total = sum(len(p["images"]) for p in user_pages)

    plan = []
    # หมายเหตุ / การแก้ไข → ต่อท้ายข้อความของสไลด์แรก (หัวข้อเป็นตัวหนา)
    extra_lines = []
    if (sub.get("cause") or "").strip():
        extra_lines += ["", "สาเหตุของปัญหา"] + clean_note(sub["cause"])
    if (sub.get("action") or "").strip():
        extra_lines += ["", "การแก้ไข"] + clean_note(sub["action"])
    if (sub.get("remark") or "").strip():
        extra_lines += ["", "หมายเหตุ"] + clean_note(sub["remark"])
    for ui, up in enumerate(user_pages):
        lines = clean_note(up.get("note"))
        if ui == 0 and extra_lines:
            lines = (lines + extra_lines) if lines else extra_lines[1:]
        imgs = [upload_dir / f for f in up["images"] if (upload_dir / f).exists()]
        if not lines and not imgs and len(user_pages) > 1:
            continue
        pg = _plan_page(lines, imgs)
        pg["cont"] = False
        plan.append(pg)

    for pi, pg in enumerate(plan):
        s = prs.slides.add_slide(blank)
        cont = f"หน้า {pi + 1}/{len(plan)}" if len(plan) > 1 else ""
        _header(s, title, subtitle, round_label, cont)
        _info_strip(s, sub, n_img_total)
        lay = pg["layout"]
        ttl = "รายละเอียด / ข้อสังเกต" + (" (ต่อ)" if pg["cont"] else "")
        if lay == "text":
            _note_card(s, LEFT, CONTENT_TOP, FULL_W, CONTENT_H, *pg["note"], ttl)
        elif lay == "side":
            nw = pg["note_w"]
            _note_card(s, LEFT, CONTENT_TOP, nw, CONTENT_H, *pg["note"], ttl)
            _image_grid(s, pg["imgs"], LEFT + nw + GAP, CONTENT_TOP, FULL_W - nw - GAP, CONTENT_H)
        elif lay == "band":
            bh = pg["note_h"]
            _note_card(s, LEFT, CONTENT_TOP, FULL_W, bh, *pg["note"], ttl)
            _image_grid(s, pg["imgs"], LEFT, CONTENT_TOP + bh + Inches(0.15), FULL_W, CONTENT_H - bh - Inches(0.15))
        elif lay == "img":
            _image_grid(s, pg["imgs"], LEFT, CONTENT_TOP, FULL_W, CONTENT_H)
        else:
            _text(s, LEFT, CONTENT_TOP, FULL_W, CONTENT_H, "ไม่มีรายละเอียดและรูปภาพ", 16, False, MUTED,
                  PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, fill=LIGHT)
        _footer(s, foot)