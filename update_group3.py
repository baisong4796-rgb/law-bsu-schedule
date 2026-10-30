#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Crawl law.bsu.by "Class Schedule 3,4,5,6,7,8 groups" xlsx,
extract Group 3 only, and generate an iCalendar (.ics) file.
Re-run weekly (every Saturday) to refresh the feed.
"""
import ssl, re, json, os, urllib.request, hashlib
from datetime import datetime, date, time, timedelta, timezone

import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX_URL = "https://law.bsu.by/pub/master%20couses/schedule_1_of_3_4_5_6_7_8gr.xlsx"
LOCAL_XLSX = os.path.join(BASE_DIR, "Class Schedule 3,4,5,6,7,8 groups.xlsx")
OUT_ICS = os.path.join(BASE_DIR, "group3_schedule.ics")
OUT_JSON = os.path.join(BASE_DIR, "group3_events.json")

GROUP3_COL = 3   # column C
DATE_COL = 1     # column A
TIME_COL = 2     # column B

WEEKDAY_RE = re.compile(r"Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|"
                        r"Понедельник|Вторник|Среда|Четверг|Пятница|Суббота|Воскресенье", re.I)
DATE_RE = re.compile(r"(\d{1,2})\.(\d{1,2})")
TIME_RE = re.compile(r"(\d{1,2})[.:](\d{2})\s*-\s*(\d{1,2})[.:](\d{2})")
YEAR_RE = re.compile(r"(\d{4})\s*/\s*(\d{4})")
ROOM_RE = re.compile(r"(?:room|classroom|аудитория|ауд\.?)\s*(\d+)", re.I)


def download():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    req = urllib.request.Request(XLSX_URL, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, context=ctx, timeout=60) as r:
            data = r.read()
        with open(LOCAL_XLSX, "wb") as f:
            f.write(data)
        return LOCAL_XLSX
    except Exception as e:
        print("download failed (%s), using local copy" % e)
        return LOCAL_XLSX if os.path.exists(LOCAL_XLSX) else None


def normalize_time(s):
    m = TIME_RE.search(s or "")
    if not m:
        return None, None
    h1, m1, h2, m2 = (int(x) for x in m.groups())
    return time(h1, m1), time(h2, m2)


def english_title(text):
    t = text.strip()
    if "/" in t:
        t = t.split("/", 1)[1]
    t = re.sub(r"\((?:seminar|lecture|lectures|практическое|лекция|лекции|семинар|практич\.?|занятие)[^)]*\)", " ", t, flags=re.I)
    t = re.sub(r"\((?:room|classroom|аудитория|ауд\.?)\s*\d*\s*\)?", " ", t, flags=re.I)
    t = re.sub(r"Lecturer\s+[A-Za-zА-Яа-яёЁ .'-]+", " ", t, flags=re.I)
    t = re.sub(r"\(\s*\)", " ", t)
    t = re.sub(r"\s{2,}", " ", t).strip(" -–—")
    return t


def room_of(text):
    m = ROOM_RE.search(text or "")
    return m.group(1) if m else ""


def extract(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    events = []
    year = 2026
    for ws in wb.worksheets:
        # find academic year in the sheet
        for row in ws.iter_rows(values_only=True):
            for v in row:
                if isinstance(v, str) and "учебн" in v.lower():
                    ym = YEAR_RE.search(v)
                    if ym:
                        year = int(ym.group(1))
        # find header row
        hdr = None
        for r in range(1, ws.max_row + 1):
            if "Group 3" in str(ws.cell(row=r, column=GROUP3_COL).value or ""):
                hdr = r
                break
        if hdr is None:
            continue
        # find first data row (a weekday in column A)
        start = None
        for r in range(hdr + 1, ws.max_row + 1):
            a = ws.cell(row=r, column=DATE_COL).value
            if isinstance(a, str) and WEEKDAY_RE.search(a):
                start = r
                break
        if start is None:
            continue
        cur_date = None
        cur_time = None
        for r in range(start, ws.max_row + 1):
            a = ws.cell(row=r, column=DATE_COL).value
            b = ws.cell(row=r, column=TIME_COL).value
            c = ws.cell(row=r, column=GROUP3_COL).value
            if isinstance(a, str) and a.strip():
                dm = DATE_RE.search(a)
                if dm:
                    cur_date = date(year, int(dm.group(2)), int(dm.group(1)))
            t1, t2 = normalize_time(b if isinstance(b, str) else "")
            if t1:
                cur_time = (t1, t2)
            if isinstance(c, str) and c.strip():
                t1, t2 = cur_time if cur_time else (None, None)
                if cur_date and t1 and t2:
                    events.append({
                        "date": cur_date.isoformat(),
                        "start": t1.strftime("%H:%M"),
                        "end": t2.strftime("%H:%M"),
                        "text": c.strip(),
                        "summary": english_title(c),
                        "room": room_of(c),
                        "sheet": ws.title,
                    })
    # dedupe & sort
    seen, uniq = set(), []
    for e in events:
        key = (e["date"], e["start"], e["summary"])
        if key in seen:
            continue
        seen.add(key)
        uniq.append(e)
    uniq.sort(key=lambda e: (e["date"], e["start"]))
    return uniq


def to_ics(events):
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//law.bsu.by//Group3 Schedule//RU",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:Group 3 Class Schedule (law.bsu.by)",
        "X-WR-TIMEZONE:Europe/Minsk",
        "BEGIN:VTIMEZONE",
        "TZID:Europe/Minsk",
        "BEGIN:STANDARD",
        "DTSTART:19700101T000000",
        "TZOFFSETFROM:+0300",
        "TZOFFSETTO:+0300",
        "TZNAME:+03",
        "END:STANDARD",
        "END:VTIMEZONE",
    ]
    now = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    for e in events:
        d = e["date"].replace("-", "")
        st = e["start"].replace(":", "")
        en = e["end"].replace(":", "")
        uid = hashlib.md5((e["date"] + e["start"] + e["summary"]).encode("utf-8")).hexdigest()
        summary = (e["summary"] or e["text"])[:150].replace("\\", "").replace(";", ",").replace("\n", " ")
        desc = e["text"].replace("\\", "").replace("\n", " ")[:500]
        loc = e["room"]
        lines += [
            "BEGIN:VEVENT",
            "UID:%s@law.bsu.by" % uid,
            "DTSTAMP:%s" % now,
            "DTSTART;TZID=Europe/Minsk:%sT%s00" % (d, st),
            "DTEND;TZID=Europe/Minsk:%sT%s00" % (d, en),
            "SUMMARY:%s" % summary,
            "STATUS:CONFIRMED",
            "SEQUENCE:0",
        ]
        if loc:
            lines.append("LOCATION:Classroom %s" % loc)
        if desc:
            lines.append("DESCRIPTION:%s" % desc)
        lines.append("END:VEVENT")
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"


if __name__ == "__main__":
    path = download()
    if not path:
        raise SystemExit("no source file")
    events = extract(path)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(events, f, ensure_ascii=False, indent=2)
    with open(OUT_ICS, "w", encoding="utf-8", newline="") as f:
        f.write(to_ics(events))
    print("events=%d -> %s" % (len(events), OUT_ICS))
    for e in events:
        print(e["date"], e["start"], "-", e["end"], "|", e["summary"], "|", e["room"])
