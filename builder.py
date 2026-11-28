"""
Near Here calendars - builder script v0.5 (Week 8)

Reads the Timetable Sheet (a local .xlsx file, or the Google Sheet's
"Publish to web" .xlsx link), checks every row, sets problem rows aside,
turns the dropdown choices into calendar rules and writes one calendar feed
per zone and audience group, plus an index page and a problems page.

Usage:
    python builder.py SHEET output_folder [YYYY-MM-DD]

SHEET is a path or an https link. The optional date is the build date
(defaults to today); calendars start from the first session on or after it.

New in v0.3 (Week 6):
  - reads the shared Google Sheet directly, so it can run every night
  - 'First date' column anchors fortnightly sessions, so nightly runs
    don't shift them to the wrong week
  - term-time calendar IDs stay the same from night to night
  - duplicate check also compares 'How often' ('1st' and '3rd Friday'
    rows are not duplicates)
  - venues must have a zone; venue postcodes are checked with postcodes.io
  - one set of feeds per zone, plus problems.html for the team
  - v0.3.1: monthly rules written as BYDAY=1FR / -1FR, which Google Calendar reads
  - v0.3.2: every series has an end date (end of the school year in the Term dates
    tab, or at least six months ahead). Google Calendar would not load the 28-session
    feed while most sessions repeated with no end. The nightly build moves it forward.

New in v0.4 (Week 7):
  - one page per area (areas/<area>.html) listing every live session by day, with
    'add to calendar' buttons for iPhone, Google Calendar and Outlook, a copy-link
    button, and a filter by audience
  - 'add just this one' calendar file for each session (sessions/<id>.ics)
  - the home page is now a postcode finder: the visitor types a postcode, the page
    asks postcodes.io where it is and points to the nearest area with calendars
  - area centres come from 'Anchor postcode' on the Zones tab (or, if blank, the
    middle of the area's venues); missing or unknown anchors are listed as warnings

New in v0.5 (Week 8):
  - printable A4 QR posters (posters/<area>.html), one for each area with live
    sessions, plus a general poster (posters/newquay.html) whose QR code opens
    the postcode finder; posters/index.html lists them for the team
  - QR codes are drawn as SVG by the 'segno' library; if it is missing the
    build still runs and simply skips the posters
  - the finder says 'under 1 km away' instead of 'about 0.0 km away'
"""
import csv
import html
import json
import re
import shutil
import sys
import tempfile
import math
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta, timezone
from difflib import SequenceMatcher
from pathlib import Path

from openpyxl import load_workbook

TZ = "Europe/London"
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
DAY_CODE = dict(zip(DAYS, ["MO", "TU", "WE", "TH", "FR", "SA", "SU"]))
PATTERN = {  # dropdown wording -> calendar rule (before BYDAY)
    "Every week": "FREQ=WEEKLY",
    "Every 2 weeks": "FREQ=WEEKLY;INTERVAL=2",
    "1st of the month": "FREQ=MONTHLY;BYSETPOS=1",
    "2nd of the month": "FREQ=MONTHLY;BYSETPOS=2",
    "3rd of the month": "FREQ=MONTHLY;BYSETPOS=3",
    "4th of the month": "FREQ=MONTHLY;BYSETPOS=4",
    "Last of the month": "FREQ=MONTHLY;BYSETPOS=-1",
}
FIRST_DATE = "First date (every 2 weeks only)"
REQUIRED = ["Session name", "Venue", "Day", "Start", "End", "How often", "Term-time only", "Audience", "Source link"]
STALE_DAYS = 60
MIN_HORIZON_DAYS = 183  # series always run at least this far ahead of the build date
TIME_RE = re.compile(r"^([01]\d|2[0-3]):([0-5]\d)$")
POSTCODE_RE = re.compile(r"^[A-Z]{1,2}\d[A-Z\d]? ?\d[A-Z]{2}$", re.I)
SITE = "https://voledx2.github.io/near-here-calendars/"
ROAD_COL = "If no sessions, send people to"


def slug(text):
    return re.sub(r"[^a-z0-9]+", "-", str(text).lower()).strip("-")


def ics_escape(text):
    return str(text).replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")


def fold(line):
    """iCalendar lines must be at most 75 octets; longer lines continue with a leading space."""
    out, b = [], line.encode("utf-8")
    while len(b) > 75:
        cut = 75 if not out else 74
        while (b[cut] & 0xC0) == 0x80:  # don't split a multi-byte character
            cut -= 1
        out.append(b[:cut].decode("utf-8"))
        b = b[cut:]
    out.append(b.decode("utf-8"))
    return "\r\n ".join(out)


def as_time(value):
    if value is None:
        return None
    if hasattr(value, "strftime"):
        return value.strftime("%H:%M")
    v = str(value).strip()
    return v if TIME_RE.match(v) else None


def as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d %b %Y"):
            try:
                return datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                pass
    return None


def read_table(ws):
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(h).strip() if h else "" for h in rows[0]]
    return [(i, dict(zip(hdr, r))) for i, r in enumerate(rows[1:], start=2) if r and r[0]]


def open_sheet(source):
    """Load the workbook from a local path or an https link (Google 'Publish to web' as .xlsx)."""
    if str(source).startswith("http"):
        req = urllib.request.Request(source, headers={"User-Agent": "near-here-builder"})
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = resp.read()
        tmp = Path(tempfile.mkdtemp()) / "sheet.xlsx"
        tmp.write_bytes(data)
        source = tmp
    return load_workbook(source, data_only=True)


def lookup_postcodes(postcodes):
    """Ask postcodes.io where each postcode is.
    Returns {postcode: (lat, lon)} with None for postcodes that don't exist, or None if offline."""
    if not postcodes:
        return {}
    try:
        found = {}
        todo = sorted(postcodes)
        for i in range(0, len(todo), 100):  # postcodes.io takes up to 100 at a time
            body = json.dumps({"postcodes": todo[i:i + 100]}).encode()
            req = urllib.request.Request("https://api.postcodes.io/postcodes", data=body,
                                         headers={"Content-Type": "application/json", "User-Agent": "near-here-builder"})
            with urllib.request.urlopen(req, timeout=30) as resp:
                for r in json.load(resp)["result"]:
                    res = r["result"]
                    found[r["query"]] = (res["latitude"], res["longitude"]) if res else None
        return found
    except Exception:
        return None


def first_on_or_after(start, weekday_name, rule, anchor=None):
    """First date on/after `start` that matches the day and pattern.
    For fortnightly sessions, `anchor` is a known session date that fixes which weeks it runs."""
    target = DAYS.index(weekday_name)
    if rule.startswith("FREQ=WEEKLY"):
        first = start + timedelta(days=(target - start.weekday()) % 7)
        if "INTERVAL=2" in rule and anchor:
            if (first - anchor).days % 14:
                first += timedelta(days=7)
        return first
    pos = int(rule.split("BYSETPOS=")[1])
    d = date(start.year, start.month, 1)
    for _ in range(14):  # look ahead up to 14 months
        matches = [d + timedelta(days=k) for k in range(31)
                   if (d + timedelta(days=k)).month == d.month and (d + timedelta(days=k)).weekday() == target]
        hit = matches[pos - 1] if pos > 0 else matches[pos]
        if hit >= start:
            return hit
        d = date(d.year + (d.month // 12), d.month % 12 + 1, 1)
    return None


def check(rows, venues, term_parts, today):
    """Return (ready_rows, problems). Each problem: (row, session, problem, fix)."""
    ready, problems, seen = [], [], {}
    for rownum, r in rows:
        name = r.get("Session name") or "(no name)"
        status = r.get("Status") or ""
        if status == "Paused":
            continue
        issues = []
        missing = [f for f in REQUIRED if not r.get(f)]
        if missing:
            issues.append(("Missing: " + ", ".join(missing), "Fill these in from the source, or ask the organiser"))
        start, end = as_time(r.get("Start")), as_time(r.get("End"))
        if r.get("Start") and not start:
            issues.append((f"Start time '{r.get('Start')}' isn't in 24-hour HH:MM format", "Write it like 10:00 or 15:30"))
        if r.get("End") and not end:
            issues.append((f"End time '{r.get('End')}' isn't in 24-hour HH:MM format", "Write it like 10:00 or 15:30"))
        if start and end and end <= start:
            issues.append(("End time is not after start time", "Check the times on the source"))
        often = r.get("How often")
        if often == "Monthly (week unknown)":
            issues.append(("Runs monthly but the source doesn't say which week", "Ask the venue which week, then pick the matching option"))
        elif often and often not in PATTERN:
            issues.append((f"Unknown 'How often' value '{often}'", "Pick a value from the dropdown"))
        anchor = as_date(r.get(FIRST_DATE))
        if often == "Every 2 weeks":
            if not anchor:
                issues.append(("Fortnightly session without a first date", f"Enter one real session date in '{FIRST_DATE}'"))
            elif r.get("Day") in DAYS and anchor.weekday() != DAYS.index(r["Day"]):
                issues.append((f"First date {anchor:%d/%m/%Y} is not a {r['Day']}", "Check the date or the day"))
        if r.get("Day") and r.get("Day") not in DAYS:
            issues.append((f"Unknown day '{r.get('Day')}'", "Pick a day from the dropdown"))
        venue = venues.get(r.get("Venue"))
        if r.get("Venue") and not venue:
            issues.append((f"Venue '{r.get('Venue')}' isn't on the Venues tab", "Add the venue with its postcode and zone"))
        elif venue and not venue.get("zone"):
            issues.append((f"Venue '{r.get('Venue')}' has no zone (outside our areas or not yet placed)",
                           "Give the venue a zone on the Venues tab, or leave it out of this proof of concept"))
        if r.get("Term-time only") == "Yes" and not term_parts:
            issues.append(("Term-time session but the Term dates tab is empty", "Fill in the Term dates tab"))
        bad_dates = [p.strip() for p in str(r.get("Cancelled dates") or "").split(",")
                     if p.strip() and not as_date(p.strip())]
        if bad_dates:
            issues.append(("Cancelled dates not in DD/MM/YYYY: " + ", ".join(bad_dates), "Rewrite as e.g. 25/12/2026"))
        verified = as_date(r.get("Last verified"))
        if not verified:
            issues.append(("No 'Last verified' date", "Check the source and enter today's date"))
        elif (today - verified).days > STALE_DAYS:
            issues.append((f"Not checked for {(today - verified).days} days", "Re-check the source and update the date"))
        # Possible duplicate: same venue, day, start time and pattern, and a similar name.
        key = (r.get("Venue"), r.get("Day"), start, often)
        for other_row, other_name in seen.get(key, []):
            if SequenceMatcher(None, slug(name), slug(other_name)).ratio() >= 0.6:
                issues.append((f"Looks like a duplicate of row {other_row} ('{other_name}')", "Delete one, or set the other to Paused"))
                break
        if start:
            seen.setdefault(key, []).append((rownum, name))
        if status == "Needs check" and not issues:
            issues.append(("Marked 'Needs check' by the person who entered it", "Confirm the details, then set Status to Ready"))
        if issues:
            problems += [(rownum, name, p, f) for p, f in issues]
        else:
            ready.append((rownum, r))
    return ready, problems


def vevents(r, venue, term_parts, build_date, stamp, horizon):
    """One VEVENT per session, or one per term part for term-time sessions."""
    day, rule0 = r["Day"], PATTERN[r["How often"]]
    if "BYSETPOS=" in rule0:
        # Write monthly rules as e.g. BYDAY=1FR or BYDAY=-1FR. Google Calendar did not
        # load feeds that used BYSETPOS (found in the Week 6 Android test).
        pos = rule0.split("BYSETPOS=")[1]
        rule = f"FREQ=MONTHLY;BYDAY={pos}{DAY_CODE[day]}"
    else:
        rule = f"{rule0};BYDAY={DAY_CODE[day]}"
    anchor = as_date(r.get(FIRST_DATE))
    start_t, end_t = as_time(r["Start"]).replace(":", "") + "00", as_time(r["End"]).replace(":", "") + "00"
    exdates = [as_date(p.strip()) for p in str(r.get("Cancelled dates") or "").split(",") if p.strip()]
    # Keep each term part's position in the full list, so calendar IDs stay the same night after night.
    windows = [(0, max(build_date, anchor or build_date), horizon)]
    if r["Term-time only"] == "Yes":
        windows = [(i, max(a, build_date), b) for i, (a, b) in enumerate(term_parts) if b >= build_date]
    location = ", ".join(x for x in (venue["name"], venue.get("address"), venue.get("postcode")) if x)
    notes = [f"Who it's for: {r['Audience']}", f"Cost: {r.get('Cost') or 'Not stated'}",
             f"How to join: {r.get('How to join') or 'Not stated'}"]
    if r["Term-time only"] == "Not stated":
        notes.append("The source doesn't say whether this runs in school holidays. Check with the venue.")
    verified = as_date(r.get("Last verified"))
    notes.append(f"Source: {r['Source link']}")
    notes.append(f"Last checked: {verified.strftime('%d %b %Y') if verified else 'unknown'}")
    out = []
    for i, a, b in windows:
        first = first_on_or_after(a, day, rule0, anchor)
        if first is None or (b and first > b):
            continue
        term = r["Term-time only"] == "Yes"
        uid = f"{slug(r['Venue'])}-{slug(r['Session name'])}-{DAY_CODE[day].lower()}{('-t' + str(i + 1)) if term else ''}@near-here"
        rr = rule + f";UNTIL={b.strftime('%Y%m%d')}T235959Z"
        lines = ["BEGIN:VEVENT", f"UID:{uid}", f"DTSTAMP:{stamp}",
                 f"SUMMARY:{ics_escape(r['Session name'])}",
                 f"DTSTART;TZID={TZ}:{first.strftime('%Y%m%d')}T{start_t}",
                 f"DTEND;TZID={TZ}:{first.strftime('%Y%m%d')}T{end_t}",
                 f"RRULE:{rr}"]
        for ex in exdates:
            if ex >= first and (not b or ex <= b):
                lines.append(f"EXDATE;TZID={TZ}:{ex.strftime('%Y%m%d')}T{start_t}")
        lines += [f"LOCATION:{ics_escape(location)}", f"DESCRIPTION:{ics_escape(chr(10).join(notes))}",
                  f"URL:{r['Source link']}", "END:VEVENT"]
        out.append(lines)
    return out


VTIMEZONE = """BEGIN:VTIMEZONE
TZID:Europe/London
BEGIN:DAYLIGHT
TZOFFSETFROM:+0000
TZOFFSETTO:+0100
TZNAME:BST
DTSTART:19700329T010000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:+0100
TZOFFSETTO:+0000
TZNAME:GMT
DTSTART:19701025T020000
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU
END:STANDARD
END:VTIMEZONE""".split("\n")


def write_feed(path, title, events):
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Near Here//Builder v0.5//EN", "CALSCALE:GREGORIAN",
             "METHOD:PUBLISH", f"X-WR-CALNAME:{ics_escape(title)}", f"X-WR-TIMEZONE:{TZ}",
             "REFRESH-INTERVAL;VALUE=DURATION:PT12H", "X-PUBLISHED-TTL:PT12H"] + VTIMEZONE
    for ev in events:
        lines += ev
    lines.append("END:VCALENDAR")
    path.write_text("\r\n".join(fold(l) for l in lines) + "\r\n", encoding="utf-8", newline="")


PAGE_STYLE = """
:root{--ink:#14232d;--muted:#52636e;--line:#d6e0e4;--bg:#f7faf9;--card:#fff;--accent:#0b6e5f;--accent-ink:#fff;--soft:#e4f2ee;--warn:#8a5a00}
*{box-sizing:border-box}
body{font-family:system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;margin:0;background:var(--bg);color:var(--ink);line-height:1.55}
main{max-width:860px;margin:0 auto;padding:24px 16px 48px}
h1{font-size:1.75rem;line-height:1.2;margin:.2rem 0 .6rem}
h2{font-size:1.2rem;margin:2rem 0 .6rem}
h3{font-size:1.05rem;margin:0 0 .2rem}
p{margin:.4rem 0 .8rem}
a{color:var(--accent)}
.lede{color:var(--muted);font-size:1.05rem}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;margin:0 0 12px}
.btns{display:flex;flex-wrap:wrap;gap:8px;margin:.6rem 0 .2rem}
.btn{display:inline-block;padding:10px 14px;border-radius:10px;border:1px solid var(--accent);background:var(--accent);color:var(--accent-ink);text-decoration:none;font-weight:600;font-size:.95rem;cursor:pointer;font-family:inherit}
.btn.alt{background:var(--card);color:var(--accent)}
.btn.small{padding:6px 10px;font-size:.85rem}
.meta{color:var(--muted);font-size:.92rem;margin:.1rem 0}
.tag{display:inline-block;background:var(--soft);border-radius:999px;padding:1px 10px;font-size:.82rem;margin-right:4px}
.note{color:var(--warn);font-size:.9rem}
label{font-weight:600;display:block;margin-bottom:6px}
input[type=text],select{font:inherit;padding:10px 12px;border:1px solid var(--line);border-radius:10px;width:100%;max-width:320px;background:#fff;color:var(--ink)}
.row{display:flex;flex-wrap:wrap;gap:8px;align-items:center}
.result{margin-top:12px}
.areas{list-style:none;padding:0;margin:0}
.areas li{margin:0 0 8px}
.small{font-size:.88rem;color:var(--muted)}
table{border-collapse:collapse;width:100%;font-size:14px}th,td{border:1px solid var(--line);padding:6px 8px;text-align:left;vertical-align:top}th{background:var(--soft)}
.wrap{overflow-x:auto}
footer{margin-top:2.5rem;border-top:1px solid var(--line);padding-top:12px}
"""


def page(title, inner, script=""):
    return (f'<!doctype html><html lang="en"><head><meta charset="utf-8">'
            f'<meta name="viewport" content="width=device-width, initial-scale=1"><title>{html.escape(title)}</title>'
            f'<style>{PAGE_STYLE}</style></head><body><main>{inner}</main>'
            f'{"<script>" + script + "</script>" if script else ""}</body></html>')


def subscribe_links(fname, title):
    """Links that add a feed in each calendar app. The feed must be online at SITE/feeds/fname."""
    https = f"{SITE}feeds/{fname}"
    webcal = "webcal://" + https.split("://", 1)[1]
    return {
        "apple": webcal,
        "google": "https://calendar.google.com/calendar/r?cid=" + urllib.parse.quote(webcal, safe=""),
        "outlook": "https://outlook.live.com/calendar/0/addfromweb?url=" + urllib.parse.quote(https, safe="")
                   + "&name=" + urllib.parse.quote(title, safe=""),
        "https": https,
    }


def how_often_words(r):
    day, often = r["Day"], r["How often"]
    words = {"Every week": f"Every {day}", "Every 2 weeks": f"Every other {day}",
             "1st of the month": f"1st {day} of the month", "2nd of the month": f"2nd {day} of the month",
             "3rd of the month": f"3rd {day} of the month", "4th of the month": f"4th {day} of the month",
             "Last of the month": f"Last {day} of the month"}[often]
    return words


def next_dates(r, term_parts, build_date, horizon, count=1):
    """The next session date(s) on or after the build date, skipping cancelled dates and school holidays."""
    rule0 = PATTERN[r["How often"]]
    anchor = as_date(r.get(FIRST_DATE))
    cancelled = {as_date(p.strip()) for p in str(r.get("Cancelled dates") or "").split(",") if p.strip()}
    windows = [(max(build_date, anchor or build_date), horizon)]
    if r["Term-time only"] == "Yes":
        windows = [(max(a, build_date), b) for a, b in term_parts if b >= build_date]
    out = []
    for a, b in windows:
        d = first_on_or_after(a, r["Day"], rule0, anchor)
        while d and d <= b and len(out) < count:
            if d not in cancelled:
                out.append(d)
            if rule0.startswith("FREQ=WEEKLY"):
                d += timedelta(days=14 if "INTERVAL=2" in rule0 else 7)
            else:
                nxt = date(d.year + (d.month // 12), d.month % 12 + 1, 1)
                d = first_on_or_after(nxt, r["Day"], rule0)
        if len(out) >= count:
            break
    return out


def area_page(zone, zone_feeds, sessions, built):
    """One page per area: subscribe buttons (with audience filter) and the sessions by day."""
    options, data = [], {}
    for aud, fname, _ in zone_feeds:
        label = "All sessions" if aud == "Everything" else aud
        n = sum(1 for s in sessions if aud in ("Everything", s["audience"]))
        options.append(f'<option value="{html.escape(fname)}">{html.escape(label)} ({n})</option>')
        data[fname] = subscribe_links(fname, f"Near Here - {zone} - {aud}")
    first = zone_feeds[0][1]
    links = data[first]
    cards = []
    for day in DAYS:
        todays = [s for s in sessions if s["day"] == day]
        if not todays:
            continue
        cards.append(f'<h2>{day}s</h2>')
        for s in sorted(todays, key=lambda x: x["start"]):
            when = f'{s["start"]}–{s["end"]} · {html.escape(s["often"])}'
            nxt = f' · next: {s["next"]}' if s["next"] else ""
            term = ""
            if s["term"] == "Yes":
                term = '<p class="meta">School term time only.</p>'
            elif s["term"] == "Not stated":
                term = '<p class="note">May not run in school holidays: check with the venue.</p>'
            place = html.escape(s["venue"])
            if s["postcode"]:
                place += (f', <a href="https://www.openstreetmap.org/search?query={urllib.parse.quote(s["postcode"])}">'
                          f'{html.escape(s["postcode"])}</a>')
            cards.append(f'''<div class="card" data-aud="{html.escape(s["audience"])}">
<h3>{html.escape(s["name"])}</h3>
<p class="meta">{when}{nxt}</p>
<p class="meta">{place}</p>
<p><span class="tag">{html.escape(s["audience"])}</span><span class="tag">{html.escape(s["cost"])}</span></p>
<p class="meta">How to join: {html.escape(s["join"])}</p>{term}
<div class="btns"><a class="btn small alt" href="../sessions/{s["file"]}">Add just this one</a>
<a class="btn small alt" href="{html.escape(s["source"])}" rel="noopener">Source (checked {s["checked"]})</a></div>
</div>''')
    inner = f'''<p class="small"><a href="../index.html">← Find another area</a></p>
<h1>{html.escape(zone)}: regular sessions</h1>
<p class="lede">{len(sessions)} regular sessions in the {html.escape(zone)} area of Newquay. Add them to your phone's calendar once and they stay up to date by themselves.</p>
<div class="card">
<label for="aud">Which sessions?</label>
<select id="aud">{"".join(options)}</select>
<div class="btns">
<a class="btn" id="b-apple" href="{links["apple"]}">iPhone or Mac</a>
<a class="btn" id="b-google" href="{links["google"]}" rel="noopener">Google Calendar</a>
<a class="btn" id="b-outlook" href="{links["outlook"]}" rel="noopener">Outlook</a>
<button class="btn alt" id="b-copy" type="button">Copy link</button>
</div>
<p class="small" id="copied" aria-live="polite">Android phones: use 'Google Calendar' on a computer or in your phone's browser, then turn the calendar on in the Google Calendar app. Times are UK time.</p>
</div>
{"".join(cards)}
<footer class="small">Rebuilt every night from the Near Here Timetable Sheet; last built {built}. Every session links to where its details came from. Spotted a mistake? Tell the venue or the Near Here team.</footer>'''
    script = f'''const L={json.dumps(data)};const sel=document.getElementById('aud');
function show(){{const f=sel.value,l=L[f];document.getElementById('b-apple').href=l.apple;document.getElementById('b-google').href=l.google;document.getElementById('b-outlook').href=l.outlook;
const aud=sel.options[sel.selectedIndex].text.replace(/ \\(\\d+\\)$/,'');document.querySelectorAll('.card[data-aud]').forEach(c=>{{c.hidden=!(aud==='All sessions'||c.dataset.aud===aud)}});
document.querySelectorAll('main h2').forEach(h=>{{let n=h.nextElementSibling,any=false;while(n&&n.tagName!=='H2'&&n.tagName!=='FOOTER'){{if(!n.hidden&&n.dataset.aud)any=true;n=n.nextElementSibling}}h.hidden=!any}})}}
sel.addEventListener('change',show);
document.getElementById('b-copy').addEventListener('click',async()=>{{const u=L[sel.value].https;const msg=document.getElementById('copied');
try{{await navigator.clipboard.writeText(u);msg.textContent='Link copied: '+u}}catch(e){{msg.textContent='Copy this link: '+u}}}});'''
    return page(f"Near Here - {zone}", inner, script)


def finder_page(zones_json, live, ready_n, blocked_n, built, has_test):
    items = "".join(f'<li><a class="btn alt" href="areas/{slug(z)}.html">{html.escape(z)} ({n} sessions)</a></li>'
                    for z, n in live)
    test = '<p class="small">Week 3 test feed: <a href="near_here_test_feed.ics">near_here_test_feed.ics</a></p>' if has_test else ""
    inner = f'''<h1>Near Here: regular sessions near you</h1>
<p class="lede">Clubs, groups and drop-ins that happen every week or month in Newquay, ready to add to your phone's calendar.</p>
<div class="card">
<form id="f" novalidate>
<label for="pc">Your postcode</label>
<div class="row"><input type="text" id="pc" name="pc" autocomplete="postal-code" placeholder="e.g. TR7 1BD" inputmode="text">
<button class="btn" type="submit">Find my area</button></div>
</form>
<div class="result" id="out" aria-live="polite"></div>
<p class="small">Your postcode is looked up with postcodes.io, a free public postcode service. This site does not keep it.</p>
</div>
<h2>Or pick an area</h2>
<ul class="areas">{items or '<li>No areas are live yet.</li>'}</ul>
<footer class="small">{ready_n} sessions are live. {blocked_n} more are waiting for missing details. Rebuilt every night from the team's Timetable Sheet; last built {built}.
<br>For the team: <a href="problems.html">problems list</a> · <a href="build_report.txt">build report</a> · <a href="posters/index.html">printable posters</a>{test}</footer>'''
    script = f'''const Z={json.dumps(zones_json)};
const out=document.getElementById('out');
function km(a,b,c,d){{const r=x=>x*Math.PI/180,R=6371,dl=r(c-a),dn=r(d-b);const h=Math.sin(dl/2)**2+Math.cos(r(a))*Math.cos(r(c))*Math.sin(dn/2)**2;return 2*R*Math.asin(Math.sqrt(h))}}
function say(h){{out.innerHTML=h}}
function esc(t){{return String(t).replace(/[&<>"]/g,c=>({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}})[c])}}
document.getElementById('f').addEventListener('submit',async e=>{{e.preventDefault();
const raw=document.getElementById('pc').value.toUpperCase().replace(/\\s+/g,'');
if(!/^[A-Z]{{1,2}}\\d[A-Z\\d]?(\\d[A-Z]{{2}})?$/.test(raw)){{say('<p>That doesn\\'t look like a UK postcode. Try something like <strong>TR7 1BD</strong>.</p>');return}}
const full=raw.length>4&&/\\d[A-Z]{{2}}$/.test(raw);
const url=full?'https://api.postcodes.io/postcodes/'+encodeURIComponent(raw):'https://api.postcodes.io/outcodes/'+encodeURIComponent(raw);
say('<p>Looking up '+esc(raw)+'…</p>');
let p;try{{const res=await fetch(url);if(res.status===404){{say('<p>We couldn\\'t find that postcode. Check it, or pick an area below.</p>');return}}p=(await res.json()).result}}catch(err){{say('<p>We couldn\\'t reach the postcode service just now. Please pick an area below.</p>');return}}
if(!p||p.latitude==null){{say('<p>That postcode has no location on record. Please pick an area below.</p>');return}}
const ranked=Z.filter(z=>z.lat!=null).map(z=>Object.assign({{d:km(p.latitude,p.longitude,z.lat,z.lon)}},z)).sort((a,b)=>a.d-b.d);
if(!ranked.length){{say('<p>Please pick an area below.</p>');return}}
const near=ranked[0];let live=ranked.filter(z=>z.live);
const road=live.find(z=>z.slug===near.road);if(!near.live&&road)live=[road].concat(live.filter(z=>z!==road));
if(near.d>15){{say('<p>That postcode is about '+Math.round(near.d)+' km from Newquay, outside the Near Here areas for now.</p>'+(live.length?'<p>The nearest area with calendars is <a class="btn" href="areas/'+live[0].slug+'.html">'+esc(live[0].name)+'</a></p>':''));return}}
const far=near.d<1?'under 1 km away':'about '+near.d.toFixed(1)+' km away';let h='<p>Your nearest area is <strong>'+esc(near.name)+'</strong> ('+far+').</p>';
if(near.live){{h+='<p><a class="btn" href="areas/'+near.slug+'.html">See '+near.sessions+' sessions in '+esc(near.name)+'</a></p>'}}
else if(live.length){{h+='<p>There are no sessions listed in '+esc(near.name)+' yet. The nearest area with sessions by road is '+esc(live[0].name)+'.</p><p><a class="btn" href="areas/'+live[0].slug+'.html">See '+live[0].sessions+' sessions in '+esc(live[0].name)+'</a></p>'}}
const shown=near.live?near.slug:(live[0]?live[0].slug:'');const more=live.filter(z=>z.slug!==near.slug&&z.slug!==shown).slice(0,2);
if(near.live&&more.length)h+='<p class="small">Also close by: '+more.map(z=>'<a href="areas/'+z.slug+'.html">'+esc(z.name)+'</a> ('+z.d.toFixed(1)+' km)').join(', ')+'. Distances are in a straight line.</p>';
say(h)}});'''
    return page("Near Here: regular sessions near you", inner, script)


POSTER_STYLE = """
@page{size:A4;margin:0}
*{box-sizing:border-box}
@media print{html,body{height:296mm;overflow:hidden}}
html,body{margin:0;background:#fff;color:#14232d;font-family:system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif}
.sheet{width:210mm;height:296mm;padding:16mm 16mm 30mm;position:relative;overflow:hidden;break-after:avoid}
.band{background:#0b6e5f;color:#fff;border-radius:6mm;padding:9mm 10mm 8mm}
.brand{font-size:12pt;letter-spacing:.14em;text-transform:uppercase;font-weight:700;opacity:.9;margin:0 0 3mm}
h1{font-size:34pt;line-height:1.05;margin:0}
.where{font-size:17pt;margin:3mm 0 0;font-weight:600}
.mid{display:flex;gap:10mm;align-items:center;margin:10mm 0 6mm}
.qr{flex:0 0 92mm}
.qr svg{width:92mm;height:92mm;display:block}
.steps{flex:1;font-size:15pt;line-height:1.35;margin:0;padding:0;list-style:none;counter-reset:s}
.steps li{counter-increment:s;margin:0 0 6mm;padding-left:13mm;position:relative}
.steps li::before{content:counter(s);position:absolute;left:0;top:-.5mm;width:9mm;height:9mm;border-radius:50%;background:#0b6e5f;color:#fff;font-weight:700;font-size:13pt;display:flex;align-items:center;justify-content:center}
.free{font-size:13pt;color:#52636e;margin:0 0 6mm}
.list{border-top:1.2mm solid #e4f2ee;padding-top:5mm;font-size:13.5pt;line-height:1.45}
.list h2{font-size:14pt;margin:0 0 2mm}
.list ul{margin:0;padding-left:6mm}
.foot{position:absolute;left:16mm;right:16mm;bottom:10mm;font-size:10.5pt;color:#52636e;border-top:.3mm solid #d6e0e4;padding-top:3mm}
.url{font-weight:700;color:#14232d;word-break:break-all}
@media screen{body{background:#e9eeec}.sheet{margin:10mm auto;background:#fff;box-shadow:0 1mm 6mm rgba(0,0,0,.15)}}
"""


def qr_svg(url):
    """The QR code as inline SVG, or None if the segno library isn't installed."""
    try:
        import segno
    except ImportError:
        return None
    return segno.make(url, error="m").svg_inline(dark="#14232d", border=4, omitsize=True)


def poster_page(title, where, url, qr, items, items_title, built):
    lis = "".join(f"<li>{html.escape(i)}</li>" for i in items)
    short = url.replace("https://", "")
    return (f'<!doctype html><html lang="en"><head><meta charset="utf-8">'
            f'<meta name="viewport" content="width=device-width, initial-scale=1"><title>{html.escape(title + " - " + where)}</title>'
            f'<style>{POSTER_STYLE}</style></head><body><div class="sheet">'
            f'<div class="band"><p class="brand">Near Here</p><h1>{html.escape(title)}</h1><p class="where">{html.escape(where)}</p></div>'
            f'<div class="mid"><div class="qr" role="img" aria-label="QR code for {html.escape(short)}">{qr}</div>'
            f'<ol class="steps"><li>Point your phone camera at the code.</li>'
            f'<li>Tap the button for your phone: iPhone or Google Calendar.</li>'
            f'<li>Done. New and changed sessions appear in your calendar by themselves.</li></ol></div>'
            f'<p class="free">Free. No app and no sign-up. You can remove it from your calendar at any time.</p>'
            f'<div class="list"><h2>{html.escape(items_title)}</h2><ul>{lis}</ul></div>'
            f'<div class="foot">No camera? Type <span class="url">{html.escape(short)}</span><br>'
            f'Details come from the organisers\' own listings and are checked by the Near Here team. Printed from the build of {built}.</div>'
            f'</div></body></html>')


def pick_examples(sessions, n=6):
    """A short, varied list of sessions for a poster: one per audience first, then the rest."""
    chosen, seen = [], set()
    for s in sessions:
        if s["audience"] not in seen:
            chosen.append(s); seen.add(s["audience"])
    for s in sessions:
        if s not in chosen:
            chosen.append(s)
    out = [f'{s["name"]}: {s["often"]}, {s["start"]}' for s in chosen[:n]]
    if len(sessions) > n:
        out.append(f"and {len(sessions) - n} more")
    return out


def write_posters(outdir, zones_live, session_info, built):
    """One A4 poster per live area plus a general one. Returns the number written, or 0 if QR codes can't be made."""
    if qr_svg(SITE) is None:
        return 0
    made = []
    for z in zones_live:
        url = f"{SITE}areas/{slug(z)}.html"
        ss = session_info[z]
        page_html = poster_page("Regular sessions near you", f"{z}, Newquay", url, qr_svg(url),
                                pick_examples(ss), f"{len(ss)} regular sessions, including:", built)
        (outdir / "posters" / f"{slug(z)}.html").write_text(page_html, encoding="utf-8")
        made.append((f"{slug(z)}.html", f"{z} ({len(ss)} sessions)"))
    items = [f"{z}: {len(session_info[z])} sessions" for z in zones_live]
    gen = poster_page("Regular clubs and groups near you", "Newquay and nearby villages", SITE, qr_svg(SITE),
                      items + ["Type your postcode to find your nearest area"], "Areas with sessions so far:", built)
    (outdir / "posters" / "newquay.html").write_text(gen, encoding="utf-8")
    made.append(("newquay.html", "General poster (opens the postcode finder)"))
    rows = "".join(f'<li><a href="{f}">{html.escape(t)}</a></li>' for f, t in made)
    (outdir / "posters" / "index.html").write_text(page("Near Here posters", f"""
<h1>Printable posters</h1>
<p>Open a poster and print it on A4 (in the print settings choose 'Fit to page' off, margins 'None' if offered). Each QR code opens the page named on the poster, so a printed poster never needs reprinting when sessions change.</p>
<ul>{rows}</ul>
<p class="small">Last built {built}. <a href="../index.html">Back to the postcode finder</a></p>"""), encoding="utf-8")
    return len(made)


def main():
    source, outdir = sys.argv[1], Path(sys.argv[2])
    build_date = date.fromisoformat(sys.argv[3]) if len(sys.argv) > 3 else date.today()
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    wb = open_sheet(source)
    rows = read_table(wb["Sessions"])
    venues = {}
    for _, v in read_table(wb["Venues"]):
        venues[str(v["Venue name"]).strip()] = {"name": str(v["Venue name"]).strip(), "address": v.get("Address"),
                                                "postcode": (v.get("Postcode") or "").strip().upper(),
                                                "zone": (v.get("Zone") or "").strip()}
    zone_rows = read_table(wb["Zones"])
    zones = [(str(z["Zone"]).strip(), (z.get("Anchor postcode") or "").strip().upper()) for _, z in zone_rows]
    # Straight-line distance can mislead (e.g. the River Gannel between Pentire and Crantock),
    # so each area can name the area people should be sent to by road when it has no sessions.
    by_road = {str(z["Zone"]).strip(): (z.get(ROAD_COL) or "").strip() for _, z in zone_rows}
    term_parts = [(as_date(t["Starts"]), as_date(t["Ends"])) for _, t in read_table(wb["Term dates"])
                  if as_date(t.get("Starts")) and as_date(t.get("Ends"))]

    ready, problems = check(rows, venues, term_parts, build_date)
    # Every series needs an end date: the end of the school year, or at least six months ahead.
    horizon = max([b for _, b in term_parts] + [build_date + timedelta(days=MIN_HORIZON_DAYS)])

    # Venue and area warnings (don't block publishing)
    warnings = []
    used = {r.get("Venue") for _, r in rows}
    to_check = {v["postcode"] for n, v in venues.items() if n in used and v["postcode"]} | {pc for _, pc in zones if pc}
    where = lookup_postcodes(to_check)
    for n, v in sorted(venues.items()):
        if n not in used:
            continue
        if not v["postcode"]:
            warnings.append((n, "No postcode", "Add the venue's postcode (shown on the area page and used to place the area)"))
        elif not POSTCODE_RE.match(v["postcode"]) or (where is not None and where.get(v["postcode"], 1) is None):
            warnings.append((n, f"Postcode '{v['postcode']}' not found", "Check the postcode"))
    postcode_note = "checked with postcodes.io" if where is not None else "not checked (postcodes.io unreachable)"

    # Area centres for the postcode finder: the anchor postcode, or the middle of the area's venues.
    zone_centre = {}
    for z, pc in zones:
        if pc and where and where.get(pc):
            zone_centre[z] = where[pc]
            continue
        if pc and where is not None:
            warnings.append((f"Area: {z}", f"Anchor postcode '{pc}' not found", "Fix the postcode on the Zones tab"))
        pts = [where[v["postcode"]] for v in venues.values()
               if where and v["zone"] == z and v["postcode"] and where.get(v["postcode"])]
        if pts:
            zone_centre[z] = (sum(a for a, _ in pts) / len(pts), sum(b for _, b in pts) / len(pts))
        if not pc:
            warnings.append((f"Area: {z}", "No anchor postcode", "Add a central postcode on the Zones tab for the postcode finder"))

    feeds = {}  # (zone, audience) -> events
    session_info = {}  # zone -> list of session dicts for the area page
    for d in ("feeds", "sessions", "areas", "posters"):
        if (outdir / d).exists():
            shutil.rmtree(outdir / d)  # drop pages and calendars that no longer have any sessions
        (outdir / d).mkdir(parents=True, exist_ok=True)
    for _, r in ready:
        v = venues[r["Venue"]]
        evs = vevents(r, v, term_parts, build_date, stamp, horizon)
        if not evs:
            continue
        for key in ((v["zone"], "Everything"), (v["zone"], r["Audience"])):
            feeds.setdefault(key, []).extend(evs)
        sid = f"{slug(r['Venue'])}-{slug(r['Session name'])}-{DAY_CODE[r['Day']].lower()}"
        write_feed(outdir / "sessions" / f"{sid}.ics", r["Session name"], evs)
        nd = next_dates(r, term_parts, build_date, horizon)
        verified = as_date(r.get("Last verified"))
        session_info.setdefault(v["zone"], []).append({
            "name": r["Session name"], "day": r["Day"], "start": as_time(r["Start"]), "end": as_time(r["End"]),
            "often": how_often_words(r), "next": nd[0].strftime("%a %d %b").replace(" 0", " ") if nd else "",
            "venue": v["name"], "postcode": v["postcode"], "audience": r["Audience"],
            "cost": r.get("Cost") if r.get("Cost") not in (None, "", "Not stated") else "Cost not stated", "join": r.get("How to join") or "Not stated",
            "term": r["Term-time only"], "source": r["Source link"], "file": f"{sid}.ics",
            "checked": verified.strftime("%d %b %Y") if verified else "unknown"})

    written = []
    for (zone, aud), evs in sorted(feeds.items(), key=lambda kv: (kv[0][0], kv[0][1] != "Everything", kv[0][1])):
        fname = f"{slug(zone)}-{slug(aud)}.ics"
        write_feed(outdir / "feeds" / fname, f"Near Here - {zone} - {aud}", evs)
        written.append((zone, aud, fname, len(evs)))

    with open(outdir / "problems.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Run date", "Sheet row", "Session", "Problem", "How to fix"])
        for p in problems:
            w.writerow([build_date.isoformat(), *p])

    built = build_date.strftime("%d %b %Y")
    zones_live = sorted({z for z, _, _, _ in written})
    for z in zones_live:
        zf = [(a, fn, n) for zz, a, fn, n in written if zz == z]
        (outdir / "areas" / f"{slug(z)}.html").write_text(area_page(z, zf, session_info[z], built), encoding="utf-8")

    posters = write_posters(outdir, zones_live, session_info, built)

    # Data quality summary
    active = [r for _, r in rows if r.get("Status") != "Paused"]
    blocked_rows = sorted({p[0] for p in problems})
    fields = ["Session name", "Venue", "Day", "Start", "End", "How often", "Audience", "Cost", "Source link", "Last verified"]
    filled = sum(1 for r in active for fld in fields if r.get(fld) not in (None, "", "Not stated"))
    completeness = filled / (len(active) * len(fields)) if active else 0
    term_known = sum(1 for r in active if r.get("Term-time only") in ("Yes", "No"))
    fresh = sum(1 for r in active if as_date(r.get("Last verified")) and (build_date - as_date(r.get("Last verified"))).days <= 30)
    sources = len({r.get("Source link") for r in active if r.get("Source link")})
    zone_names = [z for z, _ in zones] + sorted({v["zone"] for v in venues.values() if v["zone"]} - {z for z, _ in zones})
    zones_json = [{"name": z, "slug": slug(z), "lat": zone_centre.get(z, (None, None))[0],
                   "lon": zone_centre.get(z, (None, None))[1], "live": z in zones_live,
                   "sessions": len(session_info.get(z, [])), "road": slug(by_road.get(z, "")) if by_road.get(z) else ""}
                  for z in zone_names]
    placed = sum(1 for z in zones_json if z["lat"] is not None)
    summary = [
        f"Near Here builder v0.5 - build date {build_date.isoformat()}",
        f"Calendars run until: {horizon.strftime('%d %b %Y')}",
        f"Sessions in the sheet (not paused): {len(active)} from {sources} sources",
        f"Published to calendars: {len(ready)}",
        f"Held back for fixing: {len(blocked_rows)} ({len(problems)} problems)",
        f"Venue and area warnings: {len(warnings)} (postcodes {postcode_note})",
        f"Field completeness (10 key fields, 'Not stated' counts as missing): {completeness:.0%}",
        f"Sessions where term-time is known (Yes/No): {term_known} of {len(active)}",
        f"Sessions verified in the last 30 days: {fresh} of {len(active)}",
        f"Zones with live calendars: {', '.join(zones_live) or 'none'}",
        f"Areas placed on the map for the postcode finder: {placed} of {len(zones_json)}",
        f"Area pages written: {len(zones_live)}; single-session calendar files: {sum(len(v) for v in session_info.values())}",
        f"Printable QR posters: {posters}" + ("" if posters else " (segno not installed, posters skipped)"),
        "Feeds written:",
    ] + [f"  {fn}  ({n} calendar series)" for _, _, fn, n in written]
    (outdir / "build_report.txt").write_text("\n".join(summary) + "\n", encoding="utf-8")

    live = [(z, len(session_info[z])) for z in zones_live]
    (outdir / "index.html").write_text(
        finder_page(zones_json, live, len(ready), len(blocked_rows), built, (outdir / "near_here_test_feed.ics").exists()),
        encoding="utf-8")

    prow = "".join(f"<tr><td>{n}</td><td>{html.escape(s)}</td><td>{html.escape(p)}</td><td>{html.escape(f)}</td></tr>" for n, s, p, f in problems)
    wrow = "".join(f"<tr><td>{html.escape(n)}</td><td>{html.escape(p)}</td><td>{html.escape(f)}</td></tr>" for n, p, f in warnings)
    (outdir / "problems.html").write_text(page("Near Here problems list", f"""
<h1>Problems list</h1>
<p>Built {built}. Sessions below are kept off the calendars until the row is fixed in the Timetable Sheet. Fixes go live the next morning.</p>
<div class="wrap"><table><tr><th>Sheet row</th><th>Session</th><th>Problem</th><th>How to fix</th></tr>{prow or '<tr><td colspan="4">No problems</td></tr>'}</table></div>
<h2>Venue and area warnings</h2><p>These don't stop sessions being published. Postcodes {postcode_note}.</p>
<div class="wrap"><table><tr><th>Venue or area</th><th>Warning</th><th>How to fix</th></tr>{wrow or '<tr><td colspan="3">None</td></tr>'}</table></div>
<p><a href="index.html">Back to the postcode finder</a></p>"""), encoding="utf-8")
    print("\n".join(summary))


if __name__ == "__main__":
    main()
