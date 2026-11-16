"""
Near Here calendars - builder script v0.3.1 (Week 6)

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
"""
import csv
import html
import json
import re
import shutil
import sys
import tempfile
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
TIME_RE = re.compile(r"^([01]\d|2[0-3]):([0-5]\d)$")
POSTCODE_RE = re.compile(r"^[A-Z]{1,2}\d[A-Z\d]? ?\d[A-Z]{2}$", re.I)
SITE = "https://voledx2.github.io/near-here-calendars/"


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


def check_postcodes(postcodes):
    """Ask postcodes.io which postcodes exist. Returns {postcode: True/False}, or None if offline."""
    if not postcodes:
        return {}
    try:
        body = json.dumps({"postcodes": sorted(postcodes)}).encode()
        req = urllib.request.Request("https://api.postcodes.io/postcodes", data=body,
                                     headers={"Content-Type": "application/json", "User-Agent": "near-here-builder"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            res = json.load(resp)["result"]
        return {r["query"]: r["result"] is not None for r in res}
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


def vevents(r, venue, term_parts, build_date, stamp):
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
    windows = [(0, max(build_date, anchor or build_date), None)]
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
        uid = f"{slug(r['Venue'])}-{slug(r['Session name'])}-{DAY_CODE[day].lower()}{('-t' + str(i + 1)) if b else ''}@near-here"
        rr = rule + (f";UNTIL={b.strftime('%Y%m%d')}T235959Z" if b else "")
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
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Near Here//Builder v0.3.1//EN", "CALSCALE:GREGORIAN",
             "METHOD:PUBLISH", f"X-WR-CALNAME:{ics_escape(title)}", f"X-WR-TIMEZONE:{TZ}",
             "REFRESH-INTERVAL;VALUE=DURATION:PT12H", "X-PUBLISHED-TTL:PT12H"] + VTIMEZONE
    for ev in events:
        lines += ev
    lines.append("END:VCALENDAR")
    path.write_text("\r\n".join(fold(l) for l in lines) + "\r\n", encoding="utf-8", newline="")


PAGE_STYLE = ("body{font-family:system-ui,sans-serif;max-width:860px;margin:2rem auto;padding:0 16px;line-height:1.6;color:#152430}"
              "code{background:#eef3f5;padding:2px 5px;border-radius:4px}h2{margin-top:2rem}"
              "table{border-collapse:collapse;width:100%;font-size:14px}th,td{border:1px solid #d5dee2;padding:6px 8px;text-align:left;vertical-align:top}"
              "th{background:#e3f2ec}.wrap{overflow-x:auto}")


def page(title, inner):
    return (f'<!doctype html><html lang="en"><head><meta charset="utf-8">'
            f'<meta name="viewport" content="width=device-width, initial-scale=1"><title>{html.escape(title)}</title>'
            f'<style>{PAGE_STYLE}</style></head><body>{inner}</body></html>')


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
    term_parts = [(as_date(t["Starts"]), as_date(t["Ends"])) for _, t in read_table(wb["Term dates"])
                  if as_date(t.get("Starts")) and as_date(t.get("Ends"))]

    ready, problems = check(rows, venues, term_parts, build_date)

    # Venue warnings (don't block publishing, but the team should fix them before Week 7's postcode finder)
    warnings = []
    used = {r.get("Venue") for _, r in rows}
    to_check = {v["postcode"] for n, v in venues.items() if n in used and v["postcode"]}
    valid = check_postcodes(to_check)
    for n, v in sorted(venues.items()):
        if n not in used:
            continue
        if not v["postcode"]:
            warnings.append((n, "No postcode", "Add the venue's postcode (needed for the postcode finder)"))
        elif not POSTCODE_RE.match(v["postcode"]) or (valid is not None and not valid.get(v["postcode"], True)):
            warnings.append((n, f"Postcode '{v['postcode']}' not found", "Check the postcode"))
    postcode_note = "checked with postcodes.io" if valid is not None else "not checked (postcodes.io unreachable)"

    feeds = {}  # (zone, audience) -> events
    for _, r in ready:
        v = venues[r["Venue"]]
        evs = vevents(r, v, term_parts, build_date, stamp)
        for key in ((v["zone"], "Everything"), (v["zone"], r["Audience"])):
            feeds.setdefault(key, []).extend(evs)

    feeds_dir = outdir / "feeds"
    if feeds_dir.exists():
        shutil.rmtree(feeds_dir)  # drop calendars that no longer have any sessions
    feeds_dir.mkdir(parents=True, exist_ok=True)
    written = []
    for (zone, aud), evs in sorted(feeds.items(), key=lambda kv: (kv[0][0], kv[0][1] != "Everything", kv[0][1])):
        fname = f"{slug(zone)}-{slug(aud)}.ics"
        write_feed(feeds_dir / fname, f"Near Here - {zone} - {aud}", evs)
        written.append((zone, aud, fname, len(evs)))

    with open(outdir / "problems.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Run date", "Sheet row", "Session", "Problem", "How to fix"])
        for p in problems:
            w.writerow([build_date.isoformat(), *p])

    # Data quality summary
    active = [r for _, r in rows if r.get("Status") != "Paused"]
    blocked_rows = sorted({p[0] for p in problems})
    fields = ["Session name", "Venue", "Day", "Start", "End", "How often", "Audience", "Cost", "Source link", "Last verified"]
    filled = sum(1 for r in active for fld in fields if r.get(fld) not in (None, "", "Not stated"))
    completeness = filled / (len(active) * len(fields)) if active else 0
    term_known = sum(1 for r in active if r.get("Term-time only") in ("Yes", "No"))
    fresh = sum(1 for r in active if as_date(r.get("Last verified")) and (build_date - as_date(r.get("Last verified"))).days <= 30)
    sources = len({r.get("Source link") for r in active if r.get("Source link")})
    zones_live = sorted({z for z, _, _, _ in written})
    summary = [
        f"Near Here builder v0.3.1 - build date {build_date.isoformat()}",
        f"Sessions in the sheet (not paused): {len(active)} from {sources} sources",
        f"Published to calendars: {len(ready)}",
        f"Held back for fixing: {len(blocked_rows)} ({len(problems)} problems)",
        f"Venue warnings: {len(warnings)} (postcodes {postcode_note})",
        f"Field completeness (10 key fields, 'Not stated' counts as missing): {completeness:.0%}",
        f"Sessions where term-time is known (Yes/No): {term_known} of {len(active)}",
        f"Sessions verified in the last 30 days: {fresh} of {len(active)}",
        f"Zones with live calendars: {', '.join(zones_live) or 'none'}",
        "Feeds written:",
    ] + [f"  {fn}  ({n} calendar series)" for _, _, fn, n in written]
    (outdir / "build_report.txt").write_text("\n".join(summary) + "\n", encoding="utf-8")

    built = build_date.strftime("%d %b %Y")
    by_zone = {}
    for z, a, fn, n in written:
        by_zone.setdefault(z, []).append(f'<li><strong>{html.escape(a)}</strong>: <a href="feeds/{fn}">{fn}</a> ({n} series)</li>')
    zone_html = "".join(f"<h2>{html.escape(z)}</h2><ul>{''.join(items)}</ul>" for z, items in by_zone.items())
    test_link = '<p>Test feed from Week 3: <a href="near_here_test_feed.ics">near_here_test_feed.ics</a></p>' \
        if (outdir / "near_here_test_feed.ics").exists() else ""
    (outdir / "index.html").write_text(page("Near Here calendars", f"""
<h1>Near Here calendars (proof of concept)</h1>
<p>Regular community sessions in Newquay, as calendars you can subscribe to. Rebuilt automatically every night from the team's Timetable Sheet; last built {built}. Every session links to its source.</p>
<p><strong>{len(ready)} sessions</strong> are live in {len(zones_live)} areas. {len(blocked_rows)} more are waiting for missing details (<a href="problems.html">see the problems list</a>).</p>
<p>To subscribe: copy a link below, then in Google Calendar choose <em>Other calendars &gt; From URL</em>; on iPhone, open the link with <code>webcal://</code> in place of <code>https://</code>.</p>
{zone_html}{test_link}"""), encoding="utf-8")

    prow = "".join(f"<tr><td>{n}</td><td>{html.escape(s)}</td><td>{html.escape(p)}</td><td>{html.escape(f)}</td></tr>" for n, s, p, f in problems)
    wrow = "".join(f"<tr><td>{html.escape(n)}</td><td>{html.escape(p)}</td><td>{html.escape(f)}</td></tr>" for n, p, f in warnings)
    (outdir / "problems.html").write_text(page("Near Here problems list", f"""
<h1>Problems list</h1>
<p>Built {built}. Sessions below are kept off the calendars until the row is fixed in the Timetable Sheet. Fixes go live the next morning.</p>
<div class="wrap"><table><tr><th>Sheet row</th><th>Session</th><th>Problem</th><th>How to fix</th></tr>{prow or '<tr><td colspan="4">No problems</td></tr>'}</table></div>
<h2>Venue warnings</h2><p>These don't stop sessions being published, but need fixing before the postcode finder (Week 7). Postcodes {postcode_note}.</p>
<div class="wrap"><table><tr><th>Venue</th><th>Warning</th><th>How to fix</th></tr>{wrow or '<tr><td colspan="3">None</td></tr>'}</table></div>
<p><a href="index.html">Back to the calendars</a></p>"""), encoding="utf-8")
    print("\n".join(summary))


if __name__ == "__main__":
    main()
