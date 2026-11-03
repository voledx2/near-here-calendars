"""
Near Here calendars - builder script v0.2 (Week 5)

Reads the Timetable Sheet, checks every row, sets broken rows aside in a
problems report, turns the dropdown choices into calendar rules, and writes
one calendar feed per zone and audience group, plus a simple index page.

Usage:
    python builder.py Timetable_Sheet_v2.xlsx output_folder [YYYY-MM-DD]

The optional date is the "build date" (defaults to today). Calendars start
from the first session on or after it.
"""
import csv
from difflib import SequenceMatcher
import re
import sys
from datetime import date, datetime, timedelta, timezone
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
REQUIRED = ["Session name", "Venue", "Day", "Start", "End", "How often", "Term-time only", "Audience", "Source link"]
STALE_DAYS = 60
TIME_RE = re.compile(r"^([01]\d|2[0-3]):([0-5]\d)$")


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
    return None


def read_table(ws):
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(h).strip() if h else "" for h in rows[0]]
    return [(i, dict(zip(hdr, r))) for i, r in enumerate(rows[1:], start=2) if any(r)]


def first_on_or_after(start, weekday_name, rule):
    """First date on/after `start` that matches the day and pattern."""
    target = DAYS.index(weekday_name)
    if rule.startswith("FREQ=WEEKLY"):
        return start + timedelta(days=(target - start.weekday()) % 7)
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
        if r.get("How often") == "Monthly (week unknown)":
            issues.append(("Runs monthly but the source doesn't say which week", "Ask the venue which week, then pick the matching option"))
        elif r.get("How often") and r.get("How often") not in PATTERN:
            issues.append((f"Unknown 'How often' value '{r.get('How often')}'", "Pick a value from the dropdown"))
        if r.get("Day") and r.get("Day") not in DAYS:
            issues.append((f"Unknown day '{r.get('Day')}'", "Pick a day from the dropdown"))
        if r.get("Venue") and r.get("Venue") not in venues:
            issues.append((f"Venue '{r.get('Venue')}' isn't on the Venues tab", "Add the venue with its postcode and zone"))
        if r.get("Term-time only") == "Yes" and not term_parts:
            issues.append(("Term-time session but the Term dates tab is empty", "Fill in the Term dates tab"))
        bad_dates = []
        for part in str(r.get("Cancelled dates") or "").split(","):
            part = part.strip()
            if part:
                try:
                    datetime.strptime(part, "%d/%m/%Y")
                except ValueError:
                    bad_dates.append(part)
        if bad_dates:
            issues.append(("Cancelled dates not in DD/MM/YYYY: " + ", ".join(bad_dates), "Rewrite as e.g. 25/12/2026"))
        verified = as_date(r.get("Last verified"))
        if not verified:
            issues.append(("No 'Last verified' date", "Check the source and enter today's date"))
        elif (today - verified).days > STALE_DAYS:
            issues.append((f"Not checked for {(today - verified).days} days", "Re-check the source and update the date"))
        # Possible duplicate: same venue, day and start time AND a similar name.
        # (v0.1 compared times only, which wrongly flagged Chess Club and Board games,
        # two different Saturday 10:00 sessions at the library.)
        key = (r.get("Venue"), r.get("Day"), start)
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


def vevents(rownum, r, venue, term_parts, build_date, stamp):
    """One VEVENT per session, or one per term part for term-time sessions."""
    day, rule0 = r["Day"], PATTERN[r["How often"]]
    rule = f"{rule0};BYDAY={DAY_CODE[day]}"
    start_t, end_t = as_time(r["Start"]).replace(":", "") + "00", as_time(r["End"]).replace(":", "") + "00"
    exdates = []
    for part in str(r.get("Cancelled dates") or "").split(","):
        if part.strip():
            exdates.append(datetime.strptime(part.strip(), "%d/%m/%Y").date())
    windows = [(build_date, None)]
    if r["Term-time only"] == "Yes":
        windows = [(max(a, build_date), b) for a, b in term_parts if b >= build_date]
    location = f"{venue['name']}, {venue.get('address') or ''} {venue.get('postcode') or ''}".replace(" ,", ",").strip(", ")
    notes = [f"Who it's for: {r['Audience']}", f"Cost: {r.get('Cost') or 'Not stated'}", f"How to join: {r.get('How to join') or 'Not stated'}"]
    if r["Term-time only"] == "Not stated":
        notes.append("The source doesn't say whether this runs in school holidays. Check with the venue.")
    verified = as_date(r.get("Last verified"))
    notes.append(f"Source: {r['Source link']}")
    notes.append(f"Last checked: {verified.strftime('%d %b %Y') if verified else 'unknown'}")
    out = []
    for i, (a, b) in enumerate(windows):
        first = first_on_or_after(a, day, rule0)
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
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Near Here//Builder v0.2//EN", "CALSCALE:GREGORIAN",
             "METHOD:PUBLISH", f"X-WR-CALNAME:{ics_escape(title)}", f"X-WR-TIMEZONE:{TZ}",
             "REFRESH-INTERVAL;VALUE=DURATION:PT12H", "X-PUBLISHED-TTL:PT12H"] + VTIMEZONE
    for ev in events:
        lines += ev
    lines.append("END:VCALENDAR")
    path.write_text("\r\n".join(fold(l) for l in lines) + "\r\n", encoding="utf-8", newline="")


def main():
    sheet, outdir = Path(sys.argv[1]), Path(sys.argv[2])
    build_date = date.fromisoformat(sys.argv[3]) if len(sys.argv) > 3 else date.today()
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    wb = load_workbook(sheet, data_only=True)
    rows = read_table(wb["Sessions"])
    venues = {}
    for _, v in read_table(wb["Venues"]):
        if v.get("Venue name"):
            venues[v["Venue name"]] = {"name": v["Venue name"], "address": v.get("Address"),
                                       "postcode": v.get("Postcode"), "zone": v.get("Zone")}
    term_parts = [(as_date(t["Starts"]), as_date(t["Ends"])) for _, t in read_table(wb["Term dates"])
                  if as_date(t.get("Starts")) and as_date(t.get("Ends"))]

    ready, problems = check(rows, venues, term_parts, build_date)

    feeds = {}  # (zone, audience) -> events ; audience "Everything" for all
    for rownum, r in ready:
        v = venues[r["Venue"]]
        evs = vevents(rownum, r, v, term_parts, build_date, stamp)
        for key in ((v["zone"], "Everything"), (v["zone"], r["Audience"])):
            feeds.setdefault(key, []).extend(evs)

    (outdir / "feeds").mkdir(parents=True, exist_ok=True)
    written = []
    for (zone, aud), evs in sorted(feeds.items()):
        fname = f"{slug(zone)}-{slug(aud)}.ics"
        write_feed(outdir / "feeds" / fname, f"Near Here - {zone} - {aud}", evs)
        written.append((zone, aud, fname, len(evs)))

    with open(outdir / "problems.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Run date", "Sheet row", "Session", "Problem", "How to fix"])
        for p in problems:
            w.writerow([build_date.isoformat(), *p])

    # Data quality summary
    total = len([1 for _, r in rows if r.get("Status") != "Paused"])
    blocked_rows = sorted({p[0] for p in problems})
    fields = ["Session name", "Venue", "Day", "Start", "End", "How often", "Audience", "Cost", "Source link", "Last verified"]
    filled = sum(1 for _, r in rows for fld in fields
                 if r.get(fld) not in (None, "", "Not stated"))
    completeness = filled / (len(rows) * len(fields)) if rows else 0
    term_known = sum(1 for _, r in rows if r.get("Term-time only") in ("Yes", "No"))
    fresh = sum(1 for _, r in rows if as_date(r.get("Last verified")) and (build_date - as_date(r.get("Last verified"))).days <= 30)
    summary = [
        f"Near Here builder v0.2 - run for build date {build_date.isoformat()}",
        f"Rows read (not paused): {total}",
        f"Rows published to calendars: {len(ready)}",
        f"Rows held back for fixing: {len(blocked_rows)} ({len(problems)} problems)",
        f"Field completeness (10 key fields, 'Not stated' counts as missing): {completeness:.0%}",
        f"Rows where term-time is known (Yes/No): {term_known} of {len(rows)}",
        f"Rows verified in the last 30 days: {fresh} of {len(rows)}",
        "Feeds written:",
    ] + [f"  {fn}  ({n} calendar series)" for _, _, fn, n in written]
    (outdir / "build_report.txt").write_text("\n".join(summary) + "\n", encoding="utf-8")

    # Simple index page with subscribe links (full zone pages come in Week 7)
    items = "\n".join(f'<li><strong>{z} - {a}</strong>: <a href="feeds/{fn}">{fn}</a> ({n} series)</li>' for z, a, fn, n in written)
    (outdir / "index.html").write_text(f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>Near Here calendars</title>
<style>body{{font-family:system-ui,sans-serif;max-width:720px;margin:2rem auto;padding:0 16px;line-height:1.6;color:#152430}}code{{background:#eef3f5;padding:2px 5px;border-radius:4px}}</style></head>
<body><h1>Near Here calendars (proof of concept)</h1>
<p>Regular community sessions in Newquay, as calendars you can subscribe to. Built {build_date.strftime('%d %b %Y')} from the team's Timetable Sheet. Every session links to its source.</p>
<p>To subscribe: copy a link below, then in Google Calendar choose <em>Other calendars &gt; From URL</em>; on iPhone, open the link with <code>webcal://</code> in place of <code>https://</code>.</p>
<ul>{items}</ul>
<p>Test feed from Week 3: <a href="near_here_test_feed.ics">near_here_test_feed.ics</a></p>
</body></html>""", encoding="utf-8")
    print("\n".join(summary))


if __name__ == "__main__":
    main()
