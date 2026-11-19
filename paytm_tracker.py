"""
FreshBus SRP tracker - Paytm Checkin, batch mode

Setup (one time):
    pip install playwright openpyxl
    playwright install chromium

Routes file (.xlsx or .csv):
    - a "route" column like "Bengaluru - Coimbatore", OR
    - "from" and "to" columns.
    City names must be spelled the way Paytm spells them (Bengaluru, not Bangalore).
    No links and no dates in the file - the link is built automatically:
        https://paytmcheckin.com/bus/search/{From}/{To}/{YYYY-MM-DD}/1

Run (the travel date is ALWAYS given on the command line):
    python paytm_tracker.py --routes paytm_routes.xlsx --date 2026-10-09
    python paytm_tracker.py --routes paytm_routes.xlsx --date 2026-10-09 2026-10-10
    python paytm_tracker.py --routes paytm_routes.xlsx --days-ahead 1
    python paytm_tracker.py --routes paytm_routes.xlsx --date 2026-11-04 --headed --debug

Position:
    Order of the bus cards on the Paytm results page, counting every bus card
    (sponsored "Ad" cards included, sold-out and departed buses included - these
    sit at the bottom so they don't shift anyone above them).
    The grouped government strip (e.g. "KSRTC ... 10 buses available") is not a
    bus card and is not counted.

Service key (Paytm):
    Each card's data-bus-id looks like 4251::4490::2026-10-14::64::913::onward_leg.
    The 5th part (913) is FreshBus's own service key (Paytm calls it
    provider_trip_id). It is saved as service_key; the full data-bus-id is
    saved as paytm_id.
"""

import argparse
import csv
import os
import random
import re
import time
from datetime import datetime, timedelta, timezone
from urllib.parse import quote

from playwright.sync_api import sync_playwright

IST = timezone(timedelta(hours=5, minutes=30))   # dates/times always in India time, even on a UTC server

SITE = "paytm"
BASE_URL = "https://paytmcheckin.com/bus/search/{frm}/{to}/{date}/1"
CARD_SEL = '[data-testid="srp-card"]'

FIELDS = ["run_time", "site", "route", "travel_date", "service_key", "paytm_id",
          "departure", "arrival",
          "travel_duration",
          "srp_position", "rating", "rating_count", "live_tracking",
          "price_original", "price_shown", "discount_text", "discount_pct",
          "services_count",
          # Paytm extras
          "total_buses", "status", "seats_left", "single_seats", "window_seats",
          "coupon_text", "sponsored"]

# Every Paytm bus card is <div data-testid="srp-card" data-operator=... data-departure=...
# data-sold-out=... data-departed=...> with labelled children (data-testid="srp-...").
JS_CARDS = r"""
() => {
  const q = (el, t) => el.querySelector(`[data-testid="${t}"]`);
  const txt = (el, t) => { const e = q(el, t); return e ? e.textContent.trim() : ''; };
  const attr = (el, t, a) => { const e = q(el, t); return e ? (e.getAttribute(a) || '') : ''; };
  // backup source: the card's hidden React data, item.provider_trip_id
  const tripId = el => {
    for (const k of Object.keys(el)) {
      if (!(k.startsWith('__reactFiber') || k.startsWith('__reactInternalInstance'))) continue;
      let f = el[k];
      for (let d = 0; d < 10 && f; d++) {
        const it = f.memoizedProps && f.memoizedProps.item;
        if (it && it.provider_trip_id != null) return String(it.provider_trip_id);
        f = f.return;
      }
    }
    return '';
  };
  return Array.from(document.querySelectorAll('[data-testid="srp-card"]')).map(c => {
    const r = q(c, 'srp-rating');
    let rc = '';
    if (r) {
      for (const s of r.querySelectorAll('span')) {
        const t = s.textContent.trim();
        if (/^\d[\d,]*$/.test(t)) rc = t;          // "36" (the rating itself has a dot, e.g. "4.1")
      }
    }
    return {
      operator: c.getAttribute('data-operator') || txt(c, 'srp-operator-name'),
      dep_attr: c.getAttribute('data-departure') || '',
      bus_id: c.getAttribute('data-bus-id') || '',
      trip_id: tripId(c),
      sold_out: c.getAttribute('data-sold-out') === 'true' || !!q(c, 'srp-sold-out'),
      departed: c.getAttribute('data-departed') === 'true' || !!q(c, 'srp-departed'),
      dep: txt(c, 'srp-departure-time'),
      arr: txt(c, 'srp-arrival-time'),
      dur: txt(c, 'srp-duration'),
      rating: r ? (r.getAttribute('data-rating') || '') : '',
      rating_count: rc,
      price: txt(c, 'srp-display-price'),
      strike: txt(c, 'srp-slasher-price'),
      disc: txt(c, 'srp-discount'),
      seats: attr(c, 'srp-seats-left', 'data-seats'),
      single: attr(c, 'srp-single-seat-count', 'data-single-seats'),
      window: attr(c, 'srp-window-seat-count', 'data-window-seats'),
      sale: txt(c, 'srp-sale'),
      ad: !!q(c, 'srp-sponsored'),
      live: !!q(c, 'srp-feature-live_tracking')
    };
  });
}
"""

# The results list scrolls inside #dwebSearchPage, not the window, so push both.
JS_SCROLL = r"""
() => {
  const cards = document.querySelectorAll('[data-testid="srp-card"]');
  if (cards.length) cards[cards.length - 1].scrollIntoView({block: 'end'});
  const box = document.getElementById('dwebSearchPage');
  if (box) box.scrollTop = box.scrollHeight;
  window.scrollBy(0, 4000);
  return cards.length;
}
"""

JS_FORM_CITIES = r"""
() => ({
  src: (document.getElementById('dwebSourceInput') || {}).value || '',
  dst: (document.getElementById('dwebDestinationInput') || {}).value || ''
})
"""


# ----------------------------------------------------------------- helpers --
def to_number(s):
    s = (s or "").replace(",", "").strip().upper()
    try:
        return int(float(s))
    except ValueError:
        return ""


DUR_RE = re.compile(r"(\d{1,3})\s*h(?:ours?|rs?)?\.?\s*(?:(\d{1,2})\s*m(?:in(?:ute)?s?)?)?", re.I)


def norm_duration(txt):
    """'06h 10m' becomes '6h 10m' (same format as the redBus/AbhiBus files)."""
    m = DUR_RE.search(txt or "")
    if not m:
        return ""
    return f"{int(m.group(1))}h {int(m.group(2) or 0)}m"


def to_24h(txt):
    """'11:50 PM' -> '23:50'. Returns the input unchanged if it can't be read."""
    t = " ".join((txt or "").split())
    for fmt in ("%I:%M %p", "%H:%M"):
        try:
            return datetime.strptime(t, fmt).strftime("%H:%M")
        except ValueError:
            continue
    return t


OFF_PCT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*%\s*OFF", re.I)
OFF_RS_RE = re.compile(r"((?:₹|Rs\.?)\s*[\d,]+)\s*OFF", re.I)
MONEY_RE = re.compile(r"(?:₹|Rs\.?)?\s*([\d,]+(?:\.\d+)?)", re.I)


def money(s):
    m = MONEY_RE.search(s or "")
    if not m:
        return ""
    v = float(m.group(1).replace(",", ""))
    return int(v) if v == int(v) else v


def price_fields(orig_txt, shown_txt, tag_txt):
    """orig = struck-through price, shown = price charged. Discount text from the
    '₹80 off' tag; discount % worked out from the two prices."""
    orig, shown = money(orig_txt), money(shown_txt)
    dtext, dpct = "", ""
    m = OFF_PCT_RE.search(tag_txt or "")
    if m:
        dtext, dpct = f"{m.group(1)}% OFF", float(m.group(1))
    else:
        m = OFF_RS_RE.search(tag_txt or "")
        if m:
            dtext = " ".join(m.group(1).split()) + " OFF"
    if dpct == "" and orig and shown and orig > shown:
        dpct = round((orig - shown) / orig * 100, 1)
    return {"price_original": orig, "price_shown": shown,
            "discount_text": dtext, "discount_pct": dpct}


def service_key_from(card):
    """5th part of data-bus-id (4251::4490::2026-10-14::64::913::onward_leg -> 913);
    falls back to the hidden provider_trip_id."""
    parts = (card.get("bus_id") or "").split("::")
    if len(parts) >= 5 and parts[4].strip():
        return parts[4].strip()
    return (card.get("trip_id") or "").strip()


def parse_card(card):
    dep = ""
    m = re.search(r"\b(\d{2}:\d{2}):\d{2}\b", card.get("dep_attr", ""))   # "2026-10-08 23:50:00"
    if m:
        dep = m.group(1)
    if not dep:
        dep = to_24h(card.get("dep", ""))
    if card.get("departed"):
        status = "departed"
    elif card.get("sold_out"):
        status = "sold out"
    else:
        status = "available"
    info = {
        "service_key": service_key_from(card),
        "paytm_id": (card.get("bus_id") or "").strip(),
        "operator": (card.get("operator") or "").strip(),
        "departure": dep,
        "arrival": to_24h(card.get("arr", "")),
        "duration": norm_duration(card.get("dur", "")),
        "rating": card.get("rating", ""),
        "rating_count": to_number(card.get("rating_count", "")),
        "live_tracking": "Yes" if card.get("live") else "No",
        "status": status,
        "seats_left": "" if status != "available" else to_number(card.get("seats", "")),
        "single_seats": to_number(card.get("single", "")),
        "window_seats": to_number(card.get("window", "")),
        "coupon_text": " ".join((card.get("sale") or "").split()),
        "sponsored": "Yes" if card.get("ad") else "No",
    }
    info.update(price_fields(card.get("strike", ""), card.get("price", ""), card.get("disc", "")))
    return info


def is_target(info, target_l):
    return target_l in info["operator"].lower().replace(" ", "")


# ---------------------------------------------------------------- scraping --
def load_all_cards(page, pause=1.2, max_rounds=40, patience=3):
    """Paytm seems to render the full list at once; this scrolls anyway in case it
    lazy-loads, and stops when the card count has not grown for a few rounds."""
    last, idle = -1, 0
    for _ in range(max_rounds):
        n = page.locator(CARD_SEL).count()
        if n == last:
            idle += 1
            if idle >= patience:
                break
        else:
            idle, last = 0, n
        page.evaluate(JS_SCROLL)
        time.sleep(pause)
    return max(last, 0)


def page_total(page):
    """The number in 'N buses found' (None if not shown)."""
    try:
        t = page.locator('[data-testid="srp-result-count"]').first.inner_text(timeout=3000)
        m = re.search(r"(\d+)", t)
        return int(m.group(1)) if m else None
    except Exception:
        return None


def save_debug(page, tag):
    os.makedirs("debug", exist_ok=True)
    try:
        page.screenshot(path=f"debug/paytm_{tag}.png", full_page=True)
    except Exception:
        pass
    with open(f"debug/paytm_{tag}.html", "w", encoding="utf-8") as f:
        f.write(page.content())
    print(f"   debug files saved: debug/paytm_{tag}.png and .html")


def scrape_route(ctx, url, frm, to, target, route_label, travel_date, run_time, debug_tag=None):
    page = ctx.new_page()
    target_l = target.lower().replace(" ", "")
    infos, header_total = [], None
    try:
        for attempt in range(3):          # retry temporary connection errors
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=60000)
                break
            except Exception as e:
                if attempt == 2:
                    raise
                print(f"   page failed to load ({str(e).splitlines()[0][:80]}); retrying...")
                time.sleep(5)
        try:
            page.wait_for_selector(CARD_SEL, timeout=30000)
        except Exception:
            pass
        time.sleep(1.5)

        # check Paytm understood the city names
        try:
            form = page.evaluate(JS_FORM_CITIES)
            if form["src"] and form["dst"] and (
                    form["src"].lower() != frm.lower() or form["dst"].lower() != to.lower()):
                print(f"   WARNING: page shows {form['src']} -> {form['dst']}, "
                      f"not {frm} -> {to} (check the city spelling)")
        except Exception:
            pass

        load_all_cards(page)
        header_total = page_total(page)
        infos = [parse_card(c) for c in page.evaluate(JS_CARDS)]
        for pos, info in enumerate(infos, start=1):
            info["position"] = pos

        if not infos and debug_tag is not None:
            save_debug(page, debug_tag)
    finally:
        page.close()

    matches = [i for i in infos if is_target(i, target_l)]
    rows = [{
        "run_time": run_time,
        "site": SITE,
        "route": route_label,
        "travel_date": travel_date,
        "service_key": i["service_key"],
        "paytm_id": i["paytm_id"],
        "departure": i["departure"],
        "arrival": i["arrival"],
        "travel_duration": i["duration"],
        "srp_position": i["position"],
        "rating": i["rating"],
        "rating_count": i["rating_count"],
        "live_tracking": i["live_tracking"],
        "price_original": i["price_original"],
        "price_shown": i["price_shown"],
        "discount_text": i["discount_text"],
        "discount_pct": i["discount_pct"],
        "services_count": len(matches),
        "total_buses": len(infos),
        "status": i["status"],
        "seats_left": i["seats_left"],
        "single_seats": i["single_seats"],
        "window_seats": i["window_seats"],
        "coupon_text": i["coupon_text"],
        "sponsored": i["sponsored"],
    } for i in matches]
    return len(infos), header_total, rows


# ------------------------------------------------------------------- files --
def out_file_for(date, out_arg=None):
    """--out given -> that file. Otherwise freshbus_paytm_DD-MM-YYYY.csv for the travel date."""
    if out_arg:
        return out_arg
    return "freshbus_paytm_" + datetime.strptime(date, "%Y-%m-%d").strftime("%d-%m-%Y") + ".csv"


def append_csv(rows, path):
    """Append rows. If the existing file has different columns, rewrite it once
    with the current columns first. Columns you added yourself are kept at the
    end (left blank for new rows)."""
    fields = list(FIELDS)
    if os.path.exists(path):
        with open(path, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            header = reader.fieldnames or []
            old = list(reader)
        fields += [h for h in header if h and h not in fields]
        if header and header != fields:
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
                w.writeheader()
                w.writerows(old)
    new_file = not os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        if new_file:
            w.writeheader()
        w.writerows(rows)


ROUTE_SPLIT = re.compile(r"\s+[-–→]\s+|\s*(?:→|–|->)\s*")


def split_route(text):
    parts = [p.strip() for p in ROUTE_SPLIT.split(text) if p.strip()]
    if len(parts) != 2:
        parts = [p.strip() for p in text.split("-") if p.strip()]
    return parts if len(parts) == 2 else None


def load_routes(path):
    if path.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).active
        rows = list(ws.iter_rows(values_only=True))
        header = [str(h or "").strip().lower() for h in rows[0]]
        raw = [dict(zip(header, r)) for r in rows[1:]]
    else:
        with open(path, newline="", encoding="utf-8-sig") as f:
            raw = [{(k or "").strip().lower(): v for k, v in r.items()}
                   for r in csv.DictReader(f)]

    out = []
    for n, r in enumerate(raw, start=2):
        frm = " ".join(str(r.get("from") or "").split())
        to = " ".join(str(r.get("to") or "").split())
        label = " ".join(str(r.get("route") or r.get("routes") or "").split())
        if not (frm and to) and label:
            parts = split_route(label)
            if not parts:
                print(f"row {n}: can't read route '{label}' (use 'From - To'); skipped")
                continue
            frm, to = parts
        if frm and to:
            out.append({"from": frm, "to": to, "route": f"{frm} - {to}"})
    return out


def build_url(frm, to, date):
    return BASE_URL.format(frm=quote(frm), to=quote(to), date=date)


# -------------------------------------------------------------------- main --
def main():
    ap = argparse.ArgumentParser(description="FreshBus positions on Paytm Checkin")
    ap.add_argument("--routes", default="paytm_routes.xlsx")
    ap.add_argument("--out", default=None,
                    help="Output CSV. Default: one file per travel date, "
                         "e.g. freshbus_paytm_14-10-2026.csv")
    ap.add_argument("--target", default="fresh", help="Operator text to match")
    ap.add_argument("--headed", action="store_true", help="Show the browser window")
    ap.add_argument("--debug", action="store_true",
                    help="Save screenshot + HTML (debug folder) when no bus cards are found")
    when = ap.add_mutually_exclusive_group(required=True)
    when.add_argument("--date", nargs="+", metavar="YYYY-MM-DD",
                      help="One or more travel dates for ALL routes")
    when.add_argument("--days-ahead", type=int,
                      help="Use today + N days (0 = today, 1 = tomorrow)")
    a = ap.parse_args()

    if a.date:
        for d in a.date:
            try:
                datetime.strptime(d, "%Y-%m-%d")
            except ValueError:
                raise SystemExit(f"Bad date '{d}'. Use YYYY-MM-DD, e.g. 2026-10-09")
        dates = a.date
    else:
        dates = [(datetime.now(IST) + timedelta(days=a.days_ahead)).strftime("%Y-%m-%d")]

    routes = load_routes(a.routes)
    if not routes:
        raise SystemExit(f"No routes found in {a.routes}")
    jobs = [(r, d) for r in routes for d in dates]
    run_time = datetime.now(IST).isoformat(timespec="seconds")
    files_written = set()
    total_rows = 0
    missing = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not a.headed)
        ctx = browser.new_context(
            viewport={"width": 1366, "height": 900},
            locale="en-IN",
            user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"),
        )
        for i, (r, date) in enumerate(jobs, start=1):
            url = build_url(r["from"], r["to"], date)
            print(f"[{i}/{len(jobs)}] {r['route']} on {date}")
            try:
                n, header_total, rows = scrape_route(
                    ctx, url, r["from"], r["to"], a.target, r["route"], date, run_time,
                    debug_tag=i if a.debug else None)
            except Exception as e:
                print(f"   skipped: {str(e).splitlines()[0][:120]}")
                if not a.headed:
                    print("   hint: if pages keep failing, try --headed")
                missing.append(f"{r['route']} ({date}): error")
                continue

            if n == 0:
                print("   no bus cards found (page may not have loaded, or no buses on this date)")
                missing.append(f"{r['route']} ({date}): no cards")
            elif not rows:
                print(f"   {n} buses read, FreshBus not listed")
                missing.append(f"{r['route']} ({date}): FreshBus not listed")
            else:
                if header_total and header_total != n:
                    print(f"   note: page says {header_total} buses found, {n} cards read")
                live = sum(1 for x in rows if x["status"] == "available")
                print(f"   FreshBus services: {len(rows)} ({live} bookable)")
                for row in rows:
                    price = row["price_shown"]
                    if row["price_original"]:
                        price = f"{price} (was {row['price_original']})"
                    print(f"   {row['departure']} -> {row['arrival']} ({row['travel_duration'] or '-'}) | "
                          f"service {row['service_key'] or '?'} | "
                          f"position {row['srp_position']} of {n} | "
                          f"rating {row['rating'] or '-'} ({row['rating_count'] or 0}) | "
                          f"Rs {price} | {row['status']}"
                          f"{' | seats ' + str(row['seats_left']) if row['seats_left'] != '' else ''}"
                          f"{' | AD' if row['sponsored'] == 'Yes' else ''}")
                if any(not r["service_key"] for r in rows):
                    print("   WARNING: service key not found for some FreshBus cards "
                          "(Paytm page layout may have changed) - service_key left blank")
                out_path = out_file_for(date, a.out)
                append_csv(rows, out_path)
                files_written.add(out_path)
                total_rows += len(rows)

            if i < len(jobs):
                time.sleep(random.uniform(2, 4))     # be gentle with the site
        browser.close()

    print(f"\nDone. {total_rows} FreshBus rows written to: "
          + (", ".join(sorted(files_written)) if files_written else "(no file - nothing found)"))
    if missing:
        print("Check these:")
        for m in missing:
            print(f"   {m}")


if __name__ == "__main__":
    main()
