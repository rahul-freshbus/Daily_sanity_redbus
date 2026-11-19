"""
FreshBus SRP tracker - redbus only, batch mode

Setup (one time):
    pip install playwright openpyxl
    playwright install chromium

Routes file (.xlsx or .csv), columns:  route, url
    - AbhiBus and redBus links can be mixed in the same file; the site is
      detected from the link.
    - Optional "date" column (YYYY-MM-DD) per row.
    - AbhiBus only: rows with from,to,date also work (needs cities.csv).

Run:
    python redbus_tracker.py --routes routes_links.xlsx --date 2026-10-07
    python redbus_tracker.py --date 2026-11-04 --headed --debug

Date priority: "date" column > --date > --days-ahead > the date inside the link.

Position:
    AbhiBus - order among private buses (cards with "Select Seats").
    redBus  - the "N of M" position redBus itself shows for each bus card.
              Government (RTC) buses are grouped in a separate strip on redBus
              and are not counted. Sponsored cards are counted.

Service key (redBus):
    redBus keeps a hidden data record behind every bus card. For FreshBus its
    "serviceId" is FreshBus's own service key (e.g. 912). It is saved in the
    service_key column; redbus_id is redBus's own number for the card.
"""

import argparse
import csv
import os
import random
import re
import time
from datetime import datetime, timedelta
from urllib.parse import quote

from playwright.sync_api import sync_playwright

# ---------------------------------------------------------------- AbhiBus ---
JS_GET_CARDS = r"""
(tgt) => {
  const leaves = Array.from(document.querySelectorAll('*')).filter(
    el => el.textContent.trim().toLowerCase() === 'select seats'
  );
  const amt = /^(₹|Rs\.?)\s*[\d,]+(\.\d+)?$/i;
  const trk = /live.?track|gps/i;
  const isStruck = (node, root) => {
    for (let x = node; x; x = x.parentElement) {
      if ((getComputedStyle(x).textDecorationLine || '').includes('line-through')) return true;
      if (x === root) break;
    }
    return false;
  };
  const seen = new Set();
  const out = [];
  for (const n of leaves) {
    let el = n;
    while (el.parentElement) {
      const p = el.parentElement;
      const c = (p.innerText.match(/select seats/gi) || []).length;
      if (c > 1) break;
      el = p;
    }
    if (seen.has(el)) continue;
    seen.add(el);
    const amounts = [];
    for (const d of el.querySelectorAll('*')) {
      if (d.children.length === 0 && amt.test(d.textContent.trim())) {
        amounts.push({t: d.textContent.trim(), s: isStruck(d, el)});
      }
    }
    let live = trk.test(el.innerText);
    // AbhiBus shows a "Tracking" chip (the label can be hidden by CSS, so read textContent)
    if (!live) {
      for (const sp of el.querySelectorAll('span')) {
        if (sp.children.length === 0 && /^\s*(live\s*)?tracking\s*$/i.test(sp.textContent)) { live = true; break; }
      }
    }
    // fallback: the chip's map-pin icon outline
    if (!live) live = !!el.querySelector('svg path[d^="M9.02399 5.49621C10.8083"]');
    if (!live) {
      outer: for (const d of el.querySelectorAll('*')) {
        for (const a of ['alt', 'title', 'aria-label', 'class', 'src', 'data-testid']) {
          const v = d.getAttribute(a);
          if (v && trk.test(v)) { live = true; break outer; }
        }
      }
    }
    const flat = el.innerText.toLowerCase().replace(/\s/g, '');
    const durEl = el.querySelector('.travel-time');
    out.push({text: el.innerText, amounts, live,
              dur: durEl ? durEl.textContent : '',
              html: (tgt && flat.includes(tgt)) ? el.outerHTML : ''});
  }
  return out;
}
"""

JS_COUNT = r"""
() => Array.from(document.querySelectorAll('*')).filter(
  el => el.textContent.trim().toLowerCase() === 'select seats'
).length
"""

TIME_LINE = re.compile(r"^([01]?\d|2[0-3]):[0-5]\d$")
BUS_TYPE_RE = re.compile(r"\b(AC|Non[\s-]?AC)\b.*\b(Seater|Sleeper)\b", re.I)
RATING_RE = re.compile(r"\b([1-5]\.\d)\b\s*([\d,]+(?:\.\d+)?\s*[KkMm]?)\b")
ABHI_DATE_IN_URL = re.compile(r"/(\d{4}-\d{2}-\d{2})/")

# ----------------------------------------------------------------- redBus ---
# Every redBus bus card is <li class="tupleWrapper..." aria-label="OPERATOR, type.
# Departs HH:MM, arrives HH:MM. ... Rated 4.6 out of 5. Total reviewers 126. ... 4 of 226">
JS_RB_CARDS = r"""
() => {
  // the card's hidden data record (React props): {serviceId, operatorId, routeId, ...}
  const inv = el => {
    const fk = Object.keys(el).find(k => k.startsWith('__reactFiber'));
    let f = fk ? el[fk] : null;
    for (let d = 0; d < 8 && f; d++) {
      if (f.memoizedProps && f.memoizedProps.inventory) return f.memoizedProps.inventory;
      f = f.return;
    }
    return null;
  };
  return Array.from(document.querySelectorAll('li[class*="tupleWrapper"][aria-label]')).map(e => {
   const v = inv(e) || {};
   return {
  label: e.getAttribute('aria-label'),
  rbid: e.id || (v.routeId != null ? String(v.routeId) : ''),
  sid: v.serviceId != null ? String(v.serviceId) : '',
  strike: (e.querySelector('[class*="strikeOffFare"]') || {}).textContent || '',
  final: (e.querySelector('[class*="finalFare"]') || {}).textContent || '',
  disc: (e.querySelector('[class*="discount_text"]') || {}).textContent || '',
  dur: (e.querySelector('p[class*="duration___"]') || {}).textContent || '',
  live: !!e.querySelector('[class*="liveTracking"]')
   };
  });
}
"""
RB_POS_RE = re.compile(r"(\d+)\s+of\s+(\d+)\s*$")
RB_DEP_RE = re.compile(r"Departs\s+(\d{1,2}:\d{2})")
RB_ARR_RE = re.compile(r"arrives\s+(\d{1,2}:\d{2})")
RB_RATING_RE = re.compile(r"Rated\s+([\d.]+)\s+out of 5\.\s*Total reviewers\s+([\d,]+)")

ONLY_SITE = "redbus"   # this copy only handles redbus links

FIELDS = ["run_time", "site", "route", "travel_date", "service_key", "redbus_id",
          "departure", "arrival",
          "travel_duration",
          "srp_position", "rating", "rating_count", "live_tracking",
          "price_original", "price_shown", "discount_text", "discount_pct",
          "services_count"]


def site_of(url):
    u = url.lower()
    if "redbus." in u:
        return "redbus"
    if "abhibus." in u:
        return "abhibus"
    raise ValueError("Unknown site in link (expected abhibus or redbus)")


def to_number(s):
    s = s.replace(",", "").strip().upper()
    mult = 1
    if s.endswith("K"):
        mult, s = 1_000, s[:-1]
    elif s.endswith("M"):
        mult, s = 1_000_000, s[:-1]
    try:
        return int(float(s) * mult)
    except ValueError:
        return ""


# ---- travel duration
DUR_RE = re.compile(r"(\d{1,3})\s*h(?:ours?|rs?)?\.?\s*(?:(\d{1,2})\s*m(?:in(?:ute)?s?)?)?", re.I)


def norm_duration(txt):
    """'06h.55m' (AbhiBus) and '6h 55m' (redBus) both become '6h 55m'."""
    m = DUR_RE.search(txt or "")
    if not m:
        return ""
    h, mi = int(m.group(1)), int(m.group(2) or 0)
    return f"{h}h {mi}m"


# ---- price / discount helpers
OFF_PCT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*%\s*OFF", re.I)
OFF_RS_RE = re.compile(r"((?:₹|Rs\.?)\s*[\d,]+)\s*OFF", re.I)
MONEY_RE = re.compile(r"(?:₹|Rs\.?)\s*([\d,]+(?:\.\d+)?)", re.I)


def money(s):
    m = MONEY_RE.search(s or "")
    if not m:
        return ""
    v = float(m.group(1).replace(",", ""))
    return int(v) if v == int(v) else v


def price_fields(orig_txt, shown_txt, tag_txt):
    """orig = struck-through price, shown = price being charged; discount from the
    'x% OFF' tag when present, otherwise worked out from the two prices."""
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


# ---- AbhiBus parsing
def operator_name(lines):
    for i, l in enumerate(lines):
        if BUS_TYPE_RE.search(l):
            j = i - 1
            while j >= 0 and lines[j].upper() == "AD":
                j -= 1
            return lines[j] if j >= 0 else ""
    return lines[0] if lines else ""


def parse_abhibus_card(card):
    text = card["text"]
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    times = [l for l in lines if TIME_LINE.match(l)]
    rating, count = "", ""
    m = RATING_RE.search(text)
    if m:
        rating = m.group(1)
        count = to_number(m.group(2))
    amounts = card["amounts"]
    orig_txt, shown_txt = "", ""
    for k, a in enumerate(amounts):
        if a["s"]:
            orig_txt = a["t"]
            after = [x["t"] for x in amounts[k + 1:] if not x["s"]]
            shown_txt = after[0] if after else ""
            break
    m = re.search(r"From\s*((?:₹|Rs\.?)\s*[\d,]+)", text, re.I)
    if m:
        shown_txt = m.group(1)                      # "From ₹1185" is the price charged
    elif not shown_txt:
        plain = [x["t"] for x in amounts if not x["s"]]
        shown_txt = plain[0] if plain else ""
    info = {
        "operator": operator_name(lines),
        "departure": times[0] if times else "",
        "arrival": times[1] if len(times) > 1 else "",
        "rating": rating,
        "rating_count": count,
        "live_tracking": "Yes" if card["live"] else "No",
        "duration": norm_duration(card.get("dur", "")),
        "text": text,
    }
    info.update(price_fields(orig_txt, shown_txt, ""))   # % worked out from the two prices
    sv = re.search(r"Save\s*((?:₹|Rs\.?)\s*[\d,]+)", text, re.I)
    if sv:
        info["discount_text"] = "Save " + " ".join(sv.group(1).split())
    return info


# ---- redBus parsing
def parse_redbus_label(label):
    head = label.split(". Departs")[0]          # "FRESHBUS, Electric A/C Seater/Sleeper (2+1)"
    operator = head.rsplit(", ", 1)[0] if ", " in head else head
    dep = RB_DEP_RE.search(label)
    arr = RB_ARR_RE.search(label)
    rt = RB_RATING_RE.search(label)
    pos = RB_POS_RE.search(label)
    return {
        "operator": operator.strip(),
        "departure": dep.group(1) if dep else "",
        "arrival": arr.group(1) if arr else "",
        "rating": rt.group(1) if rt else "",
        "rating_count": to_number(rt.group(2)) if rt else "",
        "position": int(pos.group(1)) if pos else None,
        "total": int(pos.group(2)) if pos else None,
        "text": label,
    }


def parse_redbus_card(card):
    label = card["label"]
    info = parse_redbus_label(label)
    info["service_key"] = (card.get("sid") or "").strip()    # FreshBus's own key, e.g. 912
    info["redbus_id"] = (card.get("rbid") or "").strip()     # redBus's number, e.g. 54406289
    live = card["live"] or "live tracking" in label.lower()
    info["live_tracking"] = "Yes" if live else "No"
    dm = re.search(r"Duration\s+([^.]+)\.", label)
    info["duration"] = norm_duration(card.get("dur", "")) or norm_duration(dm.group(1) if dm else "")
    shown_txt = card["final"]
    if not money(shown_txt):                       # fall back to "Price 390 INR" in the label
        m = re.search(r"Price\s+([\d,]+)\s*INR", label)
        shown_txt = m.group(1) if m else ""
        shown_txt = "₹" + shown_txt if shown_txt else ""
    tag = card["disc"] or label
    info.update(price_fields(card["strike"], shown_txt, tag))
    return info


# ---------------------------------------------------------------- scraping --
def is_target(site, info, target_l):
    if site == "redbus":
        hay = info["operator"]
    else:
        hay = info["operator"] + " " + info["text"][:200]
    return target_l in hay.lower().replace(" ", "")


def scroll_to_load_all(page, pause=1.5, max_rounds=20, stop_when=None):
    """Returns True if it stopped early because stop_when() said we have enough."""
    last = -1
    for _ in range(max_rounds):
        count = page.evaluate(JS_COUNT)
        if stop_when and stop_when():
            return True
        if count == last:
            break
        last = count
        page.mouse.wheel(0, 5000)
        time.sleep(pause)
    return False


JS_RB_SCROLL_LAST = r"""
() => {
  const cards = document.querySelectorAll('li[class*="tupleWrapper"]');
  if (!cards.length) return 0;
  const last = cards[cards.length - 1];
  last.scrollIntoView({block: 'end'});
  let el = last.parentElement;
  while (el) {                       // also push any inner scroll box to its bottom
    if (el.scrollHeight > el.clientHeight + 5 &&
        /(auto|scroll)/.test(getComputedStyle(el).overflowY)) {
      el.scrollTop = el.scrollHeight;
    }
    el = el.parentElement;
  }
  return cards.length;
}
"""


def collect_redbus(page, target_l="", expected=None, pause=1.8, max_rounds=150, patience=8):
    """Scroll to the last loaded card again and again (this is what triggers
    redBus to load the next batch), collecting cards as we go, de-duplicated
    by the position number redBus prints."""
    found = {}
    total = None
    idle = 0
    early = False
    page.mouse.move(683, 450)
    for _ in range(max_rounds):
        new = 0
        for card in page.evaluate(JS_RB_CARDS):
            info = parse_redbus_card(card)
            key = info["position"] if info["position"] else card["label"]
            if key not in found:
                found[key] = info
                new += 1
            if info["total"]:
                total = info["total"]
        if total and len(found) >= total:
            break
        if expected and sum(1 for i in found.values() if is_target("redbus", i, target_l)) >= expected:
            early = True
            break
        idle = 0 if new else idle + 1
        if idle >= patience:
            break
        page.evaluate(JS_RB_SCROLL_LAST)
        time.sleep(0.6)
        page.mouse.wheel(0, -300)      # small up/down wiggle re-fires lazy-load triggers
        time.sleep(0.3)
        page.mouse.wheel(0, 700)
        time.sleep(pause)
    cards = sorted(found.values(), key=lambda c: c["position"] or 10**9)
    return cards, (total or len(cards)), early


JS_RB_OPERATORS = r"""
() => Array.from(document.querySelectorAll('#travelsList [role="checkbox"]')).map(e => ({
  name: (e.querySelector('[class*="titleText"]') || {}).textContent || '',
  count: (e.querySelector('[class*="metaData"]') || {}).textContent || ''
}))
"""


def read_redbus_service_count(page, target):
    """Open the 'Bus operator' filter, type the operator name into its search box and
    read the number redBus shows next to it. Nothing is ticked, so the bus list
    (and its positions) is not changed. Returns None if it cannot be read."""
    try:
        header = page.locator('#travelsList [role="button"][aria-expanded]').first
        header.wait_for(timeout=6000)
        if header.get_attribute("aria-expanded") == "false":
            header.click()
        box = page.locator('#travelsList input:not([type="checkbox"])').first
        box.wait_for(timeout=6000)
        box.click()
        box.fill("")
        box.press_sequentially(target, delay=60)
        page.wait_for_timeout(1500)
        items = page.evaluate(JS_RB_OPERATORS)
    except Exception:
        return None
    tgt = target.lower().replace(" ", "")
    total, found = 0, False
    for it in items:
        if tgt in it["name"].lower().replace(" ", ""):
            digits = re.sub(r"[^\d]", "", it["count"])
            if digits:
                total += int(digits)
                found = True
    return total if found else None


# ---- AbhiBus: read the FreshBus count with the "Bus Partner" filter
SHOWING_RE = re.compile(r"Showing\s+(\d+)\s+Buses?", re.I)

JS_AB_MARK_PARTNER_BOX = r"""
() => {
  document.querySelectorAll('[data-srp-partner],[data-srp-hdr]').forEach(e => {
    e.removeAttribute('data-srp-partner'); e.removeAttribute('data-srp-hdr');
  });
  const w = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  let n;
  while ((n = w.nextNode())) {
    if (/^\s*Bus Partners?\s*$/i.test(n.nodeValue)) {
      const head = n.parentElement.closest('.collapsible-header') || n.parentElement;
      let el = n.parentElement;
      for (let k = 0; k < 8 && el; k++, el = el.parentElement) {
        const inp = el.querySelector('input:not([type=checkbox]):not([type=radio])');
        if (inp) {
          inp.setAttribute('data-srp-partner', '1');
          head.setAttribute('data-srp-hdr', '1');
          return true;
        }
      }
    }
  }
  return false;
}
"""

JS_AB_MARK_OPTIONS = r"""
(tgt) => {
  document.querySelectorAll('[data-srp-opt],[data-srp-lbl]').forEach(e => {
    e.removeAttribute('data-srp-opt'); e.removeAttribute('data-srp-lbl');
  });
  const box = document.querySelector('[data-srp-partner="1"]');
  if (!box) return 0;
  const bt = box.getBoundingClientRect().bottom;
  const maxX = window.innerWidth * 0.4;          // the filter panel is the left side
  const hits = [];
  for (const e of document.body.querySelectorAll('*')) {
    if (e === box) continue;
    const t = (e.textContent || '').trim();
    if (!t || t.length > 80) continue;
    if (!t.toLowerCase().replace(/\s/g, '').includes(tgt)) continue;
    const r = e.getBoundingClientRect();
    if (r.width === 0 || r.height === 0 || r.left > maxX || r.top < bt - 5) continue;
    hits.push(e);
  }
  const deepest = hits.filter(e => !hits.some(o => o !== e && e.contains(o)));
  deepest.forEach(e => {
    // the label is not linked to its checkbox, so mark the checkbox next to it
    const row = e.closest('.checkbox-container') || e.parentElement;
    const cb = row ? row.querySelector('input[type=checkbox]') : null;
    e.setAttribute('data-srp-lbl', '1');
    (cb || e).setAttribute('data-srp-opt', '1');
  });
  return deepest.length;
}
"""


def abhibus_total(page):
    """The number in 'Showing N Buses on this route' (None if not on the page)."""
    for getter in (lambda: page.locator(".buses-availability").first,
                   lambda: page.get_by_text(SHOWING_RE).first):
        try:
            m = SHOWING_RE.search(getter().inner_text(timeout=2000))
            if m:
                return int(m.group(1))
        except Exception:
            continue
    return None


def _wait_total(page, predicate, timeout_ms=7000):
    waited = 0
    while waited < timeout_ms:
        t = abhibus_total(page)
        if t is not None and predicate(t):
            page.wait_for_timeout(700)               # let it settle, then read again
            t2 = abhibus_total(page)
            return t2 if t2 is not None else t
        page.wait_for_timeout(300)
        waited += 300
    return None


def _reload_abhibus(page, url):
    page.goto(url, wait_until="domcontentloaded")
    try:
        page.wait_for_function(f"({JS_COUNT})() > 0", timeout=30000)
    except Exception:
        pass
    time.sleep(2)


def _click_options(page, tgt, n):
    for k in range(n):
        page.evaluate(JS_AB_MARK_OPTIONS, tgt)       # re-mark: the list may re-render
        try:
            page.locator('[data-srp-opt="1"]').nth(k).click(timeout=3000)
        except Exception:
            page.locator('[data-srp-lbl="1"]').nth(k).click(timeout=3000)
        page.wait_for_timeout(500)


def read_abhibus_service_count(page, target, url, dump_path=None):
    """Type the operator in 'Bus Partner', tick it, read 'Showing N Buses', untick.
    Returns N, or None if it could not be read (the caller then scans the full list).
    The page is always left showing the full, unfiltered list."""
    tgt = target.lower().replace(" ", "")
    before = abhibus_total(page)
    clicked = False
    count = None
    try:
        if before is None:
            raise RuntimeError("'Showing N Buses' text not found")
        if not page.evaluate(JS_AB_MARK_PARTNER_BOX):
            raise RuntimeError("Bus Partner search box not found")
        box = page.locator('[data-srp-partner="1"]').first
        if not box.is_visible():                      # section is collapsed: open it
            page.locator('[data-srp-hdr="1"]').first.click()
            page.wait_for_timeout(800)
        box.click()
        box.fill("")
        box.press_sequentially(target, delay=60)
        page.wait_for_timeout(1500)
        n = page.evaluate(JS_AB_MARK_OPTIONS, tgt)
        if not n:
            raise RuntimeError("no matching operator in the Bus Partner list")
        clicked = True
        _click_options(page, tgt, n)
        count = _wait_total(page, lambda t: t != before)
        if count is None:
            raise RuntimeError("the bus total did not change after ticking")
        _click_options(page, tgt, n)                 # untick
        back = _wait_total(page, lambda t: t == before)
        if back is None:
            raise RuntimeError("the full list did not come back")
        time.sleep(1.5)
        return count
    except Exception as e:
        print(f"   count via Bus Partner failed: {str(e)[:90]}")
        if dump_path:
            try:
                os.makedirs("debug", exist_ok=True)
                page.screenshot(path=dump_path + ".png")
                with open(dump_path + ".html", "w", encoding="utf-8") as f:
                    f.write(page.content())
                print(f"   debug files saved: {dump_path}.png and .html")
            except Exception:
                pass
        if clicked:                                   # make sure no filter is left on
            try:
                _reload_abhibus(page, url)
            except Exception:
                pass
        return None


def scrape_route(ctx, site, url, target, route_label, travel_date, run_time,
                 debug_tag=None, expected=None, auto_count=True):
    page = ctx.new_page()
    n_total = 0
    infos = []
    early = False
    target_l = target.lower().replace(" ", "")
    try:
        for attempt in range(3):          # retry temporary connection errors
            try:
                page.goto(url, wait_until="domcontentloaded")
                break
            except Exception as e:
                if attempt == 2:
                    raise
                print(f"   page failed to load ({str(e).splitlines()[0][:80]}); retrying...")
                time.sleep(5)
        if site != "redbus":          # redBus: the first-cards wait below is enough
            try:
                page.wait_for_load_state("networkidle", timeout=20000)
            except Exception:
                pass

        if site == "redbus":
            try:
                page.wait_for_selector('li[class*="tupleWrapper"]', timeout=30000)
            except Exception:
                pass
            time.sleep(1)
            if expected:
                print(f"   FreshBus services (from routes file): {expected}")
            elif auto_count:
                expected = read_redbus_service_count(page, target)
                if expected:
                    print(f"   FreshBus services on this date (from filter): {expected}")
                else:
                    print("   could not read the service count from the filter; scanning the full list")
            infos, n_total, early = collect_redbus(page, target_l, expected)
        else:
            try:
                page.wait_for_function(f"({JS_COUNT})() > 0", timeout=30000)
            except Exception:
                pass
            time.sleep(3)
            if expected:
                print(f"   FreshBus services (from routes file): {expected}")
            elif auto_count:
                dump = f"debug/abhibus_filter_{debug_tag}" if debug_tag is not None else None
                expected = read_abhibus_service_count(page, target, url, dump)
                if expected:
                    print(f"   FreshBus services on this date (from filter): {expected}")
                else:
                    print("   could not read the service count; scanning the full list")

            def enough():
                cards_now = page.evaluate(JS_GET_CARDS, "")
                return sum(1 for c in cards_now
                           if is_target("abhibus", parse_abhibus_card(c), target_l)) >= expected

            early = scroll_to_load_all(page, stop_when=enough if expected else None)
            tgt = target.lower().replace(" ", "") if debug_tag is not None else ""
            cards = page.evaluate(JS_GET_CARDS, tgt)
            if tgt:
                dumped = [c["html"] for c in cards if c.get("html")]
                if dumped:
                    os.makedirs("debug", exist_ok=True)
                    with open(f"debug/fresh_cards_{debug_tag}.html", "w", encoding="utf-8") as f:
                        f.write("\n\n<!-- ======== next card ======== -->\n\n".join(dumped))
                    print(f"   saved {len(dumped)} matching card(s): debug/fresh_cards_{debug_tag}.html")
            for pos, card in enumerate(cards, start=1):
                info = parse_abhibus_card(card)
                info["position"] = pos
                infos.append(info)
            n_total = None if early else len(cards)   # list size unknown when stopped early

        partial = site == "redbus" and infos and not early and len(infos) < n_total
        if (not infos or partial) and debug_tag is not None:
            os.makedirs("debug", exist_ok=True)
            tag = f"{debug_tag}_partial" if partial else str(debug_tag)
            page.screenshot(path=f"debug/route_{tag}.png", full_page=True)
            with open(f"debug/route_{tag}.html", "w", encoding="utf-8") as f:
                f.write(page.content())
            print(f"   debug files saved: debug/route_{tag}.png and .html")
    finally:
        page.close()

    rows = []
    for info in infos:
        if is_target(site, info, target_l):
            rows.append({
                "run_time": run_time,
                "site": site,
                "route": route_label,
                "travel_date": travel_date,
                "service_key": info.get("service_key", ""),
                "redbus_id": info.get("redbus_id", ""),
                "departure": info["departure"],
                "arrival": info["arrival"],
                "travel_duration": info.get("duration", ""),
                "srp_position": info["position"],
                "rating": info["rating"],
                "rating_count": info["rating_count"],
                "live_tracking": info["live_tracking"],
                "price_original": info["price_original"],
                "price_shown": info["price_shown"],
                "discount_text": info["discount_text"],
                "discount_pct": info["discount_pct"],
                "services_count": expected or "",
            })
    return n_total, len(infos), rows, early, expected


# ------------------------------------------------------------------- files --
def out_file_for(date, out_arg=None):
    """--out given -> that file. Otherwise freshbus_redbus_DD-MM-YYYY.csv for the travel date."""
    if out_arg:
        return out_arg
    return "freshbus_redbus_" + datetime.strptime(date, "%Y-%m-%d").strftime("%d-%m-%Y") + ".csv"


def append_csv(rows, path):
    """Append rows. If the existing file has different columns, rewrite it once
    with the current columns first. Columns you added yourself (e.g. a service
    number typed into the file) are kept at the end, left blank for new rows."""
    fields = list(FIELDS)
    if os.path.exists(path):
        with open(path, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            header = reader.fieldnames or []
            old = list(reader)
        fields += [h for h in header if h and h not in fields]
        if header and header != fields:
            for r in old:
                r.setdefault("site", "abhibus")
                if "redbus_id" not in header and r.get("service_key"):
                    # an earlier version wrote redBus's number into service_key
                    r["redbus_id"], r["service_key"] = r["service_key"], ""
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


def load_cities(path):
    cities = {}
    with open(path, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            name = r["city"].strip()
            cities[name.lower()] = (name, r["id"].strip())
    return cities


def build_abhibus_url(frm, to, date, cities):
    for c in (frm, to):
        if c.lower() not in cities:
            raise KeyError(f"City '{c}' not in cities.csv (add its AbhiBus ID)")
    f_name, f_id = cities[frm.lower()]
    t_name, t_id = cities[to.lower()]
    return (f"https://www.abhibus.com/bus_search/"
            f"{quote(f_name)}/{f_id}/{quote(t_name)}/{t_id}/{date}/O")


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

    def get(r, *names):
        for n in names:
            v = r.get(n)
            if v not in (None, ""):
                if hasattr(v, "strftime"):          # Excel real dates
                    return v.strftime("%Y-%m-%d")
                return str(v).strip()
        return ""

    out = []
    for r in raw:
        url = get(r, "url", "link", "links")
        frm, to = get(r, "from"), get(r, "to")
        if url or (frm and to):
            out.append({
                "route": " ".join(get(r, "route", "routes").split()),
                "url": url, "from": frm, "to": to, "date": get(r, "date"),
                "services": get(r, "services", "service", "fresh_count"),
            })
    return out


def resolve_dates(row, cli_dates, days_ahead):
    row_date = (row.get("date") or "").strip()
    if row_date:
        return [row_date]
    if cli_dates:
        return list(cli_dates)
    if days_ahead is not None:
        return [(datetime.now() + timedelta(days=days_ahead)).strftime("%Y-%m-%d")]
    return [None]


# ------------------------------------------------------------ date in URLs --
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def rb_fmt(iso):
    d = datetime.strptime(iso, "%Y-%m-%d")
    return f"{d.day:02d}-{MONTHS[d.month - 1]}-{d.year}"      # 07-Oct-2026


def rb_to_iso(s):
    try:
        d, m, y = s.split("-")
        return f"{int(y):04d}-{MONTHS.index(m.title()) + 1:02d}-{int(d):02d}"
    except Exception:
        return s


def redbus_set_date(url, iso):
    """Replace (or add) both onward= and doj= with the new date."""
    val = rb_fmt(iso)
    for key in ("onward", "doj"):
        if re.search(rf"([?&]){key}=[^&]*", url):
            url = re.sub(rf"([?&]){key}=[^&]*", rf"\g<1>{key}={val}", url)
        else:
            url += ("&" if "?" in url else "?") + f"{key}={val}"
    return url


def redbus_get_date(url):
    m = re.search(r"[?&]onward=([^&]+)", url) or re.search(r"[?&]doj=([^&]+)", url)
    return rb_to_iso(m.group(1)) if m else ""


# -------------------------------------------------------------------- main --
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--routes", default="redbus_routes.xlsx")
    ap.add_argument("--cities", default="cities.csv", help="AbhiBus from/to rows only")
    ap.add_argument("--out", default=None,
                    help="Output CSV. Default: one file per travel date, "
                         "e.g. freshbus_redbus_14-10-2026.csv")
    ap.add_argument("--target", default="fresh", help="Operator text to match")
    ap.add_argument("--headed", action="store_true")
    ap.add_argument("--no-auto-count", action="store_true",
                    help="redBus: do not read the FreshBus count from the filter")
    ap.add_argument("--full", action="store_true",
                    help="Ignore the 'services' column and always scan the whole list")
    ap.add_argument("--debug", action="store_true",
                    help="Save screenshot + HTML (debug folder) when no bus cards are found")
    ap.add_argument("--date", nargs="+", metavar="YYYY-MM-DD", default=None,
                    help="One or more travel dates for ALL routes")
    ap.add_argument("--days-ahead", type=int, default=None,
                    help="Use today + N days (0 = today, 1 = tomorrow)")
    a = ap.parse_args()

    if a.date:
        for d in a.date:
            try:
                datetime.strptime(d, "%Y-%m-%d")
            except ValueError:
                raise SystemExit(f"Bad date '{d}'. Use YYYY-MM-DD, e.g. 2026-10-07")

    cities = load_cities(a.cities) if os.path.exists(a.cities) else {}
    tomorrow = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")
    routes = load_routes(a.routes)
    jobs = [(r, d) for r in routes for d in resolve_dates(r, a.date, a.days_ahead)]

    run_time = datetime.now().isoformat(timespec="seconds")
    total_rows = 0
    files_written = set()
    locked_redirect = {}      # file open in Excel -> timestamped copy used instead

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not a.headed)
        ctx = browser.new_context(
            viewport={"width": 1366, "height": 900},
            locale="en-IN",
            user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"),
        )

        for i, (r, chosen_date) in enumerate(jobs, start=1):
            direct = (r.get("url") or "").strip()
            try:
                if direct:
                    site = site_of(direct)
                    if site == "redbus":
                        if chosen_date:
                            direct = redbus_set_date(direct, chosen_date)
                        date = redbus_get_date(direct)
                    else:
                        if chosen_date:
                            direct = ABHI_DATE_IN_URL.sub(f"/{chosen_date}/", direct, count=1)
                        m = ABHI_DATE_IN_URL.search(direct)
                        date = m.group(1) if m else ""
                    label = (r.get("route") or "").strip() or direct
                    url = direct
                else:
                    site = "abhibus"
                    frm, to = r["from"].strip(), r["to"].strip()
                    date = chosen_date or tomorrow
                    label = f"{frm} - {to}"
                    url = build_abhibus_url(frm, to, date, cities)
                if site != ONLY_SITE:
                    raise ValueError(f"{site} link, this script only handles {ONLY_SITE}")
                print(f"[{i}/{len(jobs)}] ({site}) {label} on {date}")
                expected = None
                sv = (r.get("services") or "").strip()
                if sv and not a.full:
                    try:
                        expected = int(float(sv)) or None
                    except ValueError:
                        print(f"   ignoring services value '{sv}' (not a number)")
                n, seen, rows, early, expected = scrape_route(
                    ctx, site, url, a.target, label, date, run_time,
                    debug_tag=i if a.debug else None, expected=expected,
                    auto_count=not (a.full or a.no_auto_count))
            except Exception as e:
                print(f"   skipped: {e}")
                if "ERR_HTTP2" in str(e) and not a.headed:
                    print("   hint: redBus often refuses headless browsers - run with --headed")
                continue

            if seen == 0:
                print("   no bus cards found (page may not have loaded)")
            elif not rows:
                print(f"   {seen} of {n} buses read, FreshBus not listed")
            else:
                of_txt = f" of {n}" if n else ""
                if early:
                    print(f"   stopped early: found {len(rows)} of {expected} expected FreshBus buses")
                elif n and seen < n:
                    print(f"   note: only {seen} of {n} cards loaded; positions are still correct")
                if expected and not early and len(rows) != expected:
                    print(f"   WARNING: expected {expected} FreshBus buses, found {len(rows)} "
                          f"- check the 'services' number for this route")
                for row in rows:
                    print(f"   {row['departure']} -> {row['arrival']} ({row['travel_duration'] or '-'}) | "
                          f"service {row['service_key'] or '?'} | "
                          f"position {row['srp_position']}{of_txt} | "
                          f"rating {row['rating']} ({row['rating_count']}) | "
                          f"live {row['live_tracking']} | "
                          f"price {row['price_original'] or '-'} -> {row['price_shown'] or '-'} | "
                          f"{row['discount_text'] or 'no discount tag'}")
                if any(not r["service_key"] for r in rows):
                    print("   WARNING: serviceId not found for some FreshBus cards "
                          "(redBus page layout may have changed) - service_key left blank")
                target = out_file_for(date, a.out)
                out_path = locked_redirect.get(target, target)
                try:
                    append_csv(rows, out_path)
                except PermissionError:
                    out_path = target.rsplit(".", 1)[0] + "_" + datetime.now().strftime("%H%M%S") + ".csv"
                    locked_redirect[target] = out_path
                    print(f"   WARNING: {target} is locked (open in Excel?). "
                          f"Saving the rest of this run to {out_path}")
                    append_csv(rows, out_path)
                files_written.add(out_path)
                total_rows += len(rows)

            if i < len(jobs):
                time.sleep(random.uniform(4, 8))

        browser.close()

    print(f"Done. {total_rows} FreshBus row(s) saved to: "
          + (", ".join(sorted(files_written)) if files_written else "(no file - nothing found)"))


if __name__ == "__main__":
    main()
