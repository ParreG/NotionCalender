import os
import asyncio, re
from datetime import datetime, timedelta
import pytz, dateparser
from ics import Calendar, Event
from playwright.async_api import async_playwright

NOTION_URL = "https://qlok.notion.site/b4ab3c32788c4925a798a88b6a269676?v=2d2757914f0e42cfa43e4e6ce6fd0e6f"
ICS_FILENAME = "notion_calendar.ics"

TZ = pytz.timezone("Europe/Stockholm")
END_DATE = TZ.localize(datetime(2025, 12, 1))

COL_DATE  = "1"
COL_TIME  = "2"
COL_TITLE = "3"
COL_DESC  = "4"
COL_TCHR_A = "5"
COL_TCHR_B = "6"

def pick_first_line(s: str) -> str:
    if not s:
        return ""
    for part in s.split("\n"):
        p = part.strip()
        if p:
            return p
    return ""

def parse_date_sv(text):
    if not text:
        return None
    text = pick_first_line(text)
    dt = dateparser.parse(text, languages=["sv","en"], settings={
        "TIMEZONE": "Europe/Stockholm",
        "RETURN_AS_TIMEZONE_AWARE": True
    })
    if dt and dt.tzinfo is None:
        dt = TZ.localize(dt)
    return dt

def parse_time_range(s, base_date):
    if not s or not base_date:
        return None, None
    s = pick_first_line(s)
    m = re.search(r"(\d{1,2})(?:[:\.]?(\d{2}))?\s*[–-]\s*(\d{1,2})(?:[:\.]?(\d{2}))?", s)

    def ok_hm(h, mi):
        return 0 <= h <= 23 and 0 <= mi <= 59
    try:
        if not m:
            m1 = re.search(r"\b(\d{1,2})(?:[:\.]?(\d{2}))\b", s)
            if m1:
                h = int(m1.group(1)); mi = int(m1.group(2) or 0)
                if ok_hm(h, mi):
                    start = base_date.replace(hour=h, minute=mi, second=0, microsecond=0)
                    return start, start + timedelta(hours=1)
            return None, None

        h1 = int(m.group(1)); mi1 = int(m.group(2) or 0)
        h2 = int(m.group(3)); mi2 = int(m.group(4) or 0)

        if not ok_hm(h1, mi1) or not ok_hm(h2, mi2):
            return None, None

        start = base_date.replace(hour=h1, minute=mi1, second=0, microsecond=0)
        end   = base_date.replace(hour=h2, minute=mi2, second=0, microsecond=0)
        if end <= start:
            end += timedelta(days=1)
        return start, end
    except Exception:
        return None, None

async def switch_to_allt(page):
    try:
        await page.wait_for_selector("div[role='tablist']", timeout=15000)
        tabs = page.locator("div[role='tablist']").first.locator("[role='tab']")
        n = await tabs.count()
        for i in range(n):
            name = (await tabs.nth(i).inner_text() or "").strip()
            if "Allt" in name:
                await tabs.nth(i).click()
                await page.wait_for_timeout(600)
                return True
    except:
        pass
    return False

async def find_scroller(page):
    for sel in [
        ".notion-collection-view-body .notion-scroller.vertical.horizontal",
        ".notion-scroller.vertical.horizontal",
        ".notion-collection-view-body",
    ]:
        loc = page.locator(sel)
        if await loc.count() > 0:
            return loc.first
    return None

async def collect_all_rows(page):
    scroller = await find_scroller(page)
    if scroller is None:
        return []

    async def grab_cells():
        return await page.evaluate("""
        () => {
          const cells = Array.from(document.querySelectorAll('.notion-table-view-cell'));
          const rows = {};
          for (const el of cells) {
            const r = el.getAttribute('data-row-index');
            const c = el.getAttribute('data-col-index');
            if (!r || !c) continue;
            const t = (el.innerText || '').trim();
            if (!rows[r]) rows[r] = {};
            rows[r][c] = (rows[r][c] ? rows[r][c] + "\\n" : "") + t;
          }
          return rows;
        }
        """)

    try:
        await scroller.evaluate("(el)=>el.scrollTo(0, 0)")
        await page.wait_for_timeout(700)
    except:
        pass

    seen = {}
    stable = 0
    MAX_STABLE = 12
    MAX_SWEEPS = 1200

    for _ in range(MAX_SWEEPS):
        snap = await grab_cells()
        before = len(seen)
        for r, cols in snap.items():
            seen.setdefault(r, {})
            for c, txt in cols.items():
                seen[r][c] = seen[r].get(c, "") + ("\n" if c in seen[r] and seen[r][c] else "") + txt

        try:
            await scroller.evaluate("(el)=>el.scrollBy(0, el.clientHeight)")
        except:
            await page.mouse.wheel(0, 12000)
        await page.wait_for_timeout(700)

        after = len(seen)
        if after == before:
            stable += 1
        else:
            stable = 0
        if stable >= MAX_STABLE:
            break
    keys = sorted(seen.keys(), key=lambda k: int(k))
    return [seen[k] for k in keys]

async def main():
    cal = Calendar()
    added = 0
    miss_date = 0
    past_end = 0

    async with async_playwright() as p:
        headless = os.getenv("HEADLESS", "true").lower() == "true"
        browser = await p.chromium.launch(
            headless=headless,
            args=["--no-sandbox", "--disable-setuid-sandbox"]
        )
        page = await browser.new_page(viewport={"width": 1400, "height": 950})

        print("Öppnar sidan …")
        await page.goto(NOTION_URL, wait_until="domcontentloaded", timeout=120000)
        await page.wait_for_load_state("load")
        await switch_to_allt(page)
        await page.wait_for_timeout(500)

        rows = await collect_all_rows(page)
        print(f"Rader insamlade: {len(rows)}")

        for r in rows:
            date_text  = pick_first_line((r.get(COL_DATE)  or "").strip())
            time_text  = pick_first_line((r.get(COL_TIME)  or "").strip())
            title_text = pick_first_line((r.get(COL_TITLE) or "").strip())
            desc_text  = (r.get(COL_DESC)  or "").strip()

            tchr_a = pick_first_line((r.get(COL_TCHR_A) or "").strip())
            tchr_b = pick_first_line((r.get(COL_TCHR_B) or "").strip())
            teacher = tchr_a or tchr_b

            base_date = parse_date_sv(date_text)
            if not base_date:
                miss_date += 1
                continue

            if base_date > END_DATE:
                past_end += 1
                continue

            start_dt, end_dt = parse_time_range(time_text, base_date)
            if not start_dt:
                start_dt = base_date.replace(hour=9, minute=0, second=0, microsecond=0)
                end_dt = start_dt + timedelta(hours=1)

            title = title_text if title_text else base_date.strftime("Händelse %Y-%m-%d")

            meta = []
            if teacher:   meta.append(f"Lärare: {teacher}")
            if date_text: meta.append(f"Datum: {date_text}")
            if time_text: meta.append(f"Tider: {time_text}")
            description = "\n\n".join([x for x in [desc_text, "\n".join(meta)] if x])

            ev = Event()
            ev.name = title
            ev.begin = start_dt
            ev.end = end_dt
            ev.description = description
            cal.events.add(ev)
            added += 1

            print(f"📅 {start_dt.date()} {start_dt.strftime('%H:%M')}–{end_dt.strftime('%H:%M')} | {title}")

        with open(ICS_FILENAME, "w", encoding="utf-8") as f:
            f.write(cal.serialize())

        print(f"\nKLART. Events skrivna: {added}")
        print(f"Missade p.g.a. datum ej tolkade: {miss_date}")
        print(f"Uteslutna för att de låg efter {END_DATE.date()}: {past_end}")
        print(f"ICS-fil skapad: {ICS_FILENAME}")
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
