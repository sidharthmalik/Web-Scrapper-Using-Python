import asyncio
import random
import re
import os
import pandas as pd
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright
from fake_useragent import UserAgent
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import httpx


# ─────────────────────────────────────────────────────────────
# TEXT CLEANER
# ─────────────────────────────────────────────────────────────
def clean_text(value: str) -> str:
    if not value or value == "N/A":
        return "N/A"
    value = value.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    value = re.sub(r"\s+", " ", value)
    value = value.strip()
    value = re.sub(r"[^\x20-\x7E\u0900-\u097F]", "", value)
    return value if value else "N/A"


# ─────────────────────────────────────────────────────────────
# MAIN SCRAPER CLASS
# ─────────────────────────────────────────────────────────────
class LeadScraper:
    def __init__(self):
        self.ua = UserAgent()
        self.leads = []

    async def init_browser(self, playwright):
        browser = await playwright.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-blink-features=AutomationControlled",
                "--disable-dev-shm-usage",
            ]
        )
        context = await browser.new_context(
            user_agent=self.ua.random,
            viewport={"width": 1280, "height": 800},
            extra_http_headers={
                "Accept-Language": "en-US,en;q=0.9",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            }
        )
        await context.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
            Object.defineProperty(navigator, 'plugins',   { get: () => [1, 2, 3] });
        """)
        return browser, context

    async def human_delay(self, min_sec=1.5, max_sec=4.0):
        await asyncio.sleep(random.uniform(min_sec, max_sec))

    async def human_scroll(self, page):
        height = await page.evaluate("document.body.scrollHeight")
        steps  = random.randint(3, 6)
        for i in range(steps):
            await page.evaluate(f"window.scrollTo(0, {(height / steps) * (i + 1)})")
            await asyncio.sleep(random.uniform(0.3, 0.8))

    # ─────────────────────────────────────────────────────────
    # SCRAPER 1: Yellow Pages
    # ─────────────────────────────────────────────────────────
    async def scrape_yellowpages(self, query: str, location: str, max_pages: int = 3):
        print(f"\n[Yellow Pages] Searching: '{query}' in '{location}'")
        async with async_playwright() as p:
            browser, context = await self.init_browser(p)
            page = await context.new_page()

            for page_num in range(1, max_pages + 1):
                url = (
                    f"https://www.yellowpages.com/search"
                    f"?search_terms={query}&geo_location={location}&page={page_num}"
                )
                print(f"  → Page {page_num}: {url}")
                try:
                    await page.goto(url, wait_until="domcontentloaded", timeout=30000)
                    await self.human_delay(2, 4)
                    await self.human_scroll(page)

                    soup     = BeautifulSoup(await page.content(), "html.parser")
                    listings = soup.select("div.result")

                    if not listings:
                        print(f"  → No more results on page {page_num}")
                        break

                    for listing in listings:
                        name    = listing.select_one("a.business-name")
                        phone   = listing.select_one("div.phones")
                        street  = listing.select_one("div.street-address")
                        city    = listing.select_one("div.locality")
                        website = listing.select_one("a.track-visit-website")
                        cat     = listing.select_one("div.categories")

                        lead = {
                            "source":        "yellowpages",
                            "business_name": name.text.strip()    if name    else "N/A",
                            "phone":         phone.text.strip()   if phone   else "N/A",
                            "address": (
                                f"{street.text.strip() if street else ''} "
                                f"{city.text.strip()   if city   else ''}"
                            ).strip(),
                            "website":  website.get("href", "N/A") if website else "N/A",
                            "category": cat.text.strip()           if cat     else "N/A",
                            "email":    "N/A",
                            "rating":   "N/A",
                        }
                        self.leads.append(lead)
                        print(f"     ✓ {lead['business_name']} | {lead['phone']}")

                    await self.human_delay(2, 5)

                except Exception as e:
                    print(f"  → Error on page {page_num}: {e}")
                    continue

            await browser.close()

    # ─────────────────────────────────────────────────────────
    # SCRAPER 2: Google Maps
    # ─────────────────────────────────────────────────────────
    async def scrape_google_maps(self, query: str, location: str, max_results: int = 20):
        print(f"\n[Google Maps] Searching: '{query}' near '{location}'")
        async with async_playwright() as p:
            browser, context = await self.init_browser(p)
            page = await context.new_page()

            url = f"https://www.google.com/maps/search/{(query + ' in ' + location).replace(' ', '+')}"

            try:
                await page.goto(url, wait_until="networkidle", timeout=30000)
                await self.human_delay(2, 4)

                results_scraped = 0
                previous_count  = 0

                while results_scraped < max_results:
                    listings = await page.query_selector_all(
                        "div[role='feed'] > div > div > a"
                    )

                    if len(listings) == previous_count:
                        print("  → No new results loading, stopping.")
                        break

                    for listing in listings[previous_count:]:
                        if results_scraped >= max_results:
                            break
                        try:
                            await listing.click()
                            await self.human_delay(1.5, 3)

                            name_el    = await page.query_selector("h1.DUwDvf")
                            phone_el   = await page.query_selector("button[data-item-id^='phone']")
                            address_el = await page.query_selector("button[data-item-id='address']")
                            website_el = await page.query_selector("a[data-item-id='authority']")
                            rating_el  = await page.query_selector("div.F7nice span")

                            lead = {
                                "source":        "google_maps",
                                "business_name": await name_el.inner_text()                  if name_el    else "N/A",
                                "phone":         await phone_el.get_attribute("data-tooltip") if phone_el   else "N/A",
                                "address":       await address_el.inner_text()               if address_el else "N/A",
                                "website":       await website_el.get_attribute("href")      if website_el else "N/A",
                                "rating":        await rating_el.inner_text()                if rating_el  else "N/A",
                                "email":         "N/A",
                                "category":      query,
                            }

                            if lead["business_name"] != "N/A":
                                self.leads.append(lead)
                                results_scraped += 1
                                print(f"     ✓ {lead['business_name']} | {lead['phone']} | {lead['address'][:40]}")

                        except Exception as e:
                            print(f"     → Skipped one listing: {e}")
                            continue

                    previous_count = len(listings)
                    feed = await page.query_selector("div[role='feed']")
                    if feed:
                        await feed.evaluate("el => el.scrollTop += 1000")
                    await self.human_delay(1.5, 3)

            except Exception as e:
                print(f"  → Fatal error: {e}")
            finally:
                await browser.close()

    # ─────────────────────────────────────────────────────────
    # EMAIL FINDER: Hunter.io API
    # ─────────────────────────────────────────────────────────
    async def find_email_via_hunter(self, domain: str, hunter_api_key: str) -> str:
        if not hunter_api_key or domain == "N/A":
            return "N/A"
        try:
            async with httpx.AsyncClient() as client:
                resp = await client.get(
                    "https://api.hunter.io/v2/domain-search",
                    params={"domain": domain, "api_key": hunter_api_key},
                    timeout=5.0
                )
                emails = resp.json().get("data", {}).get("emails", [])
                if emails:
                    best = sorted(emails, key=lambda x: x.get("confidence", 0), reverse=True)
                    return best[0].get("value", "N/A")
        except Exception:
            pass
        return "N/A"

    async def enrich_leads_with_emails(self, hunter_api_key: str):
        print(f"\n[Hunter.io] Enriching {len(self.leads)} leads with emails...")
        for i, lead in enumerate(self.leads):
            if lead.get("website") and lead["website"] != "N/A":
                domain = (
                    lead["website"]
                    .replace("https://", "")
                    .replace("http://", "")
                    .split("/")[0]
                )
                email = await self.find_email_via_hunter(domain, hunter_api_key)
                self.leads[i]["email"] = email
                if email != "N/A":
                    print(f"  ✓ {email}  ←  {lead['business_name']}")
            await asyncio.sleep(0.5)

    # ─────────────────────────────────────────────────────────
    # EXPORT — clean CSV + formatted Excel
    # ─────────────────────────────────────────────────────────
    def export(self, filename_base: str = None):
        if not self.leads:
            print("No leads to export.")
            return

        output_dir = os.path.expanduser("~/Desktop")
        base       = filename_base or f"leads_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}"
        csv_path   = os.path.join(output_dir, f"{base}.csv")
        excel_path = os.path.join(output_dir, f"{base}.xlsx")

        # ── Clean every field ─────────────────────────────────
        cleaned = [{k: clean_text(str(v)) for k, v in lead.items()} for lead in self.leads]
        df      = pd.DataFrame(cleaned)

        # ── Deduplicate ───────────────────────────────────────
        before = len(df)
        df     = df.drop_duplicates(subset=["business_name", "phone"])
        df     = df.reset_index(drop=True)

        # ── Column order ──────────────────────────────────────
        order = ["source","business_name","email","phone","address","website","category","rating"]
        df    = df[[c for c in order if c in df.columns]]

        # ── CSV export ────────────────────────────────────────
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")

        # ── Excel export ──────────────────────────────────────
        wb = Workbook()
        ws = wb.active
        ws.title = "Leads"

        col_labels = {
            "source": "Source", "business_name": "Business Name",
            "email": "Email",   "phone": "Phone",
            "address": "Address","website": "Website",
            "category": "Category", "rating": "Rating",
        }

        header_fill = PatternFill("solid", fgColor="1C2B3A")
        odd_fill    = PatternFill("solid", fgColor="F7F7F5")
        even_fill   = PatternFill("solid", fgColor="FFFFFF")
        header_font = Font(bold=True, color="C9A84C", size=10, name="Calibri")
        body_font   = Font(size=9, name="Calibri", color="0D0D0D")
        email_font  = Font(size=9, name="Calibri", color="1A7A4A", underline="single")
        c_align     = Alignment(horizontal="center", vertical="center", wrap_text=False)
        l_align     = Alignment(horizontal="left",   vertical="center", wrap_text=False)
        border      = Border(
            bottom=Side(style="thin", color="DDDDDD"),
            right =Side(style="thin", color="DDDDDD"),
        )

        columns = list(df.columns)

        # Header row
        ws.row_dimensions[1].height = 22
        for ci, col in enumerate(columns, 1):
            cell            = ws.cell(row=1, column=ci)
            cell.value      = col_labels.get(col, col.replace("_"," ").title())
            cell.font       = header_font
            cell.fill       = header_fill
            cell.alignment  = c_align
            cell.border     = border

        # Data rows
        for ri, row in df.iterrows():
            er = ri + 2
            ws.row_dimensions[er].height = 18
            fill = odd_fill if ri % 2 == 0 else even_fill
            for ci, col in enumerate(columns, 1):
                cell           = ws.cell(row=er, column=ci)
                cell.value     = row[col]
                cell.fill      = fill
                cell.border    = border
                if col == "email" and row[col] != "N/A":
                    cell.font      = email_font
                    cell.alignment = l_align
                elif col in ("rating", "source"):
                    cell.font      = body_font
                    cell.alignment = c_align
                else:
                    cell.font      = body_font
                    cell.alignment = l_align

        # Auto-fit column widths
        for ci, col in enumerate(columns, 1):
            header_len  = len(col_labels.get(col, col)) + 4
            content_len = df[col].astype(str).str.len().max() if len(df) > 0 else 0
            width       = min(50, max(12, header_len, content_len + 2))
            if col in ("address", "website"):
                width = min(45, width)
            ws.column_dimensions[get_column_letter(ci)].width = width

        ws.freeze_panes    = "A2"
        ws.auto_filter.ref = ws.dimensions
        wb.save(excel_path)

        # ── Summary ───────────────────────────────────────────
        has_email = len(df[df["email"] != "N/A"]) if "email" in df.columns else 0
        print(f"\n{'='*52}")
        print(f"  ✅  {len(df)} leads  ({before - len(df)} duplicates removed)")
        print(f"  📧  {has_email} leads with verified emails")
        print(f"  📊  Excel → {excel_path}")
        print(f"  📄  CSV   → {csv_path}")
        print(f"{'='*52}")
        return excel_path, csv_path


# ─────────────────────────────────────────────────────────────
# CONFIGURE AND RUN
# ─────────────────────────────────────────────────────────────
async def main():
    scraper = LeadScraper()

    # ── Change these 4 lines to control what gets scraped ────
    SEARCH_QUERY    = "digital marketing agency"
    SEARCH_LOCATION = "Gurugram"
    MAX_RESULTS     = 30
    HUNTER_API_KEY  = ""       # Free key at hunter.io (25 lookups/month free)
    # ─────────────────────────────────────────────────────────

    await scraper.scrape_google_maps(SEARCH_QUERY, SEARCH_LOCATION, MAX_RESULTS)
    await scraper.scrape_yellowpages(SEARCH_QUERY, SEARCH_LOCATION, max_pages=3)

    if HUNTER_API_KEY:
        await scraper.enrich_leads_with_emails(HUNTER_API_KEY)

    scraper.export(f"{SEARCH_QUERY.replace(' ', '_')}_{SEARCH_LOCATION}")


if __name__ == "__main__":
    asyncio.run(main())