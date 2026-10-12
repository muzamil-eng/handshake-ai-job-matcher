"""
scraper.py - Module 1: Handshake browser automation & scraper
==============================================================

Two commands:

    python scraper.py login     # ONE-TIME: log in by hand, saves session to auth.json
    python scraper.py scrape    # headless scrape using the saved session

Setup:
    pip install playwright pandas openpyxl
    playwright install chromium

How login works (handles SSO + 2FA):
    1. `python scraper.py login` opens a visible Chromium window.
    2. YOU log in normally (school SSO, Duo/2FA, etc.). The script never sees
       or stores your password.
    3. When you land on your Handshake home page, return to the terminal and
       press Enter. Playwright saves cookies + localStorage to auth.json.
    4. All later `scrape` runs reuse auth.json. If Handshake expires the
       session, the scraper detects the login redirect and tells you to
       re-run `login`.
    Keep auth.json private (add it to .gitignore) - it is equivalent to
    being logged in as you.

IMPORTANT NOTES:
    * Handshake's Terms of Service restrict automated access. This script is
      for low-volume, personal use on your own account: it is rate-limited,
      capped by --max-pages / --max-jobs, and you are responsible for
      complying with their terms.
    * Handshake changes its front-end often. Every selector and URL pattern
      lives in the CONFIG block below so you can fix them in one place. Run
      with --headed --debug the first time to watch what it does, and use
      `playwright codegen --load-storage=auth.json https://app.joinhandshake.com`
      to find current selectors if something returns blank.
"""

from __future__ import annotations

import argparse
import random
import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import urlencode

import pandas as pd
from playwright.sync_api import Page, TimeoutError as PWTimeout, sync_playwright

# ----------------------------------------------------------------------------
# CONFIG - edit these, not the logic below
# ----------------------------------------------------------------------------
BASE_URL = "https://app.joinhandshake.com"
LOGIN_URL = f"{BASE_URL}/login"
# Search results URL. If Handshake changes it, run a search by hand and copy
# the path/param names from the address bar.
SEARCH_PATH = "/stu/jobs"
QUERY_PARAM = "query"
PAGE_PARAM = "page"

AUTH_FILE = Path("auth.json")
OUTPUT_FILE = Path("handshake_jobs.xlsx")

DEFAULT_KEYWORDS = ["Civil Engineering", "Project Management", "Internship"]
DEFAULT_LOCATIONS = ["Beltsville", "Baltimore", "Washington, DC", "DC", "Laurel"]

# Selector lists are tried in order until one returns text.
SELECTORS = {
    "title": ["h1", "[data-hook='job-title']"],
    "company": ["a[href*='/employers/']", "[data-hook='employer-name']"],
    "location": [
        "[data-hook='job-location']",
        "[data-hook='location']",
        "xpath=//*[contains(@class,'location')][1]",
    ],
    "description": [
        "[data-hook='details-content']",
        "[data-hook='job-description']",
        "main",
    ],
}
# Links on a results page that point at a job detail page.
JOB_LINK_SELECTOR = "a[href*='/jobs/']"
JOB_ID_RE = re.compile(r"/jobs/(\d+)")

DELAY_RANGE = (1.5, 3.5)  # seconds between page loads (be polite)

COLUMNS = [
    "Job Title",
    "Company Name",
    "Location",
    "Salary/Hourly Rate",
    "Application Deadline",
    "Job Description",
    "Job URL",
    "Date Scraped",
]


@dataclass
class ScrapeConfig:
    keywords: list[str] = field(default_factory=lambda: list(DEFAULT_KEYWORDS))
    locations: list[str] = field(default_factory=lambda: list(DEFAULT_LOCATIONS))
    max_pages: int = 3          # results pages per keyword
    max_jobs: int = 100         # hard cap per run
    headless: bool = True
    debug: bool = False
    location_filter: bool = True  # keep only jobs whose location matches


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def pause() -> None:
    time.sleep(random.uniform(*DELAY_RANGE))


def log(msg: str) -> None:
    print(f"[scraper] {msg}", flush=True)


def is_login_page(page: Page) -> bool:
    url = page.url.lower()
    return "/login" in url or "sso" in url or "sign_in" in url


def first_text(page: Page, selectors: list[str]) -> str:
    """Return the first non-empty inner text among the candidate selectors."""
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            if loc.count():
                text = loc.inner_text(timeout=2000).strip()
                if text:
                    return text
        except PWTimeout:
            continue
        except Exception:
            continue
    return ""


SALARY_RE = re.compile(
    r"\$\s?\d[\d,]*(?:\.\d+)?\s*(?:k)?"
    r"(?:\s*[-–to]+\s*\$?\s?\d[\d,]*(?:\.\d+)?\s*(?:k)?)?"
    r"(?:\s*(?:/|per\s*)\s*(?:hr|hour|yr|year|month|mo|week|wk))?",
    re.IGNORECASE,
)
DEADLINE_RE = re.compile(
    r"(?:apply\s*by|deadline|applications?\s*close[sd]?(?:\s*on)?)[:\s]+([^\n]{3,40})",
    re.IGNORECASE,
)


def extract_salary(text: str) -> str:
    m = SALARY_RE.search(text)
    return m.group(0).strip() if m else ""


def extract_deadline(text: str) -> str:
    m = DEADLINE_RE.search(text)
    return m.group(1).strip() if m else ""


def location_matches(location: str, description: str, wanted: list[str]) -> bool:
    haystack = f"{location}".lower()
    # Fall back to description if the location field came back blank.
    if not haystack.strip():
        haystack = description[:1500].lower()
    return any(w.lower() in haystack for w in wanted)


# ----------------------------------------------------------------------------
# Login
# ----------------------------------------------------------------------------
def login_and_save_session() -> None:
    """Open a visible browser, let the user log in (SSO/2FA), save auth.json."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        page.goto(LOGIN_URL)
        print(
            "\n=== MANUAL LOGIN ===\n"
            "1. Log in in the browser window (school SSO, 2FA, etc.).\n"
            "2. Wait until you see your Handshake home/feed page.\n"
            "3. Come back here and press Enter to save the session.\n"
        )
        input("Press Enter once you are fully logged in... ")
        context.storage_state(path=str(AUTH_FILE))
        browser.close()
    log(f"Session saved to {AUTH_FILE.resolve()} (keep it private).")


# ----------------------------------------------------------------------------
# Scraping
# ----------------------------------------------------------------------------
def build_search_url(keyword: str, page_num: int) -> str:
    params = {QUERY_PARAM: keyword, PAGE_PARAM: page_num}
    return f"{BASE_URL}{SEARCH_PATH}?{urlencode(params)}"


def collect_job_urls(page: Page, keyword: str, max_pages: int) -> list[str]:
    """Walk result pages for one keyword and return unique job detail URLs."""
    urls: dict[str, str] = {}
    for page_num in range(1, max_pages + 1):
        url = build_search_url(keyword, page_num)
        log(f"Search '{keyword}' page {page_num}")
        page.goto(url, wait_until="domcontentloaded")
        pause()
        if is_login_page(page):
            raise RuntimeError(
                "Session expired or invalid. Run: python scraper.py login"
            )
        try:
            page.wait_for_selector(JOB_LINK_SELECTOR, timeout=10000)
        except PWTimeout:
            log("  no job links found on this page; stopping this keyword.")
            break

        hrefs = page.eval_on_selector_all(
            JOB_LINK_SELECTOR, "els => els.map(e => e.getAttribute('href'))"
        )
        new = 0
        for href in hrefs:
            if not href:
                continue
            m = JOB_ID_RE.search(href)
            if not m:
                continue
            job_id = m.group(1)
            if job_id not in urls:
                urls[job_id] = f"{BASE_URL}/jobs/{job_id}"
                new += 1
        log(f"  +{new} new listings ({len(urls)} total for this keyword)")
        if new == 0:
            break
    return list(urls.values())


def scrape_job(page: Page, url: str) -> dict | None:
    """Open one job detail page and extract the target fields."""
    page.goto(url, wait_until="domcontentloaded")
    pause()
    if is_login_page(page):
        raise RuntimeError("Session expired. Run: python scraper.py login")
    try:
        page.wait_for_selector("h1", timeout=10000)
    except PWTimeout:
        log(f"  skipped (page did not load): {url}")
        return None

    title = first_text(page, SELECTORS["title"])
    company = first_text(page, SELECTORS["company"])
    location = first_text(page, SELECTORS["location"])
    description = first_text(page, SELECTORS["description"])
    body_text = page.locator("body").inner_text(timeout=5000)

    return {
        "Job Title": title,
        "Company Name": company,
        "Location": location,
        "Salary/Hourly Rate": extract_salary(body_text),
        "Application Deadline": extract_deadline(body_text),
        "Job Description": description[:20000],
        "Job URL": url,
        "Date Scraped": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
    }


def save_to_excel(new_rows: list[dict], path: Path = OUTPUT_FILE) -> pd.DataFrame:
    """Append rows to the workbook, de-duplicating on Job URL."""
    new_df = pd.DataFrame(new_rows, columns=COLUMNS)
    if path.exists():
        old_df = pd.read_excel(path)
        df = pd.concat([old_df, new_df], ignore_index=True)
    else:
        df = new_df
    df = df.drop_duplicates(subset="Job URL", keep="first").reset_index(drop=True)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Jobs")
        ws = writer.sheets["Jobs"]
        widths = {"A": 40, "B": 28, "C": 26, "D": 20, "E": 22, "F": 80, "G": 45, "H": 18}
        for col, w in widths.items():
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A2"
    return df


def run_scrape(cfg: ScrapeConfig) -> None:
    if not AUTH_FILE.exists():
        sys.exit("auth.json not found. Run first: python scraper.py login")

    rows: list[dict] = []
    seen_urls: set[str] = set()
    if OUTPUT_FILE.exists():
        seen_urls = set(pd.read_excel(OUTPUT_FILE)["Job URL"].dropna())

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=cfg.headless, slow_mo=50 if cfg.debug else 0)
        context = browser.new_context(storage_state=str(AUTH_FILE))
        page = context.new_page()

        try:
            # 1) Gather candidate URLs across all keywords.
            candidates: list[str] = []
            for kw in cfg.keywords:
                for u in collect_job_urls(page, kw, cfg.max_pages):
                    if u not in candidates and u not in seen_urls:
                        candidates.append(u)
            candidates = candidates[: cfg.max_jobs]
            log(f"{len(candidates)} new listings to open.")

            # 2) Visit each listing and extract fields.
            for i, url in enumerate(candidates, 1):
                log(f"[{i}/{len(candidates)}] {url}")
                try:
                    row = scrape_job(page, url)
                except RuntimeError:
                    raise
                except Exception as exc:  # keep going on one bad page
                    log(f"  error: {exc}")
                    continue
                if not row:
                    continue
                if cfg.location_filter and not location_matches(
                    row["Location"], row["Job Description"], cfg.locations
                ):
                    log(f"  filtered out (location: {row['Location'] or 'unknown'})")
                    continue
                rows.append(row)
        except RuntimeError as exc:
            log(str(exc))
        finally:
            # Refresh the session file so cookies stay current.
            try:
                context.storage_state(path=str(AUTH_FILE))
            except Exception:
                pass
            browser.close()

    if rows:
        df = save_to_excel(rows)
        log(f"Saved {len(rows)} new jobs. {OUTPUT_FILE} now has {len(df)} rows.")
    else:
        log("No new matching jobs were saved.")


# ----------------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(description="Handshake job scraper")
    sub = parser.add_subparsers(dest="cmd", required=True)

    sub.add_parser("login", help="Log in manually and save auth.json")

    sc = sub.add_parser("scrape", help="Scrape jobs using saved session")
    sc.add_argument("--keywords", nargs="+", default=DEFAULT_KEYWORDS)
    sc.add_argument("--locations", nargs="+", default=DEFAULT_LOCATIONS)
    sc.add_argument("--max-pages", type=int, default=3)
    sc.add_argument("--max-jobs", type=int, default=100)
    sc.add_argument("--headed", action="store_true", help="Show the browser")
    sc.add_argument("--debug", action="store_true", help="Slow down actions")
    sc.add_argument("--no-location-filter", action="store_true")

    args = parser.parse_args()
    if args.cmd == "login":
        login_and_save_session()
    else:
        cfg = ScrapeConfig(
            keywords=args.keywords,
            locations=args.locations,
            max_pages=args.max_pages,
            max_jobs=args.max_jobs,
            headless=not args.headed,
            debug=args.debug,
            location_filter=not args.no_location_filter,
        )
        run_scrape(cfg)


if __name__ == "__main__":
    main()
