#!/usr/bin/env python3
import csv
import re
import time
from pathlib import Path
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse, urljoin

import requests
from bs4 import BeautifulSoup


def set_page_param(url: str, page: int) -> str:
    parts = list(urlparse(url))
    qs = parse_qs(parts[4])
    qs["page"] = [str(page)]
    parts[4] = urlencode(qs, doseq=True)
    return urlunparse(parts)


def extract_item_id(href: str) -> str | None:
    if not href:
        return None
    try:
        q = parse_qs(urlparse(href).query)
        return q.get("id", [None])[0]
    except Exception:
        return None


def clean_price(text: str) -> int | None:
    if not text:
        return None
    digits = re.sub(r"[^\d]", "", text)
    return int(digits) if digits.isdigit() else None


def scrape_page(session: requests.Session, page_url: str) -> list[dict]:
    resp = session.get(page_url, timeout=20)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    items = []
    for block in soup.select("div.list_item_block"):
        # Title
        title_tag = block.select_one(".products-txt a.translate h4")
        title = title_tag.get_text(strip=True) if title_tag else None

        # Auction price (may be shown as current_price OR current_listing_price)
        price_tag = block.select_one(
            ".short-price .current_price strong, .short-price .current_listing_price strong"
        )
        price_text = price_tag.get_text(strip=True) if price_tag else None
        price_jpy = clean_price(price_text)

        # Buyout price
        buyout_tag = block.select_one(".short-price .buy_now_price strong")
        buyout_text = buyout_tag.get_text(strip=True) if buyout_tag else None
        buyout_jpy = clean_price(buyout_text)

        # Detail URL
        a = block.select_one(".products-txt a.translate")
        rel_href = a.get("href") if a else None
        detail_url = urljoin(page_url, rel_href) if rel_href else None

        # Image URL
        img = block.select_one(".products-pic img")
        img_src = img.get("src").strip() if img and img.get("src") else None
        image_url = urljoin(page_url, img_src) if img_src else None

        # Item ID
        item_id = extract_item_id(rel_href)

        if title or price_jpy or buyout_jpy or detail_url:
            items.append(
                {
                    "item_id": item_id,
                    "title": title,
                    "price_jpy": price_jpy,      # current auction price
                    "buyout_jpy": buyout_jpy,    # buyout price (optional)
                    "detail_url": detail_url,
                    "image_url": image_url,
                    "source_page": page_url,
                }
            )

    return items


def main(start_url: str, pages: int | None, out_csv: str, delay_sec: float = 1.0, until_empty: bool = False):
    out_path = Path(out_csv)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; listings-bot/1.0; +https://example.org/bot)",
        "Accept-Language": "en-US,en;q=0.9",
    }
    session = requests.Session()
    session.headers.update(headers)

    # Keep title as column B; add a new column for the translation formula.
    fieldnames = [
        "item_id",
        "title",        # <-- column B
        "price_jpy",
        "buyout_jpy",
        "detail_url",
        "image_url",
        "source_page",
        "title_en",     # formula column Excel/Sheets will evaluate
    ]

    # Derive starting page from the start_url
    try:
        qs = parse_qs(urlparse(start_url).query)
        start_page = int(qs.get("page", ["1"])[0])
    except Exception:
        start_page = 1

    with out_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()

        page = start_page
        total = 0
        excel_row = 2  # first data row (row 1 is header)

        while True:
            page_url = set_page_param(start_url, page)
            print(f"Fetching page {page}: {page_url}")

            try:
                items = scrape_page(session, page_url)
            except requests.HTTPError as e:
                print(f"[!] HTTP error on page {page}: {e}")
                break
            except Exception as e:
                print(f"[!] Unexpected error on page {page}: {e}")
                break

            if not items:
                print(f"[-] Page {page} returned zero items.")
                if until_empty:
                    break

            for row in items:
                # Excel (Microsoft 365) example:
                row["title_en"] = f'=TRANSLATE(B{excel_row},"ja","en")'
                # If opening in Google Sheets instead, use:
                # row["title_en"] = f'=GOOGLETRANSLATE(B{excel_row},"ja","en")'

                writer.writerow(row)
                excel_row += 1

            total += len(items)
            print(f"[+] Page {page}: wrote {len(items)} items (total {total})")

            # Stop conditions
            if pages is not None and page >= start_page + pages - 1:
                break
            page += 1
            time.sleep(delay_sec)

    print(f"Done. Wrote {total} rows to {out_csv}")


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Scrape 500yenshop Yahoo Auctions listings.")
    ap.add_argument("--start-url", required=True, help="The URL to page=1 (or any page).")
    ap.add_argument("--pages", type=int, default=None, help="Number of pages to fetch.")
    ap.add_argument("--until-empty", action="store_true", help="Stop when a page has no items.")
    ap.add_argument("--out", default="listings.csv", help="Output CSV filename.")
    ap.add_argument("--delay", type=float, default=1.0, help="Seconds to sleep between pages.")
    args = ap.parse_args()

    main(
        start_url=args.start_url,
        pages=args.pages,
        out_csv=args.out,
        delay_sec=args.delay,
        until_empty=args.until_empty,
    )
