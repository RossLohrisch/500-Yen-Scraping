#!/usr/bin/env python3
import csv
import re
import time
from pathlib import Path
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse, urljoin, urlsplit, urlunsplit

import requests
from bs4 import BeautifulSoup


# === ONE PLACE to control thumbnail size (pixels) ===
IMG_PX = 160  # used in =IMAGE(...,3,IMG_PX,IMG_PX) and for row/column sizing


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


def normalize_detail_url(u: str | None) -> str | None:
    """
    Use as a fallback dedupe key when item_id is missing.
    Strips fragment and sorts query params for stable comparison.
    """
    if not u:
        return None
    s = urlsplit(u)
    # Sort query params for stability
    q = parse_qs(s.query)
    q_str = urlencode(q, doseq=True)
    return urlunsplit((s.scheme, s.netloc, s.path, q_str, ""))


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


def pixels_to_points(px: int | float) -> float:
    # Excel row height is in points; 1 px ≈ 0.75 pt at 96dpi
    return float(px) * 0.75


def pixels_to_col_width(px: int | float) -> float:
    # Approximate conversion: ~7 px per "character" width; add a little padding.
    return max((float(px) / 7.0) + 2.0, 18.0)


def main(start_url: str, pages: int | None, out_csv: str, delay_sec: float = 1.0,
         until_empty: bool = False, out_xlsx: str | None = None):
    out_path = Path(out_csv)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; listings-bot/1.0; +https://example.org/bot)",
        "Accept-Language": "en-US,en;q=0.9",
    }
    session = requests.Session()
    session.headers.update(headers)

    # Keep title as column B; include an image preview and translation formula.
    fieldnames = [
        "item_id",      # A
        "title",        # B
        "price_jpy",    # C
        "buyout_jpy",   # D
        "detail_url",   # E
        "image_url",    # F
        "image_preview",# G (new) =IMAGE(Fn)
        "source_page",  # H
        "title_en",     # I =TRANSLATE(Bn,"ja","en")
    ]

    # Derive starting page from the start_url
    try:
        qs = parse_qs(urlparse(start_url).query)
        start_page = int(qs.get("page", ["1"])[0])
    except Exception:
        start_page = 1

    # Optional XLSX workbook that pre-sizes rows to fit images
    workbook = worksheet = None
    if out_xlsx:
        try:
            import xlsxwriter
        except ImportError:
            raise SystemExit("Please install xlsxwriter: pip install xlsxwriter")

        workbook = xlsxwriter.Workbook(out_xlsx)
        worksheet = workbook.add_worksheet("Listings")
        # Write headers
        for c, h in enumerate(fieldnames):
            worksheet.write(0, c, h)
        # Nice column widths
        worksheet.set_column(0, 0, 12)   # item_id
        worksheet.set_column(1, 1, 60)   # title
        worksheet.set_column(2, 3, 12)   # prices
        worksheet.set_column(4, 5, 36)   # detail_url, image_url
        worksheet.set_column(6, 6, pixels_to_col_width(IMG_PX))  # image_preview column sized to image width
        worksheet.set_column(7, 7, 30)   # source_page
        worksheet.set_column(8, 8, 36)   # title_en

    # === DEDUPING STATE ===
    seen_item_ids: set[str] = set()
    seen_detail_urls: set[str] = set()

    # Write CSV (formulas included)
    with out_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()

        page = start_page
        total_written = 0
        excel_row = 2   # first data row (row 1 is header)
        xlsx_row0 = 1   # first data row in xlsxwriter (0-based)

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

            page_written = 0
            page_skipped = 0

            for row in items:
                # Build dedupe keys
                iid = (row.get("item_id") or "").strip() or None
                norm_url = normalize_detail_url(row.get("detail_url"))

                # Decide if we've seen this already
                duplicate = False
                if iid and iid in seen_item_ids:
                    duplicate = True
                elif not iid and norm_url and norm_url in seen_detail_urls:
                    duplicate = True

                if duplicate:
                    page_skipped += 1
                    continue

                # Mark as seen
                if iid:
                    seen_item_ids.add(iid)
                elif norm_url:
                    seen_detail_urls.add(norm_url)

                # ---- Your exact image formula size (160x160) and translate formula ----
                row["image_preview"] = f'=IF(LEN(F{excel_row}),IMAGE(F{excel_row},"",3,{IMG_PX},{IMG_PX}),"")'
                row["title_en"] = f'=TRANSLATE(B{excel_row},"ja","en")'

                # CSV
                writer.writerow(row)

                # XLSX (pre-sized rows)
                if worksheet is not None:
                    for c, key in enumerate(fieldnames):
                        if key == "image_preview":
                            worksheet.write_formula(
                                xlsx_row0, c,
                                f'=IF(LEN(F{xlsx_row0+1}),IMAGE(F{xlsx_row0+1},"",3,{IMG_PX},{IMG_PX}),"")'
                            )
                        elif key == "title_en":
                            worksheet.write_formula(
                                xlsx_row0, c,
                                f'=TRANSLATE(B{xlsx_row0+1},"ja","en")'
                            )
                        else:
                            worksheet.write(xlsx_row0, c, row.get(key))

                    # Set row height to fit IMG_PX
                    worksheet.set_row(xlsx_row0, pixels_to_points(IMG_PX))

                excel_row += 1
                xlsx_row0 += 1
                total_written += 1
                page_written += 1

            print(f"[+] Page {page}: wrote {page_written} new items, skipped {page_skipped} duplicates (total {total_written})")

            # Stop conditions
            if pages is not None and page >= start_page + pages - 1:
                break
            page += 1
            time.sleep(delay_sec)

    if workbook is not None:
        workbook.close()

    print(f"Done. Wrote {total_written} unique rows to {out_csv}" + (f" and {out_xlsx}" if out_xlsx else ""))


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Scrape 500yenshop Yahoo Auctions listings.")
    ap.add_argument("--start-url", required=True, help="The URL to page=1 (or any page).")
    ap.add_argument("--pages", type=int, default=None, help="Number of pages to fetch.")
    ap.add_argument("--until-empty", action="store_true", help="Stop when a page has no items.")
    ap.add_argument("--out", default="listings.csv", help="Output CSV filename.")
    ap.add_argument("--out-xlsx", default=None, help="Also write an Excel .xlsx with pre-sized rows.")
    ap.add_argument("--delay", type=float, default=1.0, help="Seconds to sleep between pages.")
    args = ap.parse_args()

    main(
        start_url=args.start_url,
        pages=args.pages,
        out_csv=args.out,
        delay_sec=args.delay,
        until_empty=args.until_empty,
        out_xlsx=args.out_xlsx,
    )
