"""
Universal web scraper → DOCX

Paste any URL. Shopify Partner Directory profiles use the specialized scraper;
all other sites use the generic page scraper.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from urllib.parse import urlparse

import scrape_generic
import scrape_partner_profile as shopify

OUTPUT_DIR = Path(__file__).resolve().parent / "output"

SHOPIFY_PARTNER_RE = re.compile(
    r"shopify\.com/partners/directory/partner/[^/?#]+",
    re.I,
)


def normalize_url(url: str) -> str:
    url = (url or "").strip().strip('"').strip("'")
    if not url:
        return ""
    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url
    return url


def is_shopify_partner(url: str) -> bool:
    return bool(SHOPIFY_PARTNER_RE.search(url))


def prompt_url() -> str:
    try:
        raw = input("Enter URL to scrape: ").strip()
    except EOFError:
        raw = ""
    return normalize_url(raw)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Scrape any website URL to a local DOCX (Shopify Partner profiles get full review export)."
    )
    parser.add_argument(
        "--url",
        "-u",
        default="",
        help="Page URL to scrape. If omitted, you will be prompted.",
    )
    parser.add_argument(
        "--mode",
        choices=("auto", "generic", "shopify"),
        default="auto",
        help="auto = detect Shopify partner URLs; generic = always general page scrape; shopify = partner profile mode",
    )
    parser.add_argument("--headed", action="store_true", help="Show the browser window")
    parser.add_argument("--out", default="", help="Output DOCX path")
    parser.add_argument(
        "--scroll",
        type=int,
        default=8,
        help="Scroll passes for generic pages (default: 8)",
    )
    args = parser.parse_args()

    url = normalize_url(args.url) or prompt_url()
    if not url:
        print("Error: URL is required.", file=sys.stderr)
        sys.exit(1)

    parsed = urlparse(url)
    if not parsed.netloc:
        print(f"Error: invalid URL: {url}", file=sys.stderr)
        sys.exit(1)

    use_shopify = args.mode == "shopify" or (
        args.mode == "auto" and is_shopify_partner(url)
    )

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    headless = not args.headed

    if use_shopify:
        print("Mode: Shopify Partner profile (full reviews)")
        data = shopify.scrape(url, headless=headless)
        slug = url.rstrip("/").split("/")[-1].split("?")[0] or "partner"
        json_path = OUTPUT_DIR / f"{slug}_partner_profile.json"
        docx_path = Path(args.out) if args.out else OUTPUT_DIR / f"{slug}_partner_profile.docx"
        json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        shopify.build_docx(data, docx_path)
        print("\nDone.")
        print(f"  Reviews scraped: {data.get('reviews_scraped', 0)} (listed: {data.get('review_count')})")
        print(f"  JSON: {json_path}")
        print(f"  DOCX: {docx_path}")
        return

    print("Mode: generic website")
    data = scrape_generic.scrape(url, headless=headless, scroll_passes=args.scroll)
    slug = scrape_generic.slug_from_url(data.get("url") or url)
    json_path = OUTPUT_DIR / f"{slug}.json"
    docx_path = Path(args.out) if args.out else OUTPUT_DIR / f"{slug}.docx"
    json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    scrape_generic.build_docx(data, docx_path)

    print("\nDone.")
    print(f"  Title: {data.get('title', '')}")
    print(f"  Text length: {len(data.get('main_text') or '')} chars")
    print(f"  Links: {len(data.get('links') or [])} | Images: {len(data.get('images') or [])}")
    print(f"  JSON: {json_path}")
    print(f"  DOCX: {docx_path}")


if __name__ == "__main__":
    main()
