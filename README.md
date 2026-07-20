# Web Scraper → DOCX

Paste any website URL and export page content to a local Word document.

- **Any URL** → title, description, headings, main text, links, images
- **Shopify Partner Directory** URLs (e.g. `shopify.com/partners/directory/partner/...`) → full specialized export including all review pages

## Setup (once)

```bash
cd "D:/Work/Web Scrapping Bot(profile content)"
python -m pip install -r requirements.txt
python -m playwright install chromium
```

## Run

Interactive (asks for URL):

```bash
python scrape.py
```

Or pass the URL directly:

```bash
python scrape.py --url "https://example.com"
python scrape.py -u "https://www.shopify.com/partners/directory/partner/forsuccess"
```

### Options

```bash
# Show browser while scraping
python scrape.py -u "https://example.com" --headed

# Force generic mode even on a Shopify partner URL
python scrape.py -u "https://www.shopify.com/partners/directory/partner/forsuccess" --mode generic

# Custom output path
python scrape.py -u "https://example.com" --out "output/my_page.docx"

# More scrolling for long lazy-loaded pages
python scrape.py -u "https://example.com" --scroll 15
```

Files are saved under `output/` as `.docx` + `.json`.

## Notes

- Respect site terms of use; only scrape pages you’re allowed to copy.
- Shopify partner full review export can take a few minutes (many pages).
- The older entry point `scrape_partner_profile.py` still works for partner profiles.
