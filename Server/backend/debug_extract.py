"""
ShopIntel AI — Local Debug Tool
================================
This script runs LOCALLY on your machine and shows exactly:
  - What HTML the site returns
  - Which CSS selectors hit / miss
  - What JSON-LD contains
  - What prices are found in the raw HTML

Run: python3 debug_extract.py

Then COPY the output and share it — we'll fix the selectors based on
what your machine actually gets from each site.
"""

import re
import json
import requests
from bs4 import BeautifulSoup

UA_DESKTOP = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/137.0.0.0 Safari/537.36"
)
UA_MOBILE = (
    "Mozilla/5.0 (Linux; Android 13; Pixel 7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/137.0.0.0 Mobile Safari/537.36"
)

def make_headers(ua, referer="https://www.google.com/"):
    return {
        "User-Agent": ua,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "en-IN,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "cross-site",
        "Sec-Fetch-User": "?1",
        "Referer": referer,
        "Cache-Control": "max-age=0",
        "sec-ch-ua": '"Chromium";v="137","Google Chrome";v="137","Not=A?Brand";v="24"',
        "sec-ch-ua-mobile": "?1" if "Mobile" in ua else "?0",
        "sec-ch-ua-platform": '"Android"' if "Mobile" in ua else '"Windows"',
    }

def fetch(url, mobile=False):
    ua = UA_MOBILE if mobile else UA_DESKTOP
    session = requests.Session()
    try:
        r = session.get(url, headers=make_headers(ua), timeout=20, allow_redirects=True)
        return r.text, r.status_code
    except Exception as e:
        return "", str(e)

def debug_url(url, label="", mobile=False):
    print("\n" + "═"*70)
    print(f"  {label or url}")
    print("═"*70)

    html, status = fetch(url, mobile)
    print(f"  HTTP Status : {status}")
    print(f"  HTML Size   : {len(html):,} chars")

    if not html or len(html) < 200:
        print(f"  ❌ Empty/blocked response: {str(html)[:200]}")
        return

    soup = BeautifulSoup(html, "html.parser")

    # Page title
    title = soup.title.text.strip() if soup.title else "(no title)"
    print(f"  Page Title  : {title[:80]}")

    # ── JSON-LD ───────────────────────────────────────────────────────────
    print("\n  ── JSON-LD blocks ──")
    scripts = soup.find_all("script", type="application/ld+json")
    print(f"  Found {len(scripts)} JSON-LD scripts")
    for i, s in enumerate(scripts):
        try:
            d = json.loads(s.string or "")
            dtype = d.get("@type", "?") if isinstance(d, dict) else type(d).__name__
            # Print product-relevant info
            if isinstance(d, dict):
                name  = d.get("name", "")[:60]
                offers = d.get("offers", {})
                price = ""
                if isinstance(offers, dict):
                    price = offers.get("price", "")
                elif isinstance(offers, list) and offers:
                    price = offers[0].get("price", "")
                imgs  = d.get("image", "")
                img0  = (imgs[0] if isinstance(imgs, list) else imgs)[:80] if imgs else ""
                agg   = d.get("aggregateRating", {})
                rating = agg.get("ratingValue", "") if isinstance(agg, dict) else ""
                reviews = agg.get("reviewCount", "") if isinstance(agg, dict) else ""
                print(f"  [{i}] @type={dtype}")
                print(f"       name   = {name}")
                print(f"       price  = {price}")
                print(f"       image  = {img0}")
                print(f"       rating = {rating}  reviews = {reviews}")
            else:
                print(f"  [{i}] @type={dtype} (array/other)")
        except Exception as e:
            print(f"  [{i}] parse error: {e}")

    # ── OG tags ───────────────────────────────────────────────────────────
    print("\n  ── Open Graph / Meta tags ──")
    for prop in ["og:title", "og:image", "og:description",
                 "og:price:amount", "product:price:amount"]:
        tag = soup.find("meta", property=prop)
        val = tag.get("content", "")[:80] if tag else "(not found)"
        print(f"  {prop:<28} : {val}")

    # ── Raw ₹ prices in HTML ──────────────────────────────────────────────
    print("\n  ── Raw ₹ prices found in HTML ──")
    prices = re.findall(r"₹\s?[\d,]+", html[:50000])
    print(f"  {prices[:10] or '(none found)'}")

    # ── Key JSON price patterns ───────────────────────────────────────────
    print("\n  ── JSON price patterns in HTML ──")
    for pat in [
        r'"finalPrice"\s*:\s*(\d+)',
        r'"sellingPrice"\s*:\s*(\d+)',
        r'"discountedPrice"\s*:\s*(\d+)',
        r'"mrpPrice"\s*:\s*(\d+)',
        r'"price"\s*:\s*"?(\d{3,6})"?',
        r'"listingPrice"\s*:\s*(\d+)',
        r'data-price="(\d+)"',
        r'selling_price["\']?\s*:\s*["\']?(\d+)',
    ]:
        m = re.search(pat, html)
        result = m.group(1) if m else "(not found)"
        print(f"  {pat[:45]:<45} : {result}")

    # ── Important elements found ──────────────────────────────────────────
    print("\n  ── Key page elements ──")
    checks = [
        ("#productTitle",              "Amazon title"),
        (".priceToPay .a-offscreen",   "Amazon priceToPay"),
        (".a-price .a-offscreen",      "Amazon a-price"),
        ("span.a-price-whole",         "Amazon price-whole"),
        ("#landingImage",              "Amazon main image"),
        ("span.B_NuCI",               "Flipkart title (old)"),
        ("h1.yhB1nd",                 "Flipkart title (new)"),
        ("div.Nx9bqj",                "Flipkart price div"),
        ("div._30jeq3",               "Flipkart price (old)"),
        ("img._396cs4",               "Flipkart image"),
        ("h1.product__title",         "TSS title"),
        ("span.price-item--sale",     "TSS price"),
        (".product__media img",       "TSS image"),
        ("h1.pdp-name",              "Myntra title"),
        ("span.pdp-price strong",    "Myntra price"),
        ("h1.prod-name",             "Ajio title"),
        ("div.prod-sp",              "Ajio price"),
    ]
    for sel, label in checks:
        el = soup.select_one(sel)
        if el:
            txt = el.get_text(strip=True)[:50]
            src = el.get("src", "")[:50]
            print(f"  ✅ {label:<30} → {txt or src or '(found, empty)'}")
        else:
            print(f"  ❌ {label:<30} → not found")

    # ── Bot detection ──────────────────────────────────────────────────────
    print("\n  ── Bot detection signals ──")
    bot_signals = {
        "CAPTCHA": "captcha" in html.lower(),
        "Robot Check": "robot check" in html.lower(),
        "Access Denied": "access denied" in html.lower(),
        "JS-only (#app div)": bool(soup.find(id="app") or soup.find(id="root")),
        "Very short page (<5K)": len(html) < 5000,
    }
    for signal, triggered in bot_signals.items():
        print(f"  {'⚠️' if triggered else '  '} {signal}: {triggered}")

    print()


# ── PASTE YOUR ACTUAL PRODUCT URLs HERE ──────────────────────────────────────
URLS_TO_TEST = [
    # Replace these with your exact URLs from the browser
    ("Amazon",          "https://www.amazon.in/Puma-Running-Shoe-Black-White-31078901/dp/B0DLGT243T?adgrpid=1316117027997683&dib=eyJ2IjoiMSJ9.bhPlOCdfIcmgeDpatmktI6xxmVpc6go0fEe1p3l3538TJvg4IZPU7sVuzg55MQBqLe_vS0eTE-Jm6ORw_jqZA7iEpbdBuqoUvJJHsQwm0PkVGZit82cgij_oADpcFnsTsBPUqZJa8Lzgxl0ln4mi9dTEkwnsP_lal4pnehxvEYh2hDpfl2vigR92odHgdETVx3aNUrjPqpx3mS-zU3Fro_DUi_PPC9BRapLemJ6wFTmISpdyOY9AjOmON4kcfDtMnX1TEuYOzAe_VMbYDH5d8b0z5VAvKp0kjLhWGaytmn0.iEQXLnsyhjyisjeV68zvIAEvkXDC6tcD1_uIt2GebpI&dib_tag=se&hvadid=82257584952915&hvbmt=bb&hvdev=c&hvlocphy=264924&hvnetw=o&hvqmt=b&hvtargid=kwd-82258194075177%3Aloc-90&hydadcr=22286_2343049&keywords=puma%2Bnitro%2Bshoes&mcid=0269c1100330326a9b6e5b032bd27d33&qid=1791298317&sr=8-7&th=1&psc=1",           False),
    ("Flipkart",        "https://www.flipkart.com/puma-rungryp-sneakers-men/p/itm014500194b6a8?pid=SHOGTDWQB73CYAFF&lid=LSTSHOGTDWQB73CYAFFJDHWEN&marketplace=FLIPKART&store=osp%2Fcil&srno=b_1_10&otracker=browse&fm=organic&iid=9aa951e9-0e5b-4c64-9c69-eac4bb204724.SHOGTDWQB73CYAFF.SEARCH&ppt=browse&ppn=browse&ssid=1b9e06cmuo0000001791298377637&ov_redirect=true",       True),   # mobile=True
    ("Souled Store",    "https://www.thesouledstore.com/product/top-ace-mocha-men-low-top-sneakers?get=1",   False),
    ("Puma",    "https://in.puma.com/in/en/pd/smashic-comfort-casual-sneakers/394371?swatch=07",   False),
]

if __name__ == "__main__":
    for label, url, mobile in URLS_TO_TEST:
        debug_url(url, label=label, mobile=mobile)

    print("\n" + "═"*70)
    print("  Done. Copy ALL output above and share it.")
    print("  This tells us exactly what your machine sees from each store.")
    print("═"*70)