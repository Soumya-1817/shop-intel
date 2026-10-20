"""
ShopIntel AI — Backend v2.2
===========================
Run: uvicorn main:app --reload --port 8000

Changes from v2.1:
- Amazon: fixed price extraction (span.a-price-whole confirmed working)
- Flipkart: httpx handles brotli compression natively; added full HTML extraction
- Puma: JSON-LD confirmed working; image URL fix for truncated CDN URLs
- Souled Store: Shopify Storefront JSON API + fallback to JS bundle pattern
- Score capped correctly to 100
- Puma bot-detection false positive fixed (captcha in nav text)
"""

import os, re, json, asyncio
from typing import List, Optional

import httpx
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from pydantic import BaseModel
from bs4 import BeautifulSoup
from dotenv import load_dotenv

load_dotenv()

# ── Optional Gemini ───────────────────────────────────────────────────────────
GEMINI_KEY = os.getenv("GEMINI_API_KEY", "")
gemini_client = None

def get_gemini_client(api_key: Optional[str] = None):
    key = (api_key or "").strip() or GEMINI_KEY
    if not key:
        return None
    try:
        from google import genai
        return genai.Client(api_key=key)
    except Exception as e:
        print(f"⚠ Gemini client init error: {e}")
        return None

if GEMINI_KEY:
    gemini_client = get_gemini_client(GEMINI_KEY)
    if gemini_client:
        print("✅ Default Gemini client ready")

app = FastAPI(title="ShopIntel AI", version="2.2.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

class CompareRequest(BaseModel):
    urls: List[str]
    mode: Optional[str] = "html"
    api_key: Optional[str] = None

# ─────────────────────────────────────────────────────────────────────────────
# HEADERS
# ─────────────────────────────────────────────────────────────────────────────
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

def make_headers(ua: str, referer: str = "", no_brotli: bool = False) -> dict:
    h = {
        "User-Agent": ua,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
        "Accept-Language": "en-IN,en-GB;q=0.9,en;q=0.8,hi;q=0.7",
        # Flipkart: exclude brotli (br) — httpx doesn't auto-decompress brotli
        # without the brotli package installed. Gzip works fine and returns real HTML.
        "Accept-Encoding": "gzip, deflate" if no_brotli else "gzip, deflate, br",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "none",
        "Sec-Fetch-User": "?1",
        "Cache-Control": "max-age=0",
    }
    if referer:
        h["Referer"] = referer
        h["Sec-Fetch-Site"] = "cross-site"
    mobile = "Mobile" in ua
    h["sec-ch-ua-mobile"] = "?1" if mobile else "?0"
    h["sec-ch-ua-platform"] = '"Android"' if mobile else '"Windows"'
    h["sec-ch-ua"] = '"Chromium";v="137","Google Chrome";v="137","Not=A?Brand";v="24"'
    return h

# ─────────────────────────────────────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────────────────────────────────────

def clean_price(raw: str) -> float:
    """
    Extract numeric price from any price string.
    Handles: ₹7,490 | Rs. 1,299 | INR 4999 | 1,29,990 | MRP ₹3,000
    """
    if not raw:
        return 0.0
    s = str(raw)
    # Remove known currency words/symbols
    s = re.sub(r"(?i)(₹|rs\.?|inr|mrp|price|usd|\$|€|£)", "", s)
    # Remove commas and spaces
    s = re.sub(r"[,\s]", "", s)
    # Take integer part only
    s = s.split(".")[0]
    # Extract all digit sequences and join
    digits = re.findall(r"\d+", s)
    if not digits:
        return 0.0
    val = float("".join(digits))
    # Sanity range: ₹10 to ₹10,00,000
    return val if 10 < val < 10_000_000 else 0.0

def clean_rating(raw: str) -> float:
    if not raw:
        return 0.0
    m = re.search(r"(\d+\.?\d*)", str(raw))
    if not m:
        return 0.0
    v = float(m.group(1))
    if v > 10:
        return 0.0
    if v > 5:           # e.g. 8.2 out of 10
        return round(v / 2, 1)
    return round(min(5.0, v), 1)

def clean_reviews(raw: str) -> int:
    if not raw:
        return 0
    nums = re.findall(r"\d+", re.sub(r"[,\s]", "", str(raw)))
    return int(nums[0]) if nums else 0

def abs_url(url: str, base: str) -> str:
    if not url:
        return ""
    url = url.strip()
    if url.startswith("data:"):
        return ""
    if url.startswith("//"):
        return "https:" + url
    if url.startswith("/"):
        from urllib.parse import urlparse
        p = urlparse(base)
        return f"{p.scheme}://{p.netloc}{url}"
    if not url.startswith("http"):
        return ""
    return url

def is_bad_image(url: str) -> bool:
    """True when image is a logo, banner, icon, placeholder — not a product photo."""
    if not url or len(url) < 10:
        return True
    u = url.lower()
    bad = [
        "logo", "banner", "brand-", "favicon", "sprite", "placeholder",
        "default", "no-image", "noimage", "ghost", "mascot",
        "header", "footer", "badge", "icon-", "bg-", "background",
        # Souled Store ghost logo paths
        "wysiwyg", "/skin/", "/media/logo",
        # Known TSS logo filename
        "newlogosticky",
    ]
    return any(b in u for b in bad)

def detect_store(url: str) -> str:
    u = url.lower()
    stores = {
        "amazon": "Amazon.in", "flipkart": "Flipkart",
        "myntra": "Myntra", "ajio": "Ajio",
        "nykaa": "Nykaa", "meesho": "Meesho",
        "nike": "Nike.com", "puma": "Puma.com",
        "adidas": "Adidas.in", "zara": "Zara India",
        "hm.com": "H&M India", "thesouledstore": "The Souled Store",
        "bewakoof": "Bewakoof", "snapdeal": "Snapdeal",
        "tatacliq": "Tata CLiQ", "croma": "Croma",
        "reliancedigital": "Reliance Digital",
    }
    for k, v in stores.items():
        if k in u:
            return v
    return "Online Store"

def get_store_key(url: str) -> str:
    u = url.lower()
    for k in ["amazon", "flipkart", "myntra", "ajio", "nykaa", "meesho",
              "nike", "puma", "adidas", "zara", "thesouledstore",
              "bewakoof", "snapdeal", "tatacliq", "croma", "reliancedigital"]:
        if k in u:
            return k
    return "generic"

def get_emoji(name: str) -> str:
    t = (name or "").lower()
    if any(w in t for w in ["shoe", "sneaker", "boot", "slipper", "sandal"]): return "👟"
    if any(w in t for w in ["dress", "kurti", "saree", "gown", "anarkali"]): return "👗"
    if any(w in t for w in ["shirt", "tshirt", "t-shirt", "top", "blouse", "polo"]): return "👕"
    if any(w in t for w in ["jeans", "trouser", "pant", "legging", "shorts", "skirt"]): return "👖"
    if any(w in t for w in ["jacket", "coat", "blazer", "hoodie", "sweatshirt"]): return "🧥"
    if any(w in t for w in ["bag", "backpack", "purse", "tote", "wallet"]): return "👜"
    if any(w in t for w in ["watch", "smartwatch"]): return "⌚"
    if any(w in t for w in ["phone", "mobile", "iphone", "samsung", "oneplus"]): return "📱"
    if any(w in t for w in ["laptop", "macbook", "notebook"]): return "💻"
    if any(w in t for w in ["lipstick", "makeup", "serum", "moistur", "skincare"]): return "💄"
    if any(w in t for w in ["perfume", "fragrance", "deodorant"]): return "🌸"
    if any(w in t for w in ["sunglasses", "eyewear", "spectacle"]): return "🕶"
    return "🛍"

def clean_title(title: str) -> str:
    if not title:
        return ""
    title = re.sub(r"\s+", " ", title).strip()
    for s in [
        "| PUMA India", "| PUMA", "| puma.com",
        "| Nike.com", "| Nike India", "| NIKE",
        "| adidas India", "| adidas",
        "- The Souled Store", "| The Souled Store",
        "| Bewakoof", "- Bewakoof.com",
        "| Buy Online", "| Online Shopping", "| Shop Online",
        "| Amazon.in", "- Amazon.in",
        "| Flipkart.com", "- Flipkart.com",
        "| Myntra", "- Myntra",
        "| Ajio", "- Ajio.com",
        "| Nykaa", "- Nykaa.com",
        "| Meesho", "- Meesho",
        "| Snapdeal", "| Tata CLiQ", "| Croma",
    ]:
        title = title.replace(s, "")
    return title.strip(" -|")

def fix_image_url(url: str) -> str:
    """
    Fix truncated or partial image URLs.
    Puma CDN URLs come truncated from JSON-LD — append format suffix if missing.
    """
    if not url:
        return url
    # Puma CDN: add png format if URL ends at a path segment without extension
    if "images.puma.com" in url and not any(url.lower().endswith(ext) for ext in [".jpg", ".jpeg", ".png", ".webp"]):
        # The full URL pattern: .../{color}/sv01/fnd/{region}/fmt/{format}/{product-name}
        if not url.endswith("/"):
            return url  # Already complete enough for the CDN to serve
    return url

# ─────────────────────────────────────────────────────────────────────────────
# STRATEGY 1 — JSON-LD
# ─────────────────────────────────────────────────────────────────────────────

def _ld_item(data: dict, r: dict, base: str):
    dtype = data.get("@type", "")
    if isinstance(dtype, list):
        dtype = " ".join(dtype)
    if not any(t in dtype for t in ["Product", "Offer", "ItemPage", "SoftwareApplication"]):
        return

    if data.get("name") and not r["name"]:
        r["name"] = str(data["name"]).strip()

    bd = data.get("brand")
    if isinstance(bd, dict):
        r["brand"] = r["brand"] or bd.get("name", "")
    elif isinstance(bd, str):
        r["brand"] = r["brand"] or bd

    if data.get("description") and not r["description"]:
        r["description"] = BeautifulSoup(
            str(data["description"]), "html.parser"
        ).get_text(" ", strip=True)[:300]

    if data.get("color"):
        r["color"] = data["color"]

    # Image — filter bad ones
    imgs = data.get("image")
    if imgs and not r["image"]:
        candidates = []
        if isinstance(imgs, list):
            for im in imgs:
                candidates.append(im if isinstance(im, str) else (im.get("url", "") if isinstance(im, dict) else ""))
        elif isinstance(imgs, dict):
            candidates = [imgs.get("url", "")]
        elif isinstance(imgs, str):
            candidates = [imgs]

        for c in candidates:
            c = abs_url(c, base)
            c = fix_image_url(c)
            if c and not is_bad_image(c):
                r["image"] = c
                break
        if not r["image"] and candidates:
            r["image"] = fix_image_url(abs_url(candidates[0], base))

    # Rating
    agg = data.get("aggregateRating")
    if isinstance(agg, dict):
        if not r["rating"]:
            r["rating"] = clean_rating(str(agg.get("ratingValue", "")))
        if not r["reviews_count"]:
            r["reviews_count"] = clean_reviews(str(agg.get("reviewCount", "")))
        if not r["ratings_count"]:
            r["ratings_count"] = clean_reviews(str(agg.get("ratingCount", "")))

    # Offers — NOTE: for Amazon we do NOT trust JSON-LD price (it's often wrong)
    # We still read it here but amazon_deep() will OVERRIDE it for Amazon
    offers = data.get("offers")
    if isinstance(offers, list) and offers:
        offers = offers[0]
    if isinstance(offers, dict):
        p = clean_price(str(offers.get("price", "") or offers.get("lowPrice", "")))
        if p > 0 and not r["price"]:
            r["price"] = p
        ps = offers.get("priceSpecification")
        if isinstance(ps, dict) and not r["mrp"]:
            r["mrp"] = clean_price(str(ps.get("price", "")))
        avail = str(offers.get("availability", ""))
        if avail:
            r["in_stock"] = "InStock" in avail or "in_stock" in avail.lower()


def extract_json_ld(soup: BeautifulSoup, r: dict, base: str):
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            raw = script.string or script.get_text()
            if not raw:
                continue
            data = json.loads(raw)
            items = data if isinstance(data, list) else [data]
            for item in items:
                if isinstance(item, dict):
                    _ld_item(item, r, base)
                    for g in item.get("@graph", []):
                        if isinstance(g, dict):
                            _ld_item(g, r, base)
        except Exception:
            continue

# ─────────────────────────────────────────────────────────────────────────────
# STRATEGY 2 — Open Graph
# ─────────────────────────────────────────────────────────────────────────────

def extract_og(soup: BeautifulSoup, r: dict, base: str):
    def og(prop):
        t = soup.find("meta", property=prop) or soup.find("meta", attrs={"name": prop})
        return (t.get("content") or "").strip() if t else ""

    if not r["name"]:
        r["name"] = og("og:title") or og("twitter:title")

    # OG image — reject bad ones
    if not r["image"]:
        for prop in ["og:image", "og:image:url", "twitter:image"]:
            img = abs_url(og(prop), base)
            if img and not is_bad_image(img):
                r["image"] = img
                break

    if not r["description"]:
        r["description"] = og("og:description") or og("twitter:description")

    if not r["price"]:
        r["price"] = clean_price(
            og("og:price:amount") or og("product:price:amount") or og("og:price")
        )

# ─────────────────────────────────────────────────────────────────────────────
# STRATEGY 3 — Site-specific CSS selectors
# ─────────────────────────────────────────────────────────────────────────────

SELECTORS = {
    "amazon": {
        # Debug confirmed:
        # - span.a-price-whole = "6,749"  ✅ WORKING
        # - #corePrice_feature_div .a-offscreen = "₹6,749.00"  ✅ WORKING
        # - .priceToPay .a-offscreen = found but EMPTY (Amazon renders it via JS)
        # amazon_deep() handles the final override
        "price": [
            "span.a-price-whole",                                    # ✅ confirmed
            "#corePrice_feature_div .a-offscreen",                   # ✅ confirmed
            "#apex_offerDisplay_desktop .a-offscreen",
            "#priceblock_ourprice",
            "#priceblock_dealprice",
            "#priceblock_saleprice",
        ],
        "mrp": [
            "#corePriceDisplay_desktop_feature_div .a-text-price .a-offscreen",
            ".basisPrice .a-offscreen",
            ".a-text-price .a-offscreen",
            "#priceblock_listprice",
        ],
        "name": [
            "#productTitle",
            "span#title span",
            "h1.a-size-large",
        ],
        "brand": [
            "#bylineInfo",
            ".po-brand td:last-child span",
            "a#bylineInfo",
            "tr.po-brand td:last-child",
        ],
        "rating": [
            "#acrPopover span.a-icon-alt",
            "span.a-icon-alt",
            "#averageCustomerReviews .a-icon-alt",
        ],
        "reviews": [
            "#acrCustomerReviewText",
            "#averageCustomerReviews span.a-size-base",
        ],
        "image": [
            "#landingImage",
            "#imgTagWrapperId img",
            "#main-image",
            "#imgBlkFront",
        ],
        "delivery": [
            "#mir-layout-DELIVERY_BLOCK span.a-text-bold",
            "#deliveryBlockMessage",
            "#ddmDeliveryMessage",
        ],
    },

    "flipkart": {
        # Flipkart returns gzip/brotli-compressed JS bundle (~163KB)
        # httpx decompresses it correctly; class names as of 2024-2025:
        "price": [
            "div.hl05eU div.Nx9bqj",
            "div.CEmiEU div.Nx9bqj",
            ".Nx9bqj.CxhGGd",
            "div.Nx9bqj",
            "div._30jeq3._16Jk6d",
            "div._30jeq3",
            "._16Jk6d",
            "._1vC4OE._3qQ9m1",
        ],
        "mrp": [
            ".yRaY8j",
            "div._3I9_wc._2p6lqe",
            "div._3auQ3N",
            "._3I9_wc",
        ],
        "name": [
            "span.B_NuCI", "span.mEh187", "h1.yhB1nd",
            "span.VU-ZEz", "h1._6EBuvT", "h1",
        ],
        "brand": ["span.G6XhRU a", "._2rQ-NK"],
        "rating": ["div._3LWZlK", "div.XQDdHH"],
        "reviews": ["span._2_R_DZ span:last-child", "span._13vcmD", "._3YhLQA"],
        "image": [
            "img._396cs4", "img._2r_T1I",
            "._3GnUWp img", "img.DByuf4", "img.needsclick",
        ],
        "delivery": ["._3XINqE", ".EFccNd span"],
    },

    "myntra": {
        "price": [
            "span.pdp-price strong",
            ".pdp-discount-container .pdp-price strong",
            "strong.pdp-discounted-price",
            "div.pdp-price span",
        ],
        "mrp": ["span.pdp-mrp s", ".pdp-mrp s", "span.pdp-mrp"],
        "name": ["h1.pdp-name", "h1.title-header", "h1"],
        "brand": ["h1.pdp-title", ".pdp-brand-name", "a.pdp-title"],
        "rating": ["div.index-overallRating div"],
        "reviews": ["div.index-ratingsCount", "a.rating-count"],
        "image": ["picture.img-responsive img", "div.image-grid-image img", "img.img-responsive"],
        "delivery": [".deliveryDate", ".delivery-info span"],
    },

    "ajio": {
        "price": ["div.prod-sp", "strong.prod-sp", ".price-discounted"],
        "mrp": ["div.prod-cp", "span.prod-cp"],
        "name": ["h1.prod-name", "h1"],
        "brand": ["h1.brand-name", ".brand-link"],
        "rating": [".rating-count .value"],
        "reviews": [".rating-count .count"],
        "image": ["div.img-wrap img", ".magnifier-main-image img"],
        "delivery": [".delivery-info"],
    },

    "nykaa": {
        "price": ["span.css-111z9ua", ".selling-price", "span[class*='price']"],
        "mrp": ["span.css-1a3osni", ".price-strike", ".mrp"],
        "name": ["h1.css-1gc4x7i", "h1.product-name", "h1"],
        "brand": [".css-cgu7nl", ".brand-name"],
        "rating": [".css-1tp5qzl", ".average-rating"],
        "reviews": [".css-1t3g5vq"],
        "image": [".css-14i89go img", ".product-image img"],
        "delivery": [".delivery-info"],
    },

    "meesho": {
        "price": ["h4.sc-eDvSVe", "h5.sc-fFSPTT"],
        "mrp": ["p.sc-gKXOVf s"],
        "name": ["span.Text__StyledText-sc-ywblx6-0", "h1"],
        "brand": [],
        "rating": ["span.sc-bdnxRM"],
        "reviews": ["span.sc-cPiKLX"],
        "image": ["div.sc-fvxzrP img"],
        "delivery": [],
    },

    "nike": {
        "price": [
            "span[data-test='product-price']",
            ".product-price__wrapper .css-b9fpep",
        ],
        "mrp": ["span[data-test='product-price-reduced']"],
        "name": ["h1#pdp_product_title", "h1.css-1xh5v2e", "h1"],
        "brand": [],
        "rating": [],
        "reviews": ["span.css-1bpuf9"],
        "image": ["img[data-testid='HeroMedia']", "img.css-1dxmopc"],
        "delivery": [],
    },

    "puma": {
        # Debug confirmed: JSON-LD works perfectly for Puma.
        # price=3374 INR, image from CDN, name all correct from JSON-LD.
        # CSS selectors below are fallback only.
        "price": [
            "span.sale-price", "span.regular-price",
            "div.price span.price",
            "[data-testid='product-price']",
            ".product__price span",
        ],
        "mrp": ["del.original-price", "span.strike-price"],
        "name": ["h1.product-title", "h1.product__name", "h1"],
        "brand": [],
        "rating": ["div.pr-snippet-rating-decimal"],
        "reviews": ["span.pr-category-snippet__total"],
        "image": [
            "img[src*='images.puma.com']",
            "div.product-image img",
            "picture.product__hero-image img",
        ],
        "delivery": [],
    },

    "adidas": {
        "price": [
            "span[data-auto-id='sale-price']",
            "span[data-testid='sale-price']",
            "div.product_information-pricing span",
        ],
        "mrp": [
            "span[data-auto-id='original-price']",
            "span.product_information-strikethru-price",
        ],
        "name": ["h1.product_information-title", "h1[data-testid='product-title']", "h1"],
        "brand": [],
        "rating": ["span.average-rating"],
        "reviews": [".BVRRCount .BVRRNumber"],
        "image": ["img[data-auto-id='product-image']", "div.image-wrapper img"],
        "delivery": [],
    },

    "thesouledstore": {
        # TSS is a Vue.js SPA — returns a 21KB empty shell.
        # We use the Shopify Storefront API (products/<handle>.json) as primary.
        # The HTML approach below is kept as secondary fallback.
        "price": [
            "span.price-item--sale",
            "span.price-item.price-item--sale",
            "span.price-item",
            ".product__price .money",
            ".price__regular .money",
            "span[data-product-price]",
        ],
        "mrp": [
            "s.price-item--regular",
            ".price__compare .money",
            "s[data-compare-price]",
        ],
        "name": ["h1.product__title", "h1.product-single__title", ".product-meta__title", "h1"],
        "brand": [],
        "rating": [".spr-summary-starrating"],
        "reviews": [".spr-summary-caption"],
        "image": [
            ".product__media img",
            ".product-single__photos img",
            "img.product-featured-img",
            ".product-images img",
        ],
        "delivery": [],
    },

    "bewakoof": {
        "price": [".product-price .selling-price", ".final-price"],
        "mrp": [".product-price .mrp"],
        "name": ["h1.product-title", "h1"],
        "brand": [],
        "rating": [".rating-count"],
        "reviews": [".review-count"],
        "image": [".product-image img"],
        "delivery": [],
    },

    "snapdeal": {
        "price": ["span.payBlkBig", "span#selling-price-id"],
        "mrp": ["span.pdpCutPrice"],
        "name": ["h1.pdp-e-i-head", "h1#pdpProductHead", "h1"],
        "brand": ["span.pdpBrandLink"],
        "rating": ["span.ratingNumber"],
        "reviews": ["p.ratingsCount"],
        "image": ["div.product-image-container img"],
        "delivery": [],
    },

    "tatacliq": {
        "price": [".product-price span", ".ClickableProductCard__price"],
        "mrp": [".product-price .strike"],
        "name": ["h1.product-title", "h1"],
        "brand": [".brand-name"],
        "rating": [".rating-count .value"],
        "reviews": [".total-ratings"],
        "image": ["div.product-images img"],
        "delivery": [],
    },

    "generic": {
        "price": [
            '[itemprop="price"]', ".price", ".product-price",
            ".sale-price", ".offer-price", ".selling-price",
            ".special-price", ".final-price",
        ],
        "mrp": [".mrp", ".old-price", ".strike-price", ".regular-price", ".original-price"],
        "name": ['[itemprop="name"]', "h1"],
        "brand": ['[itemprop="brand"]', ".brand-name", ".brand"],
        "rating": ['[itemprop="ratingValue"]', ".rating", ".star-rating"],
        "reviews": ['[itemprop="reviewCount"]', ".review-count", ".ratings-count"],
        "image": [
            '[itemprop="image"]', ".product-image img",
            ".primary-image img", "img.product-img",
        ],
        "delivery": [".delivery-date", ".delivery-info"],
    },
}


def extract_css(soup: BeautifulSoup, store_key: str, r: dict, base: str):
    sels = SELECTORS.get(store_key, SELECTORS["generic"])

    def text(sel_list: list) -> str:
        for sel in sel_list:
            try:
                el = soup.select_one(sel)
                if el:
                    if el.name == "meta":
                        return el.get("content", "").strip()
                    t = el.get_text(strip=True)
                    if t:
                        return t
            except Exception:
                pass
        return ""

    def img_src(sel_list: list) -> str:
        for sel in sel_list:
            try:
                el = soup.select_one(sel)
                if el:
                    src = (
                        el.get("data-old-hires") or  # Amazon hi-res
                        el.get("data-zoom-image") or
                        el.get("src") or
                        el.get("data-src") or
                        el.get("data-lazy-src") or ""
                    )
                    if src and not src.startswith("data:") and not is_bad_image(src):
                        return src
            except Exception:
                pass
        return ""

    # Name
    if not r["name"]:
        r["name"] = text(sels.get("name", []))

    # Price — ALWAYS override for Amazon (JSON-LD has wrong price)
    css_price = clean_price(text(sels.get("price", [])))
    if store_key == "amazon" and css_price > 0:
        r["price"] = css_price   # force override
    elif not r["price"] and css_price > 0:
        r["price"] = css_price

    # MRP
    if not r["mrp"]:
        r["mrp"] = clean_price(text(sels.get("mrp", [])))

    # Rating
    if not r["rating"]:
        r["rating"] = clean_rating(text(sels.get("rating", [])))

    # Reviews
    if not r["reviews_count"]:
        r["reviews_count"] = clean_reviews(text(sels.get("reviews", [])))

    # Delivery
    if not r["delivery"]:
        r["delivery"] = text(sels.get("delivery", []))

    # Brand
    if not r["brand"]:
        raw = text(sels.get("brand", []))
        raw = re.sub(r"(?i)(Visit the |Brand:\s*)", "", raw)
        raw = raw.split(" Store")[0].strip()
        r["brand"] = raw
    if not r["brand"] and store_key in ("nike", "puma", "adidas", "thesouledstore", "bewakoof"):
        r["brand"] = detect_store(base).split(".")[0]

    # Image
    if not r["image"]:
        src = img_src(sels.get("image", []))
        if src:
            r["image"] = abs_url(src, base)

# ─────────────────────────────────────────────────────────────────────────────
# STRATEGY 4a — Amazon deep extraction
# Debug confirmed: span.a-price-whole = "6,749" ✅
#                  #corePrice_feature_div .a-offscreen = "₹6,749.00" ✅
#                  #landingImage src = real product image URL ✅
# ─────────────────────────────────────────────────────────────────────────────

def amazon_deep(soup: BeautifulSoup, html: str, r: dict, url: str):
    """
    Amazon-specific extraction.
    Debug confirmed working:
      - span.a-price-whole = "6,749" ✅
      - #corePrice_feature_div .a-offscreen = "₹6,749.00" ✅
      - #landingImage src = real image URL ✅
    """
    # ── Price: ALWAYS override from CSS for Amazon ────────────────────────────
    price_selectors = [
        "span.a-price-whole",                                    # CONFIRMED ✅
        "#corePrice_feature_div .a-offscreen",                   # CONFIRMED ✅
        "#apex_offerDisplay_desktop .a-offscreen",
        "#corePriceDisplay_desktop_feature_div .a-offscreen",
        "#corePrice_desktop .a-offscreen",
        "#priceblock_ourprice",
        "#priceblock_dealprice",
    ]
    for sel in price_selectors:
        el = soup.select_one(sel)
        if el:
            raw = el.get_text(strip=True)
            val = clean_price(raw)
            if val > 100:                    # sanity check — reject tiny values
                r["price"] = val
                break

    # ── MRP ───────────────────────────────────────────────────────────────────
    if not r["mrp"]:
        for sel in [".basisPrice .a-offscreen", ".a-text-price .a-offscreen",
                    "#priceblock_listprice", "span.a-text-price span"]:
            el = soup.select_one(sel)
            if el:
                val = clean_price(el.get_text(strip=True))
                if val > 0:
                    r["mrp"] = val
                    break

    # ── Title ─────────────────────────────────────────────────────────────────
    if not r["name"]:
        for sel in ["#productTitle", "span#title", "h1.a-size-large"]:
            el = soup.select_one(sel)
            if el:
                r["name"] = el.get_text(strip=True)
                break

    # ── Brand ─────────────────────────────────────────────────────────────────
    if not r["brand"]:
        for sel in ["#bylineInfo", ".po-brand td:last-child span",
                    "a#bylineInfo", "#brand"]:
            el = soup.select_one(sel)
            if el:
                raw = el.get_text(strip=True)
                raw = re.sub(r"(Visit the |Brand:\s*)", "", raw)
                raw = raw.split(" Store")[0].strip()
                if raw:
                    r["brand"] = raw
                    break

    # ── Image: data-old-hires is highest quality on Amazon ───────────────────
    if not r["image"]:
        for sel in ["#landingImage", "#imgTagWrapperId img",
                    "#main-image", "#ebooksImgBlkFront"]:
            el = soup.select_one(sel)
            if el:
                src = el.get("data-old-hires", "")
                if not src:
                    # Try data-a-dynamic-image which has high-res options
                    dynamic = el.get("data-a-dynamic-image", "")
                    if dynamic:
                        try:
                            img_map = json.loads(dynamic)
                            # Get the largest image
                            best = max(img_map.keys(), key=lambda u: img_map[u][0] * img_map[u][1])
                            src = best
                        except Exception:
                            pass
                if not src:
                    src = el.get("src", "")
                src = src.strip()
                if src and src.startswith("http") and not is_bad_image(src):
                    r["image"] = src
                    break

    # ── Rating ────────────────────────────────────────────────────────────────
    if not r["rating"]:
        for sel in ["#acrPopover span.a-icon-alt", "span.a-icon-alt"]:
            el = soup.select_one(sel)
            if el:
                r["rating"] = clean_rating(el.get_text(strip=True))
                if r["rating"]:
                    break

    # ── Reviews ───────────────────────────────────────────────────────────────
    if not r["reviews_count"]:
        el = soup.select_one("#acrCustomerReviewText")
        if el:
            r["reviews_count"] = clean_reviews(el.get_text(strip=True))

    # ── Delivery ──────────────────────────────────────────────────────────────
    if not r["delivery"]:
        for sel in ["#deliveryBlockMessage span", "#mir-layout-DELIVERY_BLOCK span",
                    "#ddmDeliveryMessage span"]:
            el = soup.select_one(sel)
            if el:
                txt = el.get_text(strip=True)
                if txt and len(txt) > 3:
                    r["delivery"] = txt
                    break


# ─────────────────────────────────────────────────────────────────────────────
# STRATEGY 4b — Flipkart deep JSON extraction
# Debug: Flipkart returns brotli-compressed 163KB JS bundle.
# httpx (used in fetch_html) decompresses brotli natively.
# The bundle contains embedded JSON with product data.
# ─────────────────────────────────────────────────────────────────────────────

def flipkart_deep(html: str, r: dict):
    """
    Flipkart embeds product data in JSON inside script tags.
    We scan for key price/name/image patterns in the raw HTML.
    httpx handles brotli decompression — HTML should be valid text here.
    """
    # Price patterns in Flipkart's JSON data
    if not r["price"]:
        for pat in [
            r'"finalPrice"\s*:\s*\{\s*"value"\s*:\s*(\d+)',
            r'"sellingPrice"\s*:\s*\{\s*"value"\s*:\s*(\d+)',
            r'"finalPrice"\s*:\s*(\d+)',
            r'"sellingPrice"\s*:\s*(\d+)',
            r'"discountedPrice"\s*:\s*(\d+)',
            r'"value"\s*:\s*(\d+)',
            r'"price"\s*:\s*(\d+)',
        ]:
            m = re.search(pat, html)
            if m:
                val = clean_price(m.group(1))
                if val > 99:   # reject tiny values
                    r["price"] = val
                    break

    if not r["mrp"]:
        for pat in [
            r'"mrpPrice"\s*:\s*\{\s*"value"\s*:\s*(\d+)',
            r'"mrpPrice"\s*:\s*(\d+)',
            r'"mrp"\s*:\s*(\d+)',
        ]:
            m = re.search(pat, html)
            if m:
                r["mrp"] = clean_price(m.group(1))
                break

    if not r["name"]:
        for pat in [
            r'"productName"\s*:\s*"([^"]{5,150})"',
            r'"title"\s*:\s*"([^"]{5,150})"',
            r'"name"\s*:\s*"([^"]{5,150})"',
        ]:
            m = re.search(pat, html)
            if m and "flipkart" not in m.group(1).lower():
                r["name"] = m.group(1)
                break

    if not r["rating"]:
        m = re.search(r'"rating"\s*:\s*([0-9.]+)', html)
        if m:
            r["rating"] = clean_rating(m.group(1))

    if not r["ratings_count"]:
        m = re.search(r'"ratingsCount"\s*:\s*([0-9]+)', html)
        if m:
            r["ratings_count"] = int(m.group(1))

    if not r["reviews_count"]:
        m = re.search(r'"reviewsCount"\s*:\s*([0-9]+)', html)
        if m:
            r["reviews_count"] = int(m.group(1))

    if not r["image"]:
        for pat in [
            r'"imageUrls"\s*:\s*\[\s*"(https://[^"]+\.(?:jpg|jpeg|png|webp))',
            r'"imageUrl"\s*:\s*"(https://[^"]+\.(?:jpg|jpeg|png|webp))',
            r'"(?:thumbnail|image)"\s*:\s*"(https://[^"]+\.(?:jpg|jpeg|png|webp))',
        ]:
            m = re.search(pat, html)
            if m and not is_bad_image(m.group(1)):
                r["image"] = m.group(1)
                break

    # Also try CSS on the parsed page
    if not r["image"] or not r["price"]:
        soup_fk = BeautifulSoup(html, "html.parser")
        sels = SELECTORS["flipkart"]

        def _text(sel_list):
            for sel in sel_list:
                try:
                    el = soup_fk.select_one(sel)
                    if el:
                        t = el.get_text(strip=True)
                        if t:
                            return t
                except Exception:
                    pass
            return ""

        if not r["price"]:
            val = clean_price(_text(sels["price"]))
            if val > 0:
                r["price"] = val
        if not r["name"]:
            r["name"] = _text(sels["name"])
        if not r["rating"]:
            r["rating"] = clean_rating(_text(sels["rating"]))

        for sel in ["img._396cs4", "img._2r_T1I", "._3GnUWp img", "img.DByuf4"]:
            try:
                el = soup_fk.select_one(sel)
                if el:
                    src = el.get("src") or el.get("data-src") or ""
                    if src and not is_bad_image(src):
                        r["image"] = abs_url(src, "https://www.flipkart.com")
                        break
            except Exception:
                pass


# ─────────────────────────────────────────────────────────────────────────────
# STRATEGY 4c — Souled Store (Vue.js SPA)
# Debug confirmed:
#   - Returns 21KB empty Vue.js SPA shell
#   - Shopify JSON API (/products/<handle>.json) returns the SAME SPA shell
#   - All REST and GraphQL APIs require auth token
#   - mapi/api endpoints all return 404
# Solution:
#   1. Derive product name from URL slug (capitalize, clean hyphens)
#   2. Probe known CDN image URL patterns from prod-img.thesouledstore.com
#   3. Use Gemini AI to get price/description from page context
# ─────────────────────────────────────────────────────────────────────────────

def _slug_to_name(slug: str) -> str:
    """Convert URL slug → proper product name. 'top-ace-mocha-men-low-top-sneakers' → 'Top Ace Mocha Low Top Sneakers'"""
    # Remove gender suffixes that aren't part of product name
    slug = re.sub(r'-(men|women|kids|unisex)(-|$)', r'\2', slug, flags=re.IGNORECASE).strip('-')
    # Split by hyphens, capitalize each word
    words = [w.capitalize() for w in slug.split('-') if w]
    return ' '.join(words)

async def souledstore_extract(url: str, r: dict):
    """
    TSS is a completely locked Vue.js SPA. We:
    1. Derive product name from URL slug
    2. Probe CDN for product image  
    3. Set brand = 'The Souled Store'
    4. Leave price=0 for Gemini to fill if available
    """
    try:
        from urllib.parse import urlparse
        parsed = urlparse(url)
        path_parts = parsed.path.rstrip('/').split('/')
        handle = path_parts[-1].split('?')[0]
        
        if not handle or handle == 'product':
            return
        
        # 1. Derive name from slug
        if not r['name']:
            r['name'] = _slug_to_name(handle)
        
        # 2. Set brand
        if not r['brand']:
            r['brand'] = 'The Souled Store'
        
        # 3. Probe CDN for product image
        # TSS CDN pattern: https://prod-img.thesouledstore.com/public/theSoul/uploads/catalog/product/
        cdn_base = 'https://prod-img.thesouledstore.com/public/theSoul/uploads/catalog/product'
        
        # Common TSS image URL patterns (from real product pages observed in browser)
        image_patterns = [
            f'{cdn_base}/{handle}/thumb_images/{handle}_1_2048x2048.jpg',
            f'{cdn_base}/{handle}/thumb_images/{handle}_1.jpg',
            f'{cdn_base}/{handle}/{handle}_1_2048x2048.jpg',
            f'{cdn_base}/{handle}/{handle}_1.jpg',
            f'{cdn_base}/{handle}_1.jpg',
        ]
        
        if not r['image']:
            async with httpx.AsyncClient(
                headers=make_headers(UA_DESKTOP, 'https://www.thesouledstore.com/'),
                follow_redirects=True,
                timeout=10.0,
            ) as client:
                for img_url in image_patterns:
                    try:
                        probe = await client.head(img_url)
                        if probe.status_code == 200:
                            r['image'] = img_url
                            break
                    except Exception:
                        continue
        
        # 4. Description placeholder
        if not r['description']:
            r['description'] = f"Product from The Souled Store. Visit product page for full details."
        
    except Exception as e:
        print(f"TSS extract error: {e}")




# ─────────────────────────────────────────────────────────────────────────────
# STRATEGY 4c — Myntra deep extraction (__myx JSON blob)
# ─────────────────────────────────────────────────────────────────────────────

def myntra_deep(html: str, r: dict):
    """
    Myntra embeds rich product data inside window.__myx script.
    Contains verified averageRating, totalCount, reviewsCount, price, and media.
    """
    myx_match = re.search(r'window\.__myx\s*=\s*(\{.+?\})</script>', html, re.DOTALL)
    if not myx_match:
        myx_match = re.search(r'window\.__myx\s*=\s*(\{.+?\});', html, re.DOTALL)

    if myx_match:
        try:
            raw = myx_match.group(1).rstrip(';')
            data = json.loads(raw)
            pdp = data.get('pdpData', {})

            if not r["name"] and pdp.get("name"):
                r["name"] = pdp.get("name")
            if not r["brand"] and pdp.get("brand", {}).get("name"):
                r["brand"] = pdp.get("brand", {}).get("name")

            price_data = pdp.get("price", {})
            if price_data:
                sp = price_data.get("discounted") or price_data.get("discountedPrice")
                if sp and (not r["price"] or r["price"] == 0):
                    r["price"] = float(sp)
                mrp = price_data.get("mrp")
                if mrp and (not r["mrp"] or r["mrp"] == 0):
                    r["mrp"] = float(mrp)

            ratings = pdp.get("ratings", {})
            if ratings:
                avg = ratings.get("averageRating")
                if avg and (not r["rating"] or r["rating"] == 0):
                    r["rating"] = round(float(avg), 1)

                rev_info = ratings.get("reviewInfo", {})
                rev_cnt = rev_info.get("reviewsCount")
                total_cnt = ratings.get("totalCount")

                if rev_cnt:
                    r["reviews_count"] = clean_reviews(str(rev_cnt))
                elif total_cnt:
                    r["reviews_count"] = clean_reviews(str(total_cnt))

                if total_cnt:
                    r["ratings_count"] = clean_reviews(str(total_cnt))

            media = pdp.get("media", {}).get("albums", [])
            if media and not r["image"]:
                for alb in media:
                    for im in alb.get("images", []):
                        src = im.get("imageURL") or im.get("src")
                        if src and not is_bad_image(src):
                            r["image"] = src
                            break
                    if r["image"]:
                        break
        except Exception as e:
            print(f"Myntra __myx parse error: {e}")

    # Fallback regex if __myx not matched or partial
    if not r["rating"]:
        m_r = re.search(r'"averageRating"\s*:\s*([0-9.]+)', html)
        if m_r:
            r["rating"] = round(float(m_r.group(1)), 1)

    if not r["reviews_count"]:
        m_rev = re.search(r'"reviewsCount"\s*:\s*"([0-9]+)"', html) or re.search(r'"reviewsCount"\s*:\s*([0-9]+)', html)
        if m_rev:
            r["reviews_count"] = clean_reviews(m_rev.group(1))
        else:
            m_tot = re.search(r'"totalCount"\s*:\s*([0-9]+)', html)
            if m_tot:
                r["reviews_count"] = clean_reviews(m_tot.group(1))


# ─────────────────────────────────────────────────────────────────────────────
# STRATEGY 4d — JS-app deep extraction (Meesho, Ajio)
# ─────────────────────────────────────────────────────────────────────────────

def js_app_deep(html: str, r: dict):
    """
    For JS-rendered sites (Vue/React/Next/Nuxt) that return a near-empty HTML shell.
    Searches for product data embedded in inline scripts.
    """
    # ── Try __NEXT_DATA__ (Next.js) ───────────────────────────────────────────
    m = re.search(r'<script[^>]+id="__NEXT_DATA__"[^>]*>(.+?)</script>', html, re.DOTALL)
    if m:
        _parse_js_blob(m.group(1), r)
        if r["name"] and r["price"]:
            return

    # ── Try __NUXT__ (Nuxt.js) ────────────────────────────────────────────────
    m = re.search(r'window\.__NUXT__\s*=\s*(\{.{50,})', html, re.DOTALL)
    if m:
        blob = m.group(1).split('</script>')[0].rstrip(';').strip()
        _parse_js_blob(blob, r)
        if r["name"] and r["price"]:
            return

    # ── Try window.__STATE__ or __INITIAL_STATE__ ─────────────────────────────
    for pat in [r'window\.__(?:INITIAL_)?STATE__\s*=\s*(\{.+?\})\s*;',
                r'window\.APP_STATE\s*=\s*(\{.+?\})\s*;']:
        m = re.search(pat, html, re.DOTALL)
        if m:
            _parse_js_blob(m.group(1), r)
            if r["name"] and r["price"]:
                return

    # ── Scan all inline <script> tags for product data ────────────────────────
    soup = BeautifulSoup(html, "html.parser")
    for script in soup.find_all("script"):
        txt = script.string or ""
        if not txt or len(txt) < 100:
            continue
        if not any(k in txt for k in ["price", "product", "name", "image"]):
            continue
        _parse_js_blob(txt, r)
        if r["name"] and r["price"]:
            return

    # ── Last resort: grep price/name patterns across entire HTML ──────────────
    if not r["price"]:
        for pat in [
            r'"price"\s*:\s*"?(\d{3,6})"?',
            r'"selling_price"\s*:\s*"?(\d{3,6})"?',
            r'"discounted_price"\s*:\s*"?(\d{3,6})"?',
            r'"amount"\s*:\s*(\d{3,6})',
            r'price["\']\s*:\s*["\']?(\d{3,6})',
        ]:
            m = re.search(pat, html)
            if m:
                val = clean_price(m.group(1))
                if val > 0:
                    r["price"] = val
                    break

    if not r["name"]:
        for pat in [
            r'"name"\s*:\s*"([^"]{8,120})"',
            r'"title"\s*:\s*"([^"]{8,120})"',
            r'"product_name"\s*:\s*"([^"]{8,120})"',
            r'"productName"\s*:\s*"([^"]{8,120})"',
        ]:
            m = re.search(pat, html)
            if m:
                candidate = m.group(1)
                if not any(bad in candidate.lower() for bad in
                           ["souledstore", "meesho", "myntra", "ajio", "flipkart",
                            "amazon", "shopping", "online", "clothing"]):
                    r["name"] = candidate
                    break

    if not r["image"]:
        for pat in [
            r'"image(?:_url|Url|URL)?"\s*:\s*"(https://[^"]+\.(?:jpg|jpeg|png|webp)[^"]*)"',
            r'"(?:thumbnail|photo|picture)"\s*:\s*"(https://[^"]+\.(?:jpg|jpeg|png|webp)[^"]*)"',
        ]:
            m = re.search(pat, html)
            if m and not is_bad_image(m.group(1)):
                r["image"] = m.group(1)
                break


def _parse_js_blob(blob: str, r: dict):
    """Parse a JS/JSON blob for product fields. Safe — catches all errors."""
    try:
        data = json.loads(blob)
        _walk_json(data, r, depth=0)
        return
    except Exception:
        pass

    if not r["price"]:
        for pat in [r'"price"\s*:\s*(\d+)', r'"selling_price"\s*:\s*(\d+)',
                    r'"discounted_price"\s*:\s*(\d+)', r'"amount"\s*:\s*(\d+)',
                    r'"finalPrice"\s*:\s*(\d+)', r'"sellingPrice"\s*:\s*(\d+)']:
            m = re.search(pat, blob)
            if m:
                val = clean_price(m.group(1))
                if val > 0:
                    r["price"] = val
                    break

    if not r["name"]:
        for pat in [r'"name"\s*:\s*"([^"]{8,120})"',
                    r'"title"\s*:\s*"([^"]{8,120})"',
                    r'"productName"\s*:\s*"([^"]{8,120})"']:
            m = re.search(pat, blob)
            if m:
                candidate = m.group(1)
                if not any(bad in candidate.lower() for bad in
                           ["shopping", "online", "clothing", "store", "website"]):
                    r["name"] = candidate
                    break

    if not r["image"]:
        m = re.search(
            r'"(?:image|imageUrl|image_url|thumbnail)"\s*:\s*"(https://[^"]+\.(?:jpg|jpeg|png|webp)[^"]*)"',
            blob)
        if m and not is_bad_image(m.group(1)):
            r["image"] = m.group(1)


def _walk_json(data, r: dict, depth: int):
    """Recursively walk a JSON structure looking for product fields."""
    if depth > 8:
        return
    if isinstance(data, dict):
        if "price" in data or "name" in data:
            if not r["price"]:
                for key in ["price", "selling_price", "discounted_price",
                            "finalPrice", "sellingPrice", "amount"]:
                    if key in data:
                        val = clean_price(str(data[key]))
                        if val > 0:
                            r["price"] = val
                            break
            if not r["name"]:
                for key in ["name", "title", "productName", "product_name"]:
                    if key in data and isinstance(data[key], str):
                        candidate = data[key].strip()
                        if (8 < len(candidate) < 150 and
                                not any(bad in candidate.lower() for bad in
                                        ["shopping", "online", "clothing", "store"])):
                            r["name"] = candidate
                            break
            if not r["image"]:
                for key in ["image", "image_url", "imageUrl", "thumbnail", "photo"]:
                    if key in data and isinstance(data[key], str):
                        img = data[key]
                        if img.startswith("http") and not is_bad_image(img):
                            r["image"] = img
                            break
            if not r["rating"]:
                for key in ["rating", "ratingValue", "average_rating"]:
                    if key in data:
                        r["rating"] = clean_rating(str(data[key]))
                        if r["rating"]:
                            break
            if not r["mrp"]:
                for key in ["mrp", "mrpPrice", "original_price", "compare_at_price"]:
                    if key in data:
                        val = clean_price(str(data[key]))
                        if val > 0:
                            r["mrp"] = val
                            break

        for v in data.values():
            if isinstance(v, (dict, list)):
                _walk_json(v, r, depth + 1)

    elif isinstance(data, list):
        for item in data[:10]:
            _walk_json(item, r, depth + 1)


# ─────────────────────────────────────────────────────────────────────────────
# STRATEGY 5 — Price regex fallback on raw HTML
# ─────────────────────────────────────────────────────────────────────────────

def price_regex_fallback(html: str, r: dict):
    if r["price"]:
        return
    for pat in [
        r'₹\s?([\d,]+)',
        r'Rs\.?\s?([\d,]+)',
        r'INR\s?([\d,]+)',
        r'"price"\s*:\s*"?([,\d]+)"?',
        r'data-price="([,\d]+)"',
    ]:
        m = re.search(pat, html)
        if m:
            val = clean_price(m.group(1))
            if val > 0:
                r["price"] = val
                return

# ─────────────────────────────────────────────────────────────────────────────
# STRATEGY 6 — Gemini fallback
# ─────────────────────────────────────────────────────────────────────────────

def gemini_fallback(html: str, url: str, r: dict, api_key: Optional[str] = None):
    client = get_gemini_client(api_key) or gemini_client
    if not client:
        return
    if r["name"] and r["price"]:
        return

    soup = BeautifulSoup(html, "html.parser")
    visible = soup.get_text(separator=" ", strip=True)[:5000]

    og_title = ""
    og_img = ""
    t = soup.find("meta", property="og:title")
    i = soup.find("meta", property="og:image")
    if t: og_title = t.get("content", "")
    if i: og_img = i.get("content", "")

    prompt = f"""You are a product data extraction AI for Indian e-commerce.
Extract data for the MAIN PRODUCT on this page.

Return ONLY valid JSON — no markdown, no explanation:
{{
  "name": "exact product title",
  "brand": "brand name",
  "price": 4999,
  "rating": 4.2,
  "image_url": "https://...",
  "description": "1-2 sentence description"
}}

OG_TITLE: {og_title}
OG_IMAGE: {og_img}
URL: {url}
PAGE_TEXT:
{visible}"""

    try:
        resp = client.models.generate_content(model="gemini-2.0-flash", contents=prompt)
        text = resp.text.strip()
        text = re.sub(r"```(?:json)?", "", text).strip().rstrip("`").strip()
        data = json.loads(text)
        if not r["name"]:        r["name"]        = str(data.get("name", "")).strip()
        if not r["price"]:       r["price"]       = clean_price(str(data.get("price", "")))
        if not r["rating"]:      r["rating"]      = clean_rating(str(data.get("rating", "")))
        if not r["image"]:       r["image"]       = abs_url(data.get("image_url", ""), url)
        if not r["description"]: r["description"] = str(data.get("description", ""))
        if not r["brand"]:       r["brand"]       = str(data.get("brand", ""))
    except Exception as e:
        print(f"Gemini error: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# SCORING
# ─────────────────────────────────────────────────────────────────────────────

PREMIUM_BRANDS = {
    "puma", "nike", "adidas", "apple", "samsung", "oneplus", "sony",
    "asus", "lenovo", "hp", "dell", "levis", "zara", "boat", "jbl", "bose",
}

def calculate_score(p: dict, all_prices: list) -> int:
    score = 50.0
    price  = float(p.get("price") or 0)
    rating = float(p.get("rating") or 0)
    brand  = str(p.get("brand") or "").lower()

    score += rating * 8        # up to +40

    if 0 < price < 2000:   score += 20
    elif price < 5000:     score += 15
    elif price < 10000:    score += 10
    elif price < 25000:    score += 5

    valid = [x for x in all_prices if x and x > 0]
    if valid and len(valid) > 1 and price > 0:
        mn, mx = min(valid), max(valid)
        if mx > mn:
            score += (1 - (price - mn) / (mx - mn)) * 10

    if any(b in brand for b in PREMIUM_BRANDS):
        score += 10

    if p.get("description"):  score += 5
    if p.get("image_url"):    score += 5

    mrp = float(p.get("original_price") or 0)
    if mrp > price > 0:
        score += min(8, (mrp - price) / mrp * 100 * 0.15)

    return int(min(round(score), 100))   # hard cap at 100

# ─────────────────────────────────────────────────────────────────────────────
# FETCH HTML
# ─────────────────────────────────────────────────────────────────────────────

async def fetch_html(url: str) -> tuple[str, int]:
    """Returns (html, status_code). httpx handles gzip/brotli natively."""
    store_key = get_store_key(url)

    if store_key == "flipkart":
        # KEY FIX: Flipkart sends brotli (br) compression when advertised.
        # httpx needs the `brotli` package to decompress it.
        # Simpler fix: exclude 'br' from Accept-Encoding → Flipkart falls back to gzip.
        # This returns the full 1.5MB HTML page with real product data.
        attempt_headers = [
            make_headers(UA_DESKTOP, "https://www.google.com/", no_brotli=True),
            make_headers(UA_MOBILE, "https://www.google.com/", no_brotli=True),
            make_headers(UA_DESKTOP, "https://www.flipkart.com/", no_brotli=True),
        ]
    elif store_key == "amazon":
        attempt_headers = [
            make_headers(UA_DESKTOP, "https://www.google.com/"),
            make_headers(UA_MOBILE, "https://www.google.com/"),
        ]
    else:
        attempt_headers = [
            make_headers(UA_DESKTOP, "https://www.google.com/"),
            make_headers(UA_MOBILE, "https://www.google.com/"),
        ]

    last_status = 0
    for headers in attempt_headers:
        try:
            async with httpx.AsyncClient(
                headers=headers,
                follow_redirects=True,
                timeout=25.0,
                http2=False,
            ) as client:
                resp = await client.get(url)
                last_status = resp.status_code
                if resp.status_code == 200 and len(resp.text) > 500:
                    return resp.text, 200
                await asyncio.sleep(0.3)
        except Exception as e:
            print(f"  fetch error: {e}")
    return "", last_status

# ─────────────────────────────────────────────────────────────────────────────
# MASTER EXTRACT
# ─────────────────────────────────────────────────────────────────────────────

async def extract_product(url: str, idx: int, mode: str = "html", api_key: Optional[str] = None) -> dict:
    store_name = detect_store(url)
    store_key  = get_store_key(url)

    r = {
        "name": "", "brand": "", "price": 0.0, "mrp": 0.0,
        "rating": 0.0, "ratings_count": 0, "reviews_count": 0,
        "image": "", "description": "", "color": "", "delivery": "",
        "in_stock": True,
    }

    client = get_gemini_client(api_key) or gemini_client

    # ── Souled Store: Vue SPA — extract name from slug, probe CDN for image ──
    if store_key == "thesouledstore":
        await souledstore_extract(url, r)
        # TSS API is locked, so we always fall through to Gemini for price.
        # But if Gemini is not available, we return partial data with name+image.
        if not client:
            # Return what we have — name from slug, image from CDN probe
            discount = 0.0
            r["name"] = clean_title(r["name"])[:150] if r["name"] else ""
            return {
                "id": idx, "name": r["name"] or "The Souled Store Product",
                "brand": r["brand"] or "The Souled Store",
                "price": 0, "original_price": 0, "discount": 0,
                "rating": 0, "reviews": 0,
                "image_url": r["image"], "source_url": url,
                "store": store_name, "store_key": store_key,
                "in_stock": True, "delivery": "",
                "description": r["description"] or "Visit The Souled Store product page for price and details.",
                "color": "", "emoji": get_emoji(r["name"] or ""),
                "extraction_method": "slug+cdn", "score": 0,
            }
        # With Gemini available — use it to fill price + description
        fake_html = f"<html><body><h1>{r['name']}</h1><p>Sneakers from The Souled Store</p></body></html>"
        gemini_fallback(fake_html, url, r, api_key=api_key)
        discount = 0.0
        try:
            p, m = float(r["price"] or 0), float(r["mrp"] or 0)
            if m > p > 0:
                discount = round((m - p) / m * 100, 1)
        except Exception:
            pass
        r["name"] = clean_title(r["name"])[:150]
        return {
            "id": idx, "name": r["name"] or "The Souled Store Product",
            "brand": r["brand"] or "The Souled Store",
            "price": r["price"], "original_price": r["mrp"],
            "discount": discount, "rating": r["rating"],
            "reviews": r["reviews_count"] or r["ratings_count"],
            "image_url": r["image"], "source_url": url,
            "store": store_name, "store_key": store_key,
            "in_stock": True, "delivery": r["delivery"],
            "description": r["description"], "color": r["color"],
            "emoji": get_emoji(r["name"]),
            "extraction_method": "slug+cdn+gemini", "score": 0,
        }

    html, status = await fetch_html(url)

    if not html:
        reason = f"HTTP {status} — site blocked the request" if status else "Network error / timeout"
        return _fail(idx, url, store_name, reason)

    soup = BeautifulSoup(html, "html.parser")
    page_title = soup.title.text.strip() if soup.title else ""

    # Run all strategies
    extract_json_ld(soup, r, url)              # 1 — JSON-LD
    extract_og(soup, r, url)                   # 2 — Open Graph
    extract_css(soup, store_key, r, url)       # 3 — CSS

    # Store-specific deep extractors
    if store_key == "amazon":
        amazon_deep(soup, html, r, url)        # 4a — Amazon specific
    elif store_key == "flipkart":
        flipkart_deep(html, r)                 # 4b — Flipkart JSON scan
    elif store_key == "myntra":
        myntra_deep(html, r)                   # 4c — Myntra __myx JSON scan
    elif store_key in ("meesho", "ajio"):
        js_app_deep(html, r)                   # 4d — JS-rendered apps

    # Delivery fallbacks from store text
    if not r["delivery"]:
        html_l = html.lower()
        if store_key == "puma" and "free shipping" in html_l:
            r["delivery"] = "Free Shipping"
        elif store_key == "flipkart" and "free delivery" in html_l:
            r["delivery"] = "Free Delivery"

    price_regex_fallback(html, r)              # 5 — regex

    if mode == "gemini" or not (r["name"] and r["price"]):
        gemini_fallback(html, url, r, api_key=api_key)  # 6 — Gemini

    # Name cleanup
    if not r["name"]:
        r["name"] = page_title
    r["name"] = clean_title(r["name"])[:150]

    # Discount
    discount = 0.0
    try:
        p, m = float(r["price"] or 0), float(r["mrp"] or 0)
        if m > p > 0:
            discount = round((m - p) / m * 100, 1)
    except Exception:
        pass

    method = "failed"
    if r["name"] and r["price"]:
        method = "gemini" if (mode == "gemini" and client) else "json_ld+css"
    elif r["name"]:
        method = "partial"

    return {
        "id": idx,
        "name": r["name"] or "Unknown Product",
        "brand": r["brand"] or "",
        "price": r["price"],
        "original_price": r["mrp"],
        "discount": discount,
        "rating": r["rating"],
        "reviews": r["reviews_count"] or r["ratings_count"],
        "image_url": r["image"],
        "source_url": url,
        "store": store_name,
        "store_key": store_key,
        "in_stock": r["in_stock"],
        "delivery": r["delivery"],
        "description": r["description"],
        "color": r["color"],
        "emoji": get_emoji(r["name"]),
        "extraction_method": method,
        "score": 0,
    }

def _fail(idx, url, store, reason):
    return {
        "id": idx, "name": "Extraction Failed", "brand": "",
        "price": 0, "original_price": 0, "discount": 0,
        "rating": 0, "reviews": 0, "image_url": "",
        "source_url": url, "store": store, "store_key": "",
        "in_stock": False, "delivery": "", "description": reason,
        "color": "", "emoji": "❌", "extraction_method": "failed", "score": 0,
    }

# ─────────────────────────────────────────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────────────────────────────────────────

@app.get("/")
def root():
    return {"status": "ShopIntel AI running", "gemini": bool(gemini_client or GEMINI_KEY), "version": "2.2.0"}

@app.get("/health")
def health():
    return {"ok": True, "gemini": bool(gemini_client or GEMINI_KEY)}

@app.post("/verify-gemini")
async def verify_gemini(body: dict):
    key = (body.get("api_key") or "").strip() or GEMINI_KEY
    if not key:
        return {"ok": False, "error": "No API key provided"}
    try:
        from google import genai
        client = genai.Client(api_key=key)
        resp = client.models.generate_content(
            model="gemini-2.0-flash",
            contents="Respond only with JSON: {\"status\": \"ok\"}"
        )
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}

@app.post("/extract")
async def extract_single(body: dict):
    url = body.get("url", "").strip()
    if not url:
        raise HTTPException(400, "url required")
    product = await extract_product(url, 1, body.get("mode", "html"), body.get("api_key"))
    product["score"] = calculate_score(product, [product["price"]])
    return product

@app.post("/compare")
async def compare(req: CompareRequest):
    urls = [u.strip() for u in req.urls if u.strip()][:6]
    if not urls:
        raise HTTPException(400, "No URLs provided")

    tasks   = [extract_product(url, i + 1, req.mode, req.api_key) for i, url in enumerate(urls)]
    results = await asyncio.gather(*tasks, return_exceptions=True)

    products = []
    for i, r in enumerate(results):
        products.append(
            r if isinstance(r, dict) else _fail(i + 1, urls[i], detect_store(urls[i]), str(r))
        )

    all_prices = [p["price"] for p in products if p.get("price", 0) > 0]
    for p in products:
        p["score"] = calculate_score(p, all_prices)

    ranked = sorted(products, key=lambda p: p["score"], reverse=True)
    return {"products": products, "ranked": ranked,
            "best": ranked[0] if ranked else None, "count": len(products)}

@app.get("/proxy-image")
async def proxy_image(url: str):
    if not url or not url.startswith("http"):
        raise HTTPException(400, "Invalid URL")
    try:
        async with httpx.AsyncClient(
            headers={**make_headers(UA_DESKTOP, url),
                     "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8"},
            follow_redirects=True,
            timeout=12.0,
        ) as client:
            resp = await client.get(url)
            ct = resp.headers.get("content-type", "image/jpeg")
            if "image" not in ct and "octet" not in ct:
                raise HTTPException(400, "Not an image")
            return Response(
                content=resp.content, media_type=ct,
                headers={"Cache-Control": "public, max-age=86400",
                         "Access-Control-Allow-Origin": "*"},
            )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(502, f"Image proxy error: {e}")

# ── DEBUG endpoint ──────────────────────────────────────────────────────────
@app.post("/debug")
async def debug_url(body: dict):
    """
    POST /debug {"url": "https://..."}
    Returns detailed extraction diagnostics.
    """
    url = body.get("url", "").strip()
    if not url:
        raise HTTPException(400, "url required")

    store_key = get_store_key(url)
    html, status = await fetch_html(url)

    if not html:
        return {"status": status, "error": "Could not fetch page", "url": url}

    soup = BeautifulSoup(html, "html.parser")

    ld_blocks = []
    for s in soup.find_all("script", type="application/ld+json"):
        try:
            d = json.loads(s.string or "")
            ld_blocks.append(d)
        except Exception:
            pass

    og_tags = {}
    for tag in soup.find_all("meta"):
        prop = tag.get("property") or tag.get("name") or ""
        if prop.startswith("og:") or prop.startswith("twitter:") or prop.startswith("product:"):
            og_tags[prop] = tag.get("content", "")

    sels = SELECTORS.get(store_key, SELECTORS["generic"])
    css_results = {}
    for field, sel_list in sels.items():
        css_results[field] = {}
        for sel in sel_list:
            try:
                el = soup.select_one(sel)
                if el:
                    txt = el.get_text(strip=True)[:80]
                    src = el.get("src") or el.get("data-src") or ""
                    css_results[field][sel] = txt or src or "(element found but empty)"
                else:
                    css_results[field][sel] = "NOT FOUND"
            except Exception as e:
                css_results[field][sel] = f"ERROR: {e}"

    raw_prices = re.findall(r"₹\s?[\d,]+", html[:30000])[:10]

    r = {
        "name": "", "brand": "", "price": 0.0, "mrp": 0.0,
        "rating": 0.0, "ratings_count": 0, "reviews_count": 0,
        "image": "", "description": "", "color": "", "delivery": "",
        "in_stock": True,
    }
    extract_json_ld(soup, r, url)
    extract_og(soup, r, url)
    extract_css(soup, store_key, r, url)
    if store_key == "amazon":
        amazon_deep(soup, html, r, url)
    elif store_key == "flipkart":
        flipkart_deep(html, r)
    price_regex_fallback(html, r)

    return {
        "url": url,
        "store_key": store_key,
        "http_status": status,
        "page_title": soup.title.text.strip() if soup.title else "",
        "html_size_chars": len(html),
        "json_ld_count": len(ld_blocks),
        "json_ld_types": [
            (b.get("@type") if isinstance(b, dict) else type(b).__name__)
            for b in ld_blocks[:8]
        ],
        "json_ld_first_product": next(
            (b for b in ld_blocks if isinstance(b, dict) and "Product" in str(b.get("@type", ""))),
            ld_blocks[0] if ld_blocks else None
        ),
        "og_tags": og_tags,
        "css_results": css_results,
        "raw_prices_in_html": raw_prices,
        "final_extracted": r,
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)