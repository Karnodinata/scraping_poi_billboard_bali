"""
Scraper for cvbalimedia.com - Billboard Locations in Bali
=========================================================
Site: https://www.cvbalimedia.com/billboard
Platform: Wix
Data: 3 billboard listings with locations in Canggu and Munggu area
"""

import requests
from bs4 import BeautifulSoup
import re
import time
import pandas as pd

# ── Configuration ──────────────────────────────────────────────────────
BASE_URL = "https://www.cvbalimedia.com/billboard"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}
OUTPUT_FILE = "data_cvbalimedia.xlsx"
BRAND = "CV Bali Media"


def resolve_goo_gl(short_url):
    """Resolve a Google Maps short link and extract coordinates."""
    try:
        resp = requests.get(short_url, headers=HEADERS, timeout=20, allow_redirects=True)
        final_url = resp.url
        # Pattern: @lat,lng
        coords = re.findall(r'@(-?\d+\.\d+),(\d+\.\d+)', final_url)
        if coords:
            return float(coords[0][0]), float(coords[0][1])
        # Pattern: !3d...!4d...
        d_coords = re.findall(r'!3d(-?\d+\.\d+)!4d(\d+\.\d+)', final_url)
        if d_coords:
            return float(d_coords[0][0]), float(d_coords[0][1])
    except Exception as e:
        print(f"  [WARN] Failed to resolve {short_url}: {e}")
    return None, None


def geocode_location(query):
    """Geocode a location using Nominatim (OSM)."""
    try:
        url = f"https://nominatim.openstreetmap.org/search?q={query}&format=json&limit=1"
        resp = requests.get(url, headers={"User-Agent": "POI-ADS-BALI/1.0"}, timeout=10)
        results = resp.json()
        if results:
            return float(results[0]["lat"]), float(results[0]["lon"])
    except Exception as e:
        print(f"  [WARN] Geocoding failed for '{query}': {e}")
    return None, None


def scrape_billboard_page():
    """Scrape billboard listings from the Wix page."""
    print(f"[1/3] Fetching {BASE_URL} ...")
    resp = requests.get(BASE_URL, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    print(f"  Page loaded, {len(resp.text)} bytes")

    # ── Extract billboard items ───────────────────────────────────────
    billboard_texts = []
    for tag in soup.find_all(["p", "span"]):
        text = tag.get_text(strip=True)
        if text.startswith("Billboard") and ("m " in text or "x" in text):
            if text not in billboard_texts:
                billboard_texts.append(text)

    print(f"\n[2/3] Found {len(billboard_texts)} billboard listings:")
    for t in billboard_texts:
        print(f"  - {t}")

    # ── Extract Google Maps links ─────────────────────────────────────
    goo_links = {}
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "maps.app.goo.gl" in href or "google.com/maps" in href:
            parent = a.find_parent()
            while parent:
                prev = parent.find_previous_sibling()
                if prev:
                    prev_text = prev.get_text(strip=True)
                    for bt in billboard_texts:
                        if bt in prev_text or prev_text in bt:
                            goo_links[bt] = href
                            break
                parent = parent.find_parent()
            if not any(bt in goo_links for bt in billboard_texts):
                goo_links[billboard_texts[0]] = href

    # ── Extract image URLs ────────────────────────────────────────────
    image_map = {}
    imgs = soup.find_all("img")
    billboard_imgs = [img for img in imgs
                      if img.get("alt", "") and
                      any(kw in img.get("alt", "").lower()
                          for kw in ["billboard", "canggu", "munggu"])]
    for i, bt in enumerate(billboard_texts):
        if i < len(billboard_imgs):
            image_map[bt] = billboard_imgs[i].get("src", "")

    # ── Get coordinates ───────────────────────────────────────────────
    print(f"\n[3/3] Resolving coordinates ...")
    records = []

    for bt in billboard_texts:
        # Parse size from text like "Billboard 5 x 10 m Jl Raya Canggu"
        size_match = re.search(r'(\d+)\s*x\s*(\d+)\s*m', bt)
        ukuran = f"{size_match.group(1)} x {size_match.group(2)} m" if size_match else ""

        # Parse location from text (everything after "m ")
        loc_match = re.search(r'\d+\s*m\s+(.*)', bt)
        lokasi = loc_match.group(1).strip() if loc_match else bt

        lat, lng = None, None

        # Method 1: Google Maps link
        if bt in goo_links:
            print(f"  Resolving Google Maps link for: {bt}")
            lat, lng = resolve_goo_gl(goo_links[bt])
            if lat:
                print(f"    -> {lat}, {lng} (from Google Maps)")

        # Method 2: Geocoding fallback
        if not lat:
            query = f"{lokasi}, Bali, Indonesia"
            print(f"  Geocoding: {query}")
            lat, lng = geocode_location(query)
            if lat:
                print(f"    -> {lat}, {lng} (from geocoding)")
            time.sleep(1)  # Rate limit

        records.append({
            "Nama Produk": bt,
            "Alamat": f"{lokasi}, Bali",
            "Jenis": "Billboard",
            "Longitude": lng,
            "Latitude": lat,
            "Brand": BRAND,
            "Ukuran": ukuran,
            "Orientasi": "Horizontal",
            "Image URL": image_map.get(bt, ""),
            "Source URL": BASE_URL,
        })

    return records


def main():
    records = scrape_billboard_page()

    if not records:
        print("\n[ERROR] No records found!")
        return

    df = pd.DataFrame(records)

    # Ensure coordinate completeness
    total = len(df)
    has_coords = df["Latitude"].notna().sum()
    print(f"\n{'='*60}")
    print(f"Total records:       {total}")
    print(f"With coordinates:    {has_coords}/{total} ({has_coords/total*100:.0f}%)")
    print(f"{'='*60}")

    # Save to Excel
    df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    print(f"\n[OK] Saved to {OUTPUT_FILE}")
    print(f"\nData preview:")
    print(df[["Nama Produk", "Alamat", "Ukuran", "Latitude", "Longitude"]].to_string(index=False))


if __name__ == "__main__":
    main()
