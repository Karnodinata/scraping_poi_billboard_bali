"""
Scraper for lokasibillboard.com - ALL billboard data (all provinces)
Extracts: kode_produk, nama_brand, jenis, ukuran, alamat, latitude, longitude, provinsi, kota, sumber
Coordinates from data-cur_lat / data-cur_long on each property detail page.
"""
import requests
import pandas as pd
import re
import time
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE_URL = "https://www.lokasibillboard.com"
API_ENDPOINT = f"{BASE_URL}/wp-json/wp/v2/estate_property"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}


def fetch_all_properties(per_page=100):
    """Fetch ALL properties via WP REST API (all provinces)."""
    page = 1
    all_items = []
    while True:
        params = {"per_page": per_page, "page": page}
        resp = requests.get(API_ENDPOINT, params=params, headers=HEADERS, timeout=30)
        if resp.status_code != 200:
            print(f"  Stopped at page {page} (status {resp.status_code})")
            break
        items = resp.json()
        if not items:
            break
        all_items.extend(items)
        print(f"  Page {page}: {len(items)} items (total: {len(all_items)})")
        page += 1
    return all_items


def fetch_taxonomy_map(taxonomy):
    """Fetch taxonomy term names (province, city, category)."""
    mapping = {}
    page = 1
    while True:
        r = requests.get(
            f"{BASE_URL}/wp-json/wp/v2/{taxonomy}",
            params={"per_page": 100, "page": page},
            headers=HEADERS, timeout=30,
        )
        if r.status_code != 200:
            break
        data = r.json()
        if not data:
            break
        for d in data:
            mapping[d["id"]] = d["name"]
        page += 1
    return mapping


def fetch_coordinates(url):
    """Visit property detail page and extract lat/lng from the map div."""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=30)
        soup = BeautifulSoup(resp.text, "html.parser")
        map_div = soup.find("div", id="googleMap_shortcode")
        if map_div:
            lat = map_div.get("data-cur_lat", "")
            lng = map_div.get("data-cur_long", "")
            if lat and lng:
                return float(lat), float(lng)
    except Exception as e:
        pass
    return None, None


def extract_field(content_html, field_name):
    """Extract a field value from content HTML (e.g., Lokasi, Ukuran, Kode Produk)."""
    match = re.search(
        rf"{field_name}\s*</strong>\s*[:\-]?\s*(.*?)</li>",
        content_html, re.IGNORECASE | re.DOTALL,
    )
    if match:
        return re.sub(r"<[^>]+>", "", match.group(1)).strip()
    match = re.search(
        rf"{field_name}\s*[:\-]?\s*(.*?)<",
        content_html, re.IGNORECASE | re.DOTALL,
    )
    if match:
        return re.sub(r"<[^>]+>", "", match.group(1)).strip()
    return ""


def main():
    # Step 1: Fetch taxonomy maps for readable names
    print("Fetching taxonomy maps...")
    province_map = fetch_taxonomy_map("property_county_state")
    city_map = fetch_taxonomy_map("property_city")
    category_map = fetch_taxonomy_map("property_category")
    print(f"  Provinces: {len(province_map)}, Cities: {len(city_map)}, Categories: {len(category_map)}")

    # Step 2: Fetch all properties
    print("\nFetching all properties from API...")
    items = fetch_all_properties()
    print(f"Total properties: {len(items)}")

    if not items:
        print("No properties found.")
        return

    # Step 3: Fetch coordinates from detail pages (parallel, batched)
    print(f"\nFetching coordinates from {len(items)} detail pages (this may take a few minutes)...")
    url_map = {item["id"]: item["link"] for item in items}
    coord_results = {}
    done = 0

    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {
            executor.submit(fetch_coordinates, url): pid
            for pid, url in url_map.items()
        }
        for future in as_completed(futures):
            pid = futures[future]
            lat, lng = future.result()
            coord_results[pid] = (lat, lng)
            done += 1
            if done % 50 == 0:
                print(f"  Progress: {done}/{len(items)}")

    print(f"  Done: {done}/{len(items)}")

    # Step 4: Build records
    records = []
    for item in items:
        pid = item["id"]
        title = item.get("title", {}).get("rendered", "")
        content_html = item.get("content", {}).get("rendered", "")
        class_list = item.get("class_list", [])

        # Resolve taxonomy IDs to names
        provinsi_ids = item.get("property_county_state", [])
        city_ids = item.get("property_city", [])
        cat_ids = item.get("property_category", [])

        provinsi = ", ".join(province_map.get(pid2, str(pid2)) for pid2 in provinsi_ids)
        kota = ", ".join(city_map.get(pid2, str(pid2)) for pid2 in city_ids)
        jenis = ", ".join(category_map.get(pid2, str(pid2)) for pid2 in cat_ids)

        alamat = extract_field(content_html, "Lokasi")
        ukuran = extract_field(content_html, "Ukuran")
        kode = extract_field(content_html, "Kode Produk")
        lat, lng = coord_results.get(pid, (None, None))

        records.append({
            "kode_produk": kode,
            "nama": title,
            "jenis": jenis,
            "ukuran": ukuran,
            "alamat": alamat,
            "provinsi": provinsi,
            "kota": kota,
            "latitude": lat,
            "longitude": lng,
            "sumber": "lokasibillboard.com",
        })

    df = pd.DataFrame(records)

    # Stats
    has_coords = df["latitude"].notna().sum()
    print(f"\nRecords with coordinates: {has_coords}/{len(df)}")

    output_path = "data_lokasi_billboard.xlsx"
    df.to_excel(output_path, index=False)
    print(f"Saved {len(df)} records to {output_path}")


if __name__ == "__main__":
    main()
