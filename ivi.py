import argparse
import json
import re
import sys
import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE_URL = "https://www.ivi.co.id"
DEFAULT_CATEGORY_URL = f"{BASE_URL}/sewa-billboard/bali/"
OUTPUT_FILE = "data_ivi.xlsx"
BRAND_NAME = "IVI (PT. Inti Vision Indonesia)"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

def get_product_list(target_page=None):
    """
    Mengambil daftar item produk billboard/videotron dari IVI di wilayah Bali.
    Jika target_page diisi (misal: 3), hanya scrape halaman tersebut.
    Jika None, scrape otomatis semua halaman hingga halaman terakhir.
    """
    products = []
    page = 1 if target_page is None else target_page
    
    print(f"[*] Mengumpulkan daftar lokasi dari {BASE_URL} ...")
    
    while True:
        url = f"{DEFAULT_CATEGORY_URL}{page}/" if page > 1 else DEFAULT_CATEGORY_URL
        print(f"[*] Memeriksa halaman {page}: {url} ...")
        
        try:
            resp = requests.get(url, headers=HEADERS, timeout=15)
            if resp.status_code == 404:
                print(f"[i] Halaman {page} tidak ditemukan (404). Selesai menjelajahi halaman.")
                break
            resp.raise_for_status()
        except Exception as e:
            print(f"[!] Gagal mengakses halaman {page}: {e}")
            break

        soup = BeautifulSoup(resp.text, "html.parser")
        cards = soup.select(".billboard-item-card")
        
        if not cards:
            print(f"[i] Tidak ada kartu produk ditemukan pada halaman {page}.")
            break

        print(f"[+] Halaman {page}: ditemukan {len(cards)} titik iklan.")
        
        for card in cards:
            # URL Link
            link_el = card.find("a")
            href = link_el["href"] if link_el and "href" in link_el.attrs else ""
            if href and not href.startswith("http"):
                href = BASE_URL + href

            # Jenis (Billboard / Videotron / Baliho)
            badge_el = card.select_one(".billboard-item-card__badge")
            jenis = badge_el.get_text(strip=True) if badge_el else "Billboard"

            # Ukuran
            ukuran_el = card.select_one(".billboard-item-card__badge--ukuran")
            ukuran = ukuran_el.get_text(strip=True) if ukuran_el else ""

            # Judul / Nama Produk
            title_el = card.select_one(".billboard-item-card__title")
            title = title_el.get_text(strip=True) if title_el else ""

            # Lokasi Ringkas (misal Denpasar)
            loc_el = card.select_one(".billboard-item-card__loc span")
            area = loc_el.get_text(strip=True) if loc_el else ""

            if href:
                products.append({
                    "title": title,
                    "jenis": jenis,
                    "ukuran": ukuran,
                    "area": area,
                    "url": href
                })

        # Jika user hanya meminta 1 halaman tertentu, hentikan loop
        if target_page is not None:
            break

        # Cek apakah ada tombol halaman berikutnya / link pagination berikutnya
        next_page = page + 1
        page = next_page

    print(f"[+] Total titik iklan terkumpul: {len(products)}")
    return products

def scrape_product_detail(item):
    """
    Mengambil data detail produk termasuk Alamat lengkap, Koordinat (Lat/Lng),
    Spesifikasi, dan metadata dari JSON-LD maupun halaman web.
    """
    url = item["url"]
    max_retries = 3
    resp = None

    for attempt in range(max_retries):
        try:
            resp = requests.get(url, headers=HEADERS, timeout=15)
            resp.raise_for_status()
            break
        except Exception as e:
            if attempt == max_retries - 1:
                print(f"[!] Gagal mengambil {url}: {e}")
                return None
            time.sleep(1)

    soup = BeautifulSoup(resp.text, "html.parser")

    lat, lng = None, None
    street_address, locality, region = "", "", ""
    deskripsi = ""

    # 1. Ekstraksi utama melalui schema.org JSON-LD
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string)
            if data.get("@type") == "Place":
                geo = data.get("geo", {})
                lat = geo.get("latitude")
                lng = geo.get("longitude")
                
                addr = data.get("address", {})
                street_address = addr.get("streetAddress", "")
                locality = addr.get("addressLocality", "")
                region = addr.get("addressRegion", "")
                deskripsi = data.get("description", "")
        except Exception:
            pass

    # 2. Fallback regex jika JSON-LD belum memuat koordinat
    if lat is None or lng is None:
        m = re.search(r'latitude\s*([-\d.]+)\s*dan\s*longitude\s*([-\d.]+)', resp.text, re.I)
        if m:
            try:
                lat = float(m.group(1))
                lng = float(m.group(2))
            except ValueError:
                pass

    if lat is None or lng is None:
        m2 = re.search(r'["\']latitude["\']\s*:\s*([-\d.]+)\s*,\s*["\']longitude["\']\s*:\s*([-\d.]+)', resp.text)
        if m2:
            try:
                lat = float(m2.group(1))
                lng = float(m2.group(2))
            except ValueError:
                pass

    # Format alamat lengkap
    alamat_parts = []
    if street_address:
        alamat_parts.append(street_address)
    else:
        alamat_parts.append(item["title"])

    if locality and locality.lower() not in (street_address or "").lower():
        alamat_parts.append(locality)
    if region and region.lower() not in (street_address or "").lower():
        alamat_parts.append(region)

    full_address = ", ".join(alamat_parts)

    # Ekstraksi ukuran dari spesifikasi jika di listing belum ada
    ukuran = item["ukuran"]
    if not ukuran:
        for spec in soup.select(".spec-item, .spec-info"):
            txt = spec.get_text(" ", strip=True)
            m_uk = re.search(r'(\d+m\s*x\s*\d+m)', txt, re.I)
            if m_uk:
                ukuran = m_uk.group(1)
                break

    return {
        "Nama Produk": item["title"],
        "Alamat": full_address,
        "Jenis": item["jenis"],
        "Longitude": float(lng) if lng is not None else None,
        "Latitude": float(lat) if lat is not None else None,
        "Brand": BRAND_NAME,
        "Ukuran": ukuran,
        "Kota / Area": locality or item["area"],
        "URL": url
    }

def main():
    parser = argparse.ArgumentParser(description="Scraper Sewa Billboard IVI Bali (https://www.ivi.co.id)")
    parser.add_argument(
        "--page",
        type=int,
        default=None,
        help="Nomor halaman tertentu yang ingin di-scrape (default: semua halaman Bali 1 s/d selesai)"
    )
    parser.add_argument(
        "--output",
        type=str,
        default=OUTPUT_FILE,
        help=f"Nama file output Excel (default: {OUTPUT_FILE})"
    )
    args = parser.parse_args()

    print("=" * 60)
    print("      SCRAPER BILLBOARD & VIDEOTRON IVI BALI         ")
    print(f"      Target: {DEFAULT_CATEGORY_URL}                ")
    print("=" * 60)

    start_time = time.time()
    products = get_product_list(target_page=args.page)

    if not products:
        print("[!] Tidak ada titik iklan yang ditemukan untuk di-scrape.")
        return

    print(f"\n[*] Mengambil detail untuk {len(products)} titik iklan menggunakan multithreading...")
    scraped_data = []

    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_item = {executor.submit(scrape_product_detail, p): p for p in products}
        completed = 0
        for future in as_completed(future_to_item):
            res = future.result()
            if res:
                scraped_data.append(res)
            completed += 1
            print(f"\r Progress: {completed}/{len(products)} selesai...", end="", flush=True)

    print(f"\n[+] Sukses mengambil data untuk {len(scraped_data)} titik iklan.")

    # Simpan ke DataFrame
    df = pd.DataFrame(scraped_data)

    # Urutkan kolom utama sesuai kebutuhan
    columns_order = [
        "Nama Produk",
        "Alamat",
        "Jenis",
        "Longitude",
        "Latitude",
        "Brand",
        "Ukuran",
        "Kota / Area",
        "URL"
    ]
    columns_order = [c for c in columns_order if c in df.columns]
    df = df[columns_order]

    # Simpan ke file Excel
    out_file = args.output
    print(f"[*] Menyimpan data ke file Excel: {out_file} ...")
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data IVI Bali")
        
        # Auto-adjust column width
        worksheet = writer.sheets["Data IVI Bali"]
        for col in worksheet.columns:
            col_letter = col[0].column_letter
            max_len = max(len(str(cell.value or '')) for cell in col)
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

    elapsed = time.time() - start_time
    print(f"[OK] Berhasil! Data telah disimpan di '{out_file}'.")
    print(f"[i] Waktu eksekusi: {elapsed:.2f} detik.")
    print(f"\nRingkasan Data:")
    print(f"- Total baris: {len(df)}")
    print(f"- Jenis media: {df['Jenis'].value_counts().to_dict()}")

if __name__ == "__main__":
    main()
