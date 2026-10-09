import time
import re
import requests
import pandas as pd
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE_URL = "https://thebaligen.com"
PRODUK_URL = f"{BASE_URL}/produk"
OUTPUT_FILE = "data_baligen.xlsx"
BRAND_NAME = "The Bali Gen"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

def get_product_list():
    """Mengambil daftar URL produk dari halaman utama /produk."""
    print(f"[*] Mengakses {PRODUK_URL} ...")
    response = requests.get(PRODUK_URL, headers=HEADERS, timeout=15)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")
    items = soup.select(".project-item")
    print(f"[+] Ditemukan {len(items)} produk.")

    products = []
    for item in items:
        # Kategori / Jenis
        cat_elem = item.select_one(".project-category")
        if cat_elem:
            jenis = cat_elem.get_text(strip=True)
        else:
            jenis = item.get("data-category", "").title()

        # Judul produk
        title_elem = item.find("h4")
        title = title_elem.get_text(strip=True) if title_elem else ""

        # Link detail
        link_elem = item.find("a")
        href = link_elem["href"] if link_elem and "href" in link_elem.attrs else ""
        if href and not href.startswith("http"):
            href = BASE_URL + href

        if href:
            products.append({
                "title": title,
                "jenis": jenis,
                "url": href
            })

    return products

def scrape_product_detail(product_info):
    """Mengambil detail produk: Alamat, Koordinat (Lat, Long), Ukuran, dll."""
    url = product_info["url"]
    
    # Retry mechanism
    max_retries = 3
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

    # Ambil Alamat dan Ukuran dari card info
    alamat = ""
    ukuran = ""
    for item in soup.select(".contact-info-item"):
        h5 = item.find("h5")
        span = item.find("span")
        if h5 and span:
            label = h5.get_text(strip=True).lower()
            val = span.get_text(strip=True)
            if "alamat" in label:
                alamat = val
            elif "ukuran" in label:
                ukuran = val

    # Ambil Latitude dan Longitude
    lat, lng = None, None
    # 1. Regex dari javascript google maps (var pos = { lat: ..., lng: ... };)
    m = re.search(r"var pos\s*=\s*\{\s*lat:\s*([-\d.]+),\s*lng:\s*([-\d.]+)\s*\};", resp.text)
    if m:
        try:
            lat = float(m.group(1))
            lng = float(m.group(2))
        except ValueError:
            pass

    # 2. Alternatif: dari teks "Lat: ..., Lng: ..."
    if lat is None or lng is None:
        m2 = re.search(r"Lat:\s*([-\d.]+),\s*Lng:\s*([-\d.]+)", resp.text)
        if m2:
            try:
                lat = float(m2.group(1))
                lng = float(m2.group(2))
            except ValueError:
                pass

    # 3. Alternatif: dari URL google maps
    if lat is None or lng is None:
        m3 = re.search(r"google\.com/maps\?q=([-\d.]+),([-\d.]+)", resp.text)
        if m3:
            try:
                lat = float(m3.group(1))
                lng = float(m3.group(2))
            except ValueError:
                pass

    return {
        "Nama Produk": product_info["title"],
        "Jenis": product_info["jenis"],
        "Alamat": alamat,
        "Latitude": lat,
        "Longitude": lng,
        "Brand": BRAND_NAME,
        "Ukuran": ukuran,
        "URL": url
    }

def main():
    print("=" * 60)
    print("       SCRAPER IKLAN LUAR RUANG - THE BALI GEN        ")
    print("=" * 60)

    start_time = time.time()
    products = get_product_list()

    if not products:
        print("[!] Tidak ada produk yang ditemukan untuk di-scrape.")
        return

    print(f"\n[*] Mengambil detail untuk {len(products)} produk menggunakan multithreading...")
    scraped_data = []

    # Scraping concurrent untuk efisiensi
    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_prod = {executor.submit(scrape_product_detail, p): p for p in products}
        completed = 0
        for future in as_completed(future_to_prod):
            res = future.result()
            if res:
                scraped_data.append(res)
            completed += 1
            print(f"\r Progress: {completed}/{len(products)} selesai...", end="", flush=True)

    print(f"\n[+] Sukses mengambil data untuk {len(scraped_data)} produk.")

    # Simpan ke DataFrame
    df = pd.DataFrame(scraped_data)

    # Urutkan kolom sesuai kebutuhan
    columns_order = [
        "Nama Produk",
        "Alamat",
        "Jenis",
        "Longitude",
        "Latitude",
        "Brand",
        "Ukuran",
        "URL"
    ]
    # Filter kolom yang ada
    columns_order = [c for c in columns_order if c in df.columns]
    df = df[columns_order]

    # Simpan ke Excel dengan styling openpyxl
    print(f"[*] Menyimpan data ke file Excel: {OUTPUT_FILE} ...")
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data BaliGen")
        
        # Auto-adjust column width
        worksheet = writer.sheets["Data BaliGen"]
        for col in worksheet.columns:
            col_letter = col[0].column_letter
            max_len = max(len(str(cell.value or '')) for cell in col)
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

    elapsed = time.time() - start_time
    print(f"[OK] Berhasil! Data telah disimpan di '{OUTPUT_FILE}'.")
    print(f"[i] Waktu eksekusi: {elapsed:.2f} detik.")
    print(f"\nRingkasan Data:")
    print(f"- Total baris: {len(df)}")
    print(f"- Jenis media: {df['Jenis'].value_counts().to_dict()}")

if __name__ == "__main__":
    main()
