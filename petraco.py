import argparse
import html
import json
import re
import time
import urllib.parse
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup
import pandas as pd

BASE_URL = "https://petraco.id"
API_URL = f"{BASE_URL}/wp-json/wp/v2/product"
BALI_LOCATION_SLUG = "bali"
OUTPUT_FILE = "data_petraco.xlsx"
BRAND_NAME = "Petraco (Petraco Media)"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
}

BALI_REGENCIES = {
    "denpasar": "Kota Denpasar",
    "wahidin": "Kota Denpasar",
    "cargo": "Kota Denpasar",
    "bypass ngurah rai": "Kota Denpasar",
    "badung": "Kabupaten Badung",
    "canggu": "Kabupaten Badung",
    "kuta": "Kabupaten Badung",
    "uluwatu": "Kabupaten Badung",
    "jimbaran": "Kabupaten Badung",
    "sunset road": "Kabupaten Badung",
    "gianyar": "Kabupaten Gianyar",
    "mas": "Kabupaten Gianyar",
    "ubud": "Kabupaten Gianyar",
    "tabanan": "Kabupaten Tabanan",
    "buleleng": "Kabupaten Buleleng",
    "karangasem": "Kabupaten Karangasem",
    "klungkung": "Kabupaten Klungkung",
    "bangli": "Kabupaten Bangli",
    "jembrana": "Kabupaten Jembrana",
}

def detect_regency(text):
    text_lower = text.lower()
    for key, regency in BALI_REGENCIES.items():
        if key in text_lower:
            return regency
    return "Bali"

def get_bali_location_id(session):
    """
    Mengambil ID term taksonomi lokasi untuk 'bali' dari WP REST API.
    """
    try:
        resp = session.get(f"{BASE_URL}/wp-json/wp/v2/location", headers=HEADERS, timeout=15)
        resp.raise_for_status()
        for loc in resp.json():
            if loc.get("slug") == BALI_LOCATION_SLUG:
                return loc.get("id")
    except Exception as e:
        print(f"[!] Gagal mengambil ID lokasi dari REST API: {e}")
    # Default ID 98 untuk Bali di petraco.id jika API taksonomi gagal
    return 98

def fetch_bali_products(session, location_id):
    """
    Mengambil semua data produk di Bali via REST API.
    """
    print(f"[*] Mengambil data produk untuk lokasi ID {location_id} (Bali)...")
    products = []
    page = 1
    
    while True:
        params = {
            "location": location_id,
            "per_page": 50,
            "page": page
        }
        try:
            resp = session.get(API_URL, params=params, headers=HEADERS, timeout=15)
            if resp.status_code in [400, 404]:
                break
            resp.raise_for_status()
            data = resp.json()
            if not data:
                break
            products.extend(data)
            page += 1
        except Exception as e:
            print(f"[!] Error saat fetch produk halaman {page}: {e}")
            break
            
    print(f"[+] Total produk ditemukan dari API: {len(products)}")
    return products

def extract_coords_from_embed(iframe_src, session):
    """
    Mengekstrak koordinat lat & lng dari Google Maps embed iframe.
    """
    if not iframe_src:
        return None, None
    try:
        resp = session.get(iframe_src, headers=HEADERS, timeout=10)
        if resp.status_code == 200:
            # Pola 1: [lat, lng] di JavaScript embed
            matches = re.findall(r'\[\s*(-[89]\.\d+)\s*,\s*(11[45]\.\d+)\s*\]', resp.text)
            if matches:
                return float(matches[0][0]), float(matches[0][1])
            # Pola 2: -8.xxx, 115.xxx
            m2 = re.findall(r'(-8\.\d+)[,\s]+(115\.\d+)', resp.text)
            if m2:
                return float(m2[0][0]), float(m2[0][1])
    except Exception as e:
        print(f"    [!] Error mengambil embed maps: {e}")
    return None, None

def scrape_product_detail(product, session):
    """
    Mengunjungi halaman detail produk untuk mengekstrak spesifikasi lengkap dan koordinat.
    """
    url = product.get("link", "")
    raw_title = product.get("title", {}).get("rendered", "")
    title = html.unescape(raw_title).replace("\xa0", " ").strip()
    
    data = {
        "title": title,
        "alamat": "",
        "jenis": "Billboard",
        "ukuran": "",
        "orientasi": "",
        "latitude": None,
        "longitude": None,
        "regency": "Bali",
        "foto": "",
        "url": url
    }
    
    try:
        resp = session.get(url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        
        # 1. Alamat & Query Peta
        iframe = soup.find("iframe")
        iframe_src = iframe.get("src") if iframe else ""
        if iframe_src:
            m_q = re.search(r'maps\?q=([^&"]+)', iframe_src)
            if m_q:
                query_addr = urllib.parse.unquote(m_q.group(1))
                data["alamat"] = query_addr.replace(", Indonesia", "").strip()
                
            # Ambil koordinat GPS dari embed iframe
            lat, lng = extract_coords_from_embed(iframe_src, session)
            data["latitude"] = lat
            data["longitude"] = lng
            
        # Fallback alamat dari judul jika belum ada
        if not data["alamat"]:
            m_at = re.search(r'(?:at|on)\s+(Jl\.?[^–\-|]+)', title, re.I)
            if m_at:
                data["alamat"] = m_at.group(1).strip()
            else:
                data["alamat"] = title

        # 2. Teks Konten & Spesifikasi (Ukuran, Orientasi, Jenis)
        paras = [p.get_text(strip=True) for p in soup.find_all("p") if len(p.get_text(strip=True)) > 20 and "petraco" not in p.get_text(strip=True).lower()[:20]]
        full_text = " ".join(paras)
        
        # Ukuran (misal 5m x 10m atau 4x6)
        m_size = re.search(r'(\d+(?:\.\d+)?\s*(?:m|meter)?\s*[xX*]\s*\d+(?:\.\d+)?\s*(?:m|meter)?)', full_text, re.I)
        if m_size:
            data["ukuran"] = m_size.group(1).strip()
            
        # Orientasi (Vertical / Horizontal)
        m_orient = re.search(r'(vertical|horizontal)', full_text, re.I)
        if m_orient:
            data["orientasi"] = m_orient.group(1).capitalize()
            
        # Jenis Media
        title_lower = title.lower()
        if "led" in title_lower or "digital" in title_lower or "videotron" in title_lower:
            data["jenis"] = "LED Digital OOH / Videotron"
        else:
            data["jenis"] = "Billboard"
            
        # 3. Wilayah (Kabupaten / Kota)
        combined_text = f"{title} {data['alamat']} {url}"
        data["regency"] = detect_regency(combined_text)
        
        # 4. Foto
        for img in soup.find_all("img"):
            dsrc = img.get("data-src") or img.get("src") or ""
            if "uploads" in dsrc and "logo" not in dsrc.lower() and "petraco-1" not in dsrc.lower() and not dsrc.startswith("data:"):
                data["foto"] = dsrc
                break

    except Exception as e:
        print(f"    [!] Error scraping detail ({url}): {e}")
        
    return data

def main():
    parser = argparse.ArgumentParser(description="Scraper Titik Billboard & Media Luar Ruang Petraco Bali (petraco.id)")
    parser.add_argument("--output", type=str, default=OUTPUT_FILE, help=f"Nama file output Excel (default: {OUTPUT_FILE})")
    args = parser.parse_args()

    print("=" * 65)
    print("      SCRAPER TITIK IKLAN LUAR RUANG - PETRACO.ID (BALI)      ")
    print("=" * 65)
    
    start_time = time.time()
    session = requests.Session()
    session.headers.update(HEADERS)
    
    # 1. Ambil ID Lokasi Bali
    bali_id = get_bali_location_id(session)
    
    # 2. Ambil seluruh produk Bali dari API
    products = fetch_bali_products(session, bali_id)
    if not products:
        print("[!] Tidak ada produk yang ditemukan.")
        return

    # 3. Scrape detail dan koordinat tiap titik
    rows = []
    print(f"\n[*] Mengambil detail dan koordinat dari {len(products)} titik iklan...")
    
    for idx, prod in enumerate(products, 1):
        item = scrape_product_detail(prod, session)
        
        lat = item["latitude"]
        lng = item["longitude"]
        title = item["title"]
        addr = item["alamat"]
        
        print(f"[{idx}/{len(products)}] {title[:40]}... ", end="", flush=True)
        if lat is not None and lng is not None:
            print(f"-> OK (Lat: {lat}, Lng: {lng})")
        else:
            print("-> [!] Koordinat KOSONG!")

        rows.append({
            "Nama Produk": title,
            "Alamat": addr,
            "Jenis": item["jenis"],
            "Longitude": lng,
            "Latitude": lat,
            "Brand": BRAND_NAME,
            "Ukuran": item["ukuran"],
            "Orientasi": item["orientasi"],
            "Kabupaten / Kota": item["regency"],
            "Foto": item["foto"],
            "URL": item["url"]
        })
        
        time.sleep(0.3)
        
    df = pd.DataFrame(rows)
    
    # Validasi kelengkapan koordinat
    empty_coords = df[df["Latitude"].isna() | df["Longitude"].isna()]
    if not empty_coords.empty:
        print(f"\n[!] PERINGATAN: Terdapat {len(empty_coords)} data dengan koordinat kosong!")
    else:
        print(f"\n[+] Semua ({len(df)}) data memiliki koordinat Latitude & Longitude lengkap!")

    # Susun urutan kolom
    cols_order = [
        "Nama Produk",
        "Alamat",
        "Jenis",
        "Longitude",
        "Latitude",
        "Brand",
        "Ukuran",
        "Orientasi",
        "Kabupaten / Kota",
        "Foto",
        "URL"
    ]
    df = df[cols_order]

    # Simpan ke Excel
    out_file = args.output
    print(f"\n[*] Menyimpan {len(df)} baris data ke '{out_file}'...")
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        sheet_name = "Data Petraco Bali"
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        
        worksheet = writer.sheets[sheet_name]
        for col in worksheet.columns:
            col_letter = col[0].column_letter
            max_len = max(len(str(cell.value or '')) for cell in col)
            worksheet.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 60)

    elapsed = time.time() - start_time
    print(f"\n[OK] Sukses! File telah disimpan di: {out_file}")
    print(f"[i] Waktu selesai: {elapsed:.2f} detik.")
    print(f"\nRingkasan Data:")
    print(f"- Total baris: {len(df)}")
    print(f"- Tipe Media: {df['Jenis'].value_counts().to_dict()}")
    print(f"- Wilayah: {df['Kabupaten / Kota'].value_counts().to_dict()}")

if __name__ == "__main__":
    main()
