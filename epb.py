import argparse
import html
import json
import re
import time
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup
import pandas as pd

BASE_URL = "https://www.epb.co.id"
START_URL = f"{BASE_URL}/locations/province/bali/"
OUTPUT_FILE = "data_epb.xlsx"
BRAND_NAME = "EPB (PT. Eka Panca Bersama)"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7",
}

def get_listing_pages():
    """
    Mengambil semua URL halaman listing provinsi Bali (termasuk pagination).
    """
    print(f"[*] Mengakses halaman awal: {START_URL}")
    session = requests.Session()
    session.headers.update(HEADERS)
    
    pages = [START_URL]
    try:
        resp = session.get(START_URL, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        
        # Cari link pagination untuk provinsi bali
        page_links = soup.select("a[href*='/locations/province/bali/page/']")
        for link in page_links:
            href = link.get("href")
            full_url = urljoin(BASE_URL, href)
            if full_url not in pages:
                pages.append(full_url)
    except Exception as e:
        print(f"[!] Gagal mengecek pagination: {e}")
        
    return pages

def collect_listings(page_urls):
    """
    Mengambil seluruh metadata listing dari atribut data-location pada halaman-halaman listing.
    """
    session = requests.Session()
    session.headers.update(HEADERS)
    
    unique_items = {}
    
    for idx, page_url in enumerate(page_urls, 1):
        print(f"[*] Mengambil data listing dari Halaman {idx}: {page_url} ...")
        try:
            resp = session.get(page_url, timeout=15)
            resp.raise_for_status()
            soup = BeautifulSoup(resp.text, "html.parser")
            
            elements = soup.select("[data-location]")
            for el in elements:
                raw_json = el.get("data-location")
                if not raw_json:
                    continue
                try:
                    data = json.loads(html.unescape(raw_json))
                    loc_id = data.get("id") or data.get("kode_lokasi") or data.get("slug")
                    if loc_id and loc_id not in unique_items:
                        unique_items[loc_id] = data
                except Exception as ex:
                    print(f"    [!] Error parsing JSON data-location: {ex}")
                    
        except Exception as e:
            print(f"[!] Error saat request ke {page_url}: {e}")
            
    print(f"[+] Berhasil mengumpulkan {len(unique_items)} listing unik di Bali.")
    return list(unique_items.values())

def fetch_detail_coordinates(listing, session):
    """
    Mengunjungi halaman detail listing untuk mengekstrak latitude, longitude,
    serta informasi tambahan seperti kota dari JSON-LD schema atau tag HTML.
    """
    slug = listing.get("slug")
    if not slug:
        return None, None, None
        
    detail_url = f"{BASE_URL}/locations/{slug}"
    lat = None
    lng = None
    city = None
    
    try:
        resp = session.get(detail_url, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        
        # Cara 1: Ambil dari elemen dengan data-lat dan data-lng
        lat_el = soup.find(attrs={"data-lat": True, "data-lng": True})
        if lat_el:
            try:
                lat = float(lat_el["data-lat"])
                lng = float(lat_el["data-lng"])
            except:
                pass
                
        # Cara 2: Ambil dari schema.org JSON-LD (@type: Place)
        for script in soup.find_all("script", type="application/ld+json"):
            if not script.string:
                continue
            try:
                d = json.loads(script.string)
                if isinstance(d, dict) and d.get("@type") == "Place":
                    geo = d.get("geo", {})
                    if lat is None and "latitude" in geo:
                        lat = float(geo["latitude"])
                    if lng is None and "longitude" in geo:
                        lng = float(geo["longitude"])
                    addr = d.get("address", {})
                    if isinstance(addr, dict) and "addressLocality" in addr:
                        city = addr["addressLocality"]
            except:
                pass
                
        # Cara 3: Regex fallback pada HTML jika belum ketemu
        if lat is None or lng is None:
            m_lat = re.search(r'data-lat=["\']([-\d.]+)["\']', resp.text)
            m_lng = re.search(r'data-lng=["\']([-\d.]+)["\']', resp.text)
            if m_lat and m_lng:
                lat = float(m_lat.group(1))
                lng = float(m_lng.group(1))

    except Exception as e:
        print(f"    [!] Gagal mengambil detail ({slug}): {e}")
        
    return lat, lng, city

def main():
    parser = argparse.ArgumentParser(description="Scraper Titik Billboard & Videotron EPB Bali (epb.co.id)")
    parser.add_argument("--output", type=str, default=OUTPUT_FILE, help=f"Nama file output Excel (default: {OUTPUT_FILE})")
    args = parser.parse_args()

    print("=" * 65)
    print("      SCRAPER TITIK IKLAN LUAR RUANG - EPB.CO.ID (BALI)        ")
    print("=" * 65)
    
    start_time = time.time()
    
    # 1. Dapatkan daftar halaman
    page_urls = get_listing_pages()
    print(f"[+] Total halaman listing: {len(page_urls)}")
    
    # 2. Kumpulkan seluruh listing dari halaman
    listings = collect_listings(page_urls)
    if not listings:
        print("[!] Tidak ada listing yang ditemukan.")
        return

    # 3. Kunjungi setiap halaman detail untuk mengambil koordinat lat/lng
    session = requests.Session()
    session.headers.update(HEADERS)
    
    records = []
    print(f"\n[*] Mengambil koordinat dan detail dari {len(listings)} titik iklan...")
    
    for idx, item in enumerate(listings, 1):
        slug = item.get("slug", "")
        kode = item.get("kode_lokasi", "")
        alamat = item.get("alamat", "")
        jenis = item.get("tipe_media", "Billboard")
        ukuran = item.get("ukuran", "")
        orientasi = item.get("orientasi", "")
        traffic = item.get("traffic", 0)
        sisi = item.get("sisi", 1)
        foto = item.get("foto_lokasi", "")
        detail_url = f"{BASE_URL}/locations/{slug}" if slug else ""

        # Nama produk yang informatif
        nama_produk = f"{jenis} - {alamat}"
        if ukuran:
            nama_produk += f" ({ukuran})"

        print(f"[{idx}/{len(listings)}] Memproses: {kode or slug} - {alamat[:40]}...", end=" ", flush=True)
        
        lat, lng, city = fetch_detail_coordinates(item, session)
        
        if lat is not None and lng is not None:
            print(f"-> OK (Lat: {lat}, Lng: {lng})")
        else:
            print("-> [!] Koordinat tidak ditemukan!")

        records.append({
            "Nama Produk": nama_produk,
            "Alamat": alamat,
            "Jenis": jenis,
            "Longitude": lng,
            "Latitude": lat,
            "Brand": BRAND_NAME,
            "Ukuran": ukuran,
            "Orientasi": orientasi,
            "Kode": kode,
            "Kabupaten / Kota": city or "Denpasar",
            "Traffic (Kendaraan/Hari)": traffic,
            "Sisi": sisi,
            "Foto": foto,
            "URL": detail_url
        })
        
        # Jeda ringan agar sopan terhadap server
        time.sleep(0.3)

    df = pd.DataFrame(records)
    
    # Validasi koordinat tidak kosong
    empty_coords = df[df["Latitude"].isna() | df["Longitude"].isna()]
    if not empty_coords.empty:
        print(f"\n[!] PERINGATAN: Terdapat {len(empty_coords)} data dengan koordinat kosong!")
    else:
        print(f"\n[+] Semua ({len(df)}) data memiliki koordinat Latitude & Longitude lengkap!")

    # Urutkan kolom
    cols_order = [
        "Nama Produk",
        "Alamat",
        "Jenis",
        "Longitude",
        "Latitude",
        "Brand",
        "Ukuran",
        "Orientasi",
        "Kode",
        "Kabupaten / Kota",
        "Traffic (Kendaraan/Hari)",
        "Sisi",
        "Foto",
        "URL"
    ]
    df = df[cols_order]

    # Simpan ke Excel
    out_file = args.output
    print(f"\n[*] Menyimpan {len(df)} baris data ke '{out_file}'...")
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        sheet_name = "Data EPB Bali"
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        
        # Auto-adjust lebar kolom
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
    print(f"- Kota: {df['Kabupaten / Kota'].value_counts().to_dict()}")

if __name__ == "__main__":
    main()
