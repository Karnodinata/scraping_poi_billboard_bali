import argparse
import os
import re
import time
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup
import pandas as pd

BASE_URL = "https://karyaku.co.id"
START_URL = f"{BASE_URL}/loc/bali/"
OUTPUT_FILE = "data_karyaku.xlsx"
BRAND_NAME = "Utama Mega Karya (Karyaku.co.id)"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7",
}

BALI_REGENCIES = {
    "denpasar": "Kota Denpasar",
    "daun puri": "Kota Denpasar",
    "dauh puri": "Kota Denpasar",
    "badung": "Kabupaten Badung",
    "canggu": "Kabupaten Badung",
    "kuta": "Kabupaten Badung",
    "jimbaran": "Kabupaten Badung",
    "tuban": "Kabupaten Badung",
    "petitenget": "Kabupaten Badung",
    "petingetet": "Kabupaten Badung",
    "seminyak": "Kabupaten Badung",
    "kerobokan": "Kabupaten Badung",
    "batu belig": "Kabupaten Badung",
    "dalung": "Kabupaten Badung",
    "padang luwih": "Kabupaten Badung",
    "kartika plaza": "Kabupaten Badung",
    "bandara": "Kabupaten Badung",
    "kemayoran": "Kabupaten Badung",
    "sunset road": "Kabupaten Badung",
    "gianyar": "Kabupaten Gianyar",
    "ubud": "Kabupaten Gianyar",
    "bedulu": "Kabupaten Gianyar",
    "pantai lebih": "Kabupaten Gianyar",
    "tabanan": "Kabupaten Tabanan",
    "buleleng": "Kabupaten Buleleng",
    "singaraja": "Kabupaten Buleleng",
    "karangasem": "Kabupaten Karangasem",
    "klungkung": "Kabupaten Klungkung",
    "bangli": "Kabupaten Bangli",
    "jembrana": "Kabupaten Jembrana",
    "negara": "Kabupaten Jembrana",
}

def detect_regency(text):
    text_lower = text.lower()
    for key, regency in BALI_REGENCIES.items():
        if key in text_lower:
            return regency
    return "Bali"

def get_all_item_urls():
    """
    Mengambil semua URL item iklan dari halaman arsip lokasi Bali beserta pagination-nya.
    """
    print(f"[*] Mengumpulkan URL item iklan dari: {START_URL}")
    session = requests.Session()
    session.headers.update(HEADERS)
    
    item_urls = []
    page = 1
    
    while True:
        url = f"{START_URL}page/{page}/" if page > 1 else START_URL
        print(f"    -> Memeriksa Halaman {page}: {url} ...")
        
        try:
            resp = session.get(url, timeout=15)
            if resp.status_code == 404:
                break
            resp.raise_for_status()
        except Exception as e:
            print(f"[!] Gagal mengakses halaman {page}: {e}")
            break
            
        soup = BeautifulSoup(resp.text, "html.parser")
        links = [a["href"] for a in soup.find_all("a", href=True) if "/item/" in a["href"]]
        # Deduplikasi per halaman sambil menjaga urutan
        page_items = list(dict.fromkeys(links))
        
        if not page_items:
            break
            
        new_count = 0
        for l in page_items:
            full_l = urljoin(BASE_URL, l)
            if full_l not in item_urls:
                item_urls.append(full_l)
                new_count += 1
                
        print(f"       Ditemukan {new_count} item baru di halaman ini.")
        
        # Cek link ke halaman berikutnya
        if not soup.select(f"a[href*='/page/{page+1}/']"):
            break
            
        page += 1
        time.sleep(0.3)
        
    print(f"[+] Total item iklan unik yang ditemukan: {len(item_urls)}")
    return item_urls

def scrape_item_detail(url, session):
    """
    Mengekstrak seluruh informasi spesifikasi, koordinat, dan detail dari halaman item.
    """
    data = {
        "url": url,
        "title": "",
        "alamat": "",
        "jenis": "Billboard",
        "ukuran": "",
        "muka": "",
        "latitude": None,
        "longitude": None,
        "regency": "Bali",
        "foto": "",
        "deskripsi": ""
    }
    
    try:
        resp = session.get(url, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        
        # 1. Judul / Nama Produk
        h1 = soup.find("h1")
        if h1:
            data["title"] = h1.get_text(strip=True).replace("\xa0", " ")
            
        # 2. Jenis Media
        title_lower = data["title"].lower()
        if "videotron" in title_lower or "video tron" in title_lower:
            data["jenis"] = "Videotron"
        elif "baliho" in title_lower:
            data["jenis"] = "Billboard / Baliho"
        else:
            data["jenis"] = "Billboard"
            
        # 3. Alamat
        addr_el = soup.select_one(".cd-address, [class*='address']")
        if addr_el:
            raw_addr = addr_el.get_text(separator=" ", strip=True)
            # Hilangkan label "Address" di awal
            clean_addr = re.sub(r'^Address\s*:?\s*', '', raw_addr, flags=re.I).strip()
            data["alamat"] = clean_addr.replace("\xa0", " ")
        else:
            # Fallback dari judul jika format "di Bali, Jl. ..."
            m_title_addr = re.search(r'di Bali[,\s]*(Jl\..*)', data["title"], re.I)
            if m_title_addr:
                data["alamat"] = m_title_addr.group(1).strip()
                
        # 4. Koordinat (Latitude & Longitude)
        # Cara A: Dari link Google Maps directions (destination=lat,lng)
        maps_a = soup.find("a", href=re.compile(r'destination=([-\d.]+),([-\d.]+)'))
        if maps_a:
            m_coords = re.search(r'destination=([-\d.]+),([-\d.]+)', maps_a["href"])
            if m_coords:
                try:
                    data["latitude"] = float(m_coords.group(1))
                    data["longitude"] = float(m_coords.group(2))
                except:
                    pass
                    
        # Cara B: Dari elemen .cd-gps
        if data["latitude"] is None or data["longitude"] is None:
            gps_el = soup.select_one(".cd-gps, [class*='gps']")
            if gps_el:
                gps_txt = gps_el.get_text(separator=" ", strip=True)
                clean_gps = re.sub(r'^GPS\s*:?\s*', '', gps_txt, flags=re.I).strip()
                parts = clean_gps.split(",")
                if len(parts) == 2:
                    try:
                        data["latitude"] = float(parts[0].strip())
                        data["longitude"] = float(parts[1].strip())
                    except:
                        pass
                        
        # 5. Ukuran & Muka dari teks konten
        full_text = soup.get_text(separator="\n", strip=True)
        m_size = re.search(r'Ukuran\s*:?\s*([0-9.,]+\s*[xX*]\s*[0-9.,]+(?:\s*m)?)', full_text, re.I)
        if m_size:
            data["ukuran"] = m_size.group(1).strip()
            
        m_muka = re.search(r'([1-4]\s*Muka)', full_text, re.I)
        if m_muka:
            data["muka"] = m_muka.group(1).strip()
            
        # 6. Kabupaten / Kota
        combined_text = f"{data['title']} {data['alamat']} {url}"
        data["regency"] = detect_regency(combined_text)
        
        # 7. Foto
        img_el = soup.select_one(".ctdl-item-featured-image img, .entry-content img, article img")
        if img_el and img_el.has_attr("src"):
            data["foto"] = img_el["src"]

    except Exception as e:
        print(f"    [!] Error scraping {url}: {e}")
        
    return data

def main():
    parser = argparse.ArgumentParser(description="Scraper Titik Billboard & Videotron Karyaku.co.id (Bali)")
    parser.add_argument("--output", type=str, default=OUTPUT_FILE, help=f"Nama file output Excel (default: {OUTPUT_FILE})")
    args = parser.parse_args()

    print("=" * 65)
    print("      SCRAPER TITIK IKLAN LUAR RUANG - KARYAKU.CO.ID (BALI)    ")
    print("=" * 65)
    
    start_time = time.time()
    
    # 1. Kumpulkan semua URL item
    item_urls = get_all_item_urls()
    if not item_urls:
        print("[!] Tidak ada item yang ditemukan.")
        return
        
    # 2. Scrape detail dari masing-masing item
    session = requests.Session()
    session.headers.update(HEADERS)
    
    rows = []
    print(f"\n[*] Mengambil data dan koordinat dari {len(item_urls)} titik iklan...")
    
    for idx, url in enumerate(item_urls, 1):
        item_data = scrape_item_detail(url, session)
        
        lat = item_data["latitude"]
        lng = item_data["longitude"]
        title = item_data["title"]
        addr = item_data["alamat"]
        
        print(f"[{idx}/{len(item_urls)}] {title[:35]}... ", end="", flush=True)
        if lat is not None and lng is not None:
            print(f"-> OK (Lat: {lat}, Lng: {lng})")
        else:
            print("-> [!] Koordinat KOSONG!")
            
        # Bentuk baris sesuai format standar POI-ADS-BALI
        rows.append({
            "Nama Produk": title,
            "Alamat": addr,
            "Jenis": item_data["jenis"],
            "Longitude": lng,
            "Latitude": lat,
            "Brand": BRAND_NAME,
            "Ukuran": item_data["ukuran"],
            "Orientasi": item_data["muka"],
            "Kabupaten / Kota": item_data["regency"],
            "Foto": item_data["foto"],
            "URL": url
        })
        
        time.sleep(0.2)
        
    df = pd.DataFrame(rows)
    
    # Verifikasi koordinat
    empty_coords = df[df["Latitude"].isna() | df["Longitude"].isna()]
    if not empty_coords.empty:
        print(f"\n[!] PERINGATAN: Terdapat {len(empty_coords)} data dengan koordinat kosong!")
    else:
        print(f"\n[+] Semua ({len(df)}) data memiliki koordinat Latitude & Longitude lengkap!")

    # Atur urutan kolom
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

    # Simpan ke Excel dengan openpyxl auto-adjust column width
    out_file = args.output
    print(f"\n[*] Menyimpan {len(df)} baris data ke '{out_file}'...")
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        sheet_name = "Data Karyaku Bali"
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
