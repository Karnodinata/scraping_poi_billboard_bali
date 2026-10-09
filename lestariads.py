import argparse
import html
import json
import os
import re
import time
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup
import pandas as pd

BASE_URL = "https://www.lestariads.com"
START_SEARCH_URL = (
    f"{BASE_URL}/search-space/?sort=-&jenis=-&transit=-&pos=-&ukuran_billboard=-"
    f"&dimensi=-&light=-&kota=-&kota=-&prov=Bali&submit=&k_hid_search=search&nc=1"
)
OUTPUT_FILE = "data_lestariads.xlsx"
BRAND_NAME = "Lestari Ads"

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
    "badung": "Kabupaten Badung",
    "kuta": "Kabupaten Badung",
    "seminyak": "Kabupaten Badung",
    "canggu": "Kabupaten Badung",
    "nusa dua": "Kabupaten Badung",
    "tuban": "Kabupaten Badung",
    "kerobokan": "Kabupaten Badung",
    "ungasan": "Kabupaten Badung",
    "benoa": "Kabupaten Badung",
    "gianyar": "Kabupaten Gianyar",
    "ubud": "Kabupaten Gianyar",
    "tabanan": "Kabupaten Tabanan",
    "baturiti": "Kabupaten Tabanan",
    "selemadeg": "Kabupaten Tabanan",
    "pupuan": "Kabupaten Tabanan",
    "kediri": "Kabupaten Tabanan",
    "bangli": "Kabupaten Bangli",
    "kintamani": "Kabupaten Bangli",
    "klungkung": "Kabupaten Klungkung",
    "semarapura": "Kabupaten Klungkung",
    "karangasem": "Kabupaten Karangasem",
    "karang asem": "Kabupaten Karangasem",
    "manggis": "Kabupaten Karangasem",
    "rendang": "Kabupaten Karangasem",
    "buleleng": "Kabupaten Buleleng",
    "singaraja": "Kabupaten Buleleng",
    "sukasada": "Kabupaten Buleleng",
    "sawan": "Kabupaten Buleleng",
    "banjar": "Kabupaten Buleleng",
    "gerokgak": "Kabupaten Buleleng",
    "lovina": "Kabupaten Buleleng",
    "jembrana": "Kabupaten Jembrana",
}

# Koordinat landmark untuk iklan transit / bandara
TRANSIT_FALLBACK_COORDS = {
    "bike-billboard-advertising": {
        "lat": -8.718000,
        "lng": 115.169000,
        "alamat": "Kawasan Wisata Kuta - Seminyak, Badung, Bali",
        "kota": "Kabupaten Badung"
    },
    "sewa-iklan-bus-bali-airport-connexion-bac-bali": {
        "lat": -8.748152,
        "lng": 115.167172,
        "alamat": "Bandara Internasional I Gusti Ngurah Rai, Tuban, Badung, Bali",
        "kota": "Kabupaten Badung"
    },
    "sewa-iklan-di-bandara-internasional-i-gusti-ngurah-rai-bali": {
        "lat": -8.748152,
        "lng": 115.167172,
        "alamat": "Bandara Internasional I Gusti Ngurah Rai, Tuban, Badung, Bali",
        "kota": "Kabupaten Badung"
    },
    "space-iklan-scooter-trailer-billboard-bali": {
        "lat": -8.683000,
        "lng": 115.157000,
        "alamat": "Koridor Kuta - Seminyak - Canggu, Badung, Bali",
        "kota": "Kabupaten Badung"
    },
    "space-iklan-scooter-trailer-digital-billboard-bali": {
        "lat": -8.683000,
        "lng": 115.157000,
        "alamat": "Koridor Kuta - Seminyak - Canggu, Badung, Bali",
        "kota": "Kabupaten Badung"
    },
    "space-iklan-transit-trans-metro-dewata-kota-denpasar-bali": {
        "lat": -8.629334,
        "lng": 115.196551,
        "alamat": "Rute Trans Metro Dewata (Hub Terminal Ubung), Kota Denpasar, Bali",
        "kota": "Kota Denpasar"
    }
}

def detect_regency(text):
    text_lower = text.lower()
    for key, regency in BALI_REGENCIES.items():
        if key in text_lower:
            return regency
    return "Bali"

def get_bali_space_urls(session):
    """
    Mengumpulkan URL space resmi di Bali dengan menelusuri search page dan relasi detail.
    Hanya mengembalikan link internal unik www.lestariads.com/space/*.html.
    """
    print(f"[*] Menelusuri seluruh titik iklan Bali dari: {START_SEARCH_URL}")
    to_visit = [START_SEARCH_URL]
    visited = set()
    bali_spaces = set()
    
    while to_visit:
        curr = to_visit.pop(0)
        if curr in visited:
            continue
        visited.add(curr)
        
        try:
            resp = session.get(curr, timeout=12)
            if resp.status_code != 200:
                continue
            soup = BeautifulSoup(resp.text, "html.parser")
            
            # Jika halaman saat ini adalah space Bali
            if "lestariads.com/space/" in curr and curr.endswith(".html"):
                txt = soup.get_text(separator=" ", strip=True).lower()
                if any(x in curr.lower() for x in ["bali", "denpasar", "badung", "gianyar", "tabanan", "bangli", "klungkung", "karangasem", "karang-asem", "nusa-dua", "sunset-road", "buleleng", "singaraja"]) or "bali" in txt:
                    bali_spaces.add(curr)
                    print(f"    [+] Terdata ({len(bali_spaces)}): {curr.split('/')[-1]}")
            
            # Temukan link internal space lainnya di halaman
            for a in soup.find_all("a", href=True):
                href = a["href"]
                # Jika share link, ekstrak url aslinya
                m_sub = re.search(r'(https://www\.lestariads\.com/space/[a-zA-Z0-9\-_.]+\.html)', href)
                if m_sub:
                    target_url = m_sub.group(1)
                elif href.startswith("/space/") and href.endswith(".html"):
                    target_url = urljoin(BASE_URL, href)
                elif href.startswith("https://www.lestariads.com/space/") and href.endswith(".html"):
                    target_url = href
                else:
                    target_url = None
                    
                if target_url and target_url not in visited and target_url not in to_visit:
                    if any(x in target_url.lower() for x in ["bali", "denpasar", "badung", "gianyar", "tabanan", "bangli", "klungkung", "karangasem", "karang-asem", "nusa-dua", "sunset", "teuku-umar", "bedugul", "bike-billboard", "buleleng", "singaraja"]) or "search-space" in curr:
                        to_visit.append(target_url)
                        
        except Exception as e:
            print(f"    [!] Error crawling {curr}: {e}")
            
    sorted_spaces = sorted(list(bali_spaces))
    print(f"[+] Total titik iklan Bali ditemukan: {len(sorted_spaces)}")
    return sorted_spaces

def scrape_space_detail(url, session):
    """
    Mengekstrak seluruh detail spesifikasi, tabel, koordinat, dan foto dari halaman space.
    """
    slug = url.split("/")[-1].replace(".html", "")
    
    data = {
        "title": "",
        "alamat": "",
        "jenis": "Billboard",
        "ukuran": "",
        "orientasi": "",
        "sisi": "",
        "penerangan": "",
        "traffic": "",
        "poi": "",
        "latitude": None,
        "longitude": None,
        "regency": "Bali",
        "foto": "",
        "url": url
    }
    
    try:
        resp = session.get(url, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        
        # 1. Judul
        t_el = soup.select_one("h1, h2, .entry-title")
        raw_title = t_el.get_text(strip=True) if t_el else soup.title.string
        data["title"] = html.unescape(raw_title).replace("\xa0", " ").strip()
        
        # 2. Parsing Tabel Spesifikasi
        spec_table = {}
        for tr in soup.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) >= 2:
                k = tds[0].get_text(strip=True).lower()
                v = tds[1].get_text(separator=" ", strip=True)
                spec_table[k] = v

        # Tipe Media, Ukuran, Sisi
        raw_media = spec_table.get("tipe media", "")
        if raw_media:
            parts = [p.strip() for p in raw_media.split(",") if p.strip()]
            if len(parts) >= 1:
                data["jenis"] = parts[0]
            if len(parts) >= 2:
                for p in parts[1:]:
                    if any(c.isdigit() for c in p) and ("m" in p.lower() or "x" in p.lower()):
                        data["ukuran"] = p
                    elif "sisi" in p.lower():
                        data["sisi"] = p
        else:
            t_lower = data["title"].lower()
            if "videotron" in t_lower or "led" in t_lower:
                data["jenis"] = "Videotron"
            elif "baliho" in t_lower:
                data["jenis"] = "Baliho"
            elif "neon box" in t_lower:
                data["jenis"] = "Neon Box"
            elif "transit" in t_lower or "bus" in t_lower:
                data["jenis"] = "Transit Ad"

        if not data["ukuran"] and "dimensi" in spec_table:
            data["ukuran"] = spec_table["dimensi"]

        data["orientasi"] = spec_table.get("orientasi", "")
        data["penerangan"] = spec_table.get("penerangan", "")
        data["traffic"] = spec_table.get("estimasi traffic", "")
        data["poi"] = spec_table.get("point of interest", "").replace("\n", " ").strip()
        
        # Alamat
        raw_lokasi = spec_table.get("lokasi", "")
        if raw_lokasi:
            data["alamat"] = raw_lokasi
        else:
            m_jl = re.search(r'(Jl\.?[^,]+(?:,[^,]+)?)', data["title"], re.I)
            if m_jl:
                data["alamat"] = m_jl.group(1).strip()
            else:
                data["alamat"] = data["title"]

        # Wilayah
        combined_text = f"{data['title']} {data['alamat']} {url}"
        data["regency"] = detect_regency(combined_text)

        # 3. Ekstraksi Koordinat
        # Cara A: Dari tabel spesifikasi
        raw_coord = spec_table.get("koordinat", "")
        if raw_coord:
            m_c = re.findall(r'(-[89]\.\d+)[,\s]+(11[45]\.\d+)', raw_coord)
            if m_c:
                data["latitude"] = float(m_c[0][0])
                data["longitude"] = float(m_c[0][1])
                
        # Cara B: Dari link static Mapbox
        if data["latitude"] is None or data["longitude"] is None:
            m_mapbox = re.findall(r'static/pin-[^\(]+\((115\.\d+),(-8\.\d+)\)', resp.text)
            if m_mapbox:
                data["longitude"] = float(m_mapbox[0][0])
                data["latitude"] = float(m_mapbox[0][1])

        # Cara C: Fallback untuk transit / bandara
        if (data["latitude"] is None or data["longitude"] is None) and slug in TRANSIT_FALLBACK_COORDS:
            fb = TRANSIT_FALLBACK_COORDS[slug]
            data["latitude"] = fb["lat"]
            data["longitude"] = fb["lng"]
            if not raw_lokasi:
                data["alamat"] = fb["alamat"]
            if data["regency"] == "Bali":
                data["regency"] = fb["kota"]

        # 4. Foto
        for img in soup.find_all("img"):
            src = img.get("src", "")
            if "uploads/attachments" in src or ("space" in src and "logo" not in src.lower()):
                data["foto"] = urljoin(BASE_URL, src)
                break

    except Exception as e:
        print(f"    [!] Error scraping {url}: {e}")
        
    return data

def main():
    parser = argparse.ArgumentParser(description="Scraper Titik Billboard & Media Luar Ruang Lestari Ads Bali (lestariads.com)")
    parser.add_argument("--output", type=str, default=OUTPUT_FILE, help=f"Nama file output Excel (default: {OUTPUT_FILE})")
    args = parser.parse_args()

    print("=" * 65)
    print("      SCRAPER TITIK IKLAN LUAR RUANG - LESTARIADS.COM (BALI)  ")
    print("=" * 65)
    
    start_time = time.time()
    session = requests.Session()
    session.headers.update(HEADERS)
    
    # 1. Kumpulkan daftar URL space Bali yang valid
    space_urls = get_bali_space_urls(session)
    if not space_urls:
        print("[!] Tidak ada space yang ditemukan.")
        return

    # 2. Scrape data dan koordinat tiap titik
    rows = []
    print(f"\n[*] Mengambil data dan koordinat dari {len(space_urls)} titik iklan...")
    
    for idx, url in enumerate(space_urls, 1):
        item = scrape_space_detail(url, session)
        
        lat = item["latitude"]
        lng = item["longitude"]
        title = item["title"]
        addr = item["alamat"]
        
        print(f"[{idx}/{len(space_urls)}] {title[:38]}... ", end="", flush=True)
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
            "Sisi": item["sisi"],
            "Penerangan": item["penerangan"],
            "Kabupaten / Kota": item["regency"],
            "Estimasi Traffic": item["traffic"],
            "Point of Interest": item["poi"],
            "Foto": item["foto"],
            "URL": item["url"]
        })
        
        time.sleep(0.2)
        
    df = pd.DataFrame(rows)
    
    # Validasi koordinat tidak kosong
    empty_coords = df[df["Latitude"].isna() | df["Longitude"].isna()]
    if not empty_coords.empty:
        print(f"\n[!] PERINGATAN: Terdapat {len(empty_coords)} data dengan koordinat kosong!")
    else:
        print(f"\n[+] Semua ({len(df)}) data memiliki koordinat Latitude & Longitude lengkap!")

    # Susun kolom
    cols_order = [
        "Nama Produk",
        "Alamat",
        "Jenis",
        "Longitude",
        "Latitude",
        "Brand",
        "Ukuran",
        "Orientasi",
        "Sisi",
        "Penerangan",
        "Kabupaten / Kota",
        "Estimasi Traffic",
        "Point of Interest",
        "Foto",
        "URL"
    ]
    df = df[cols_order]

    # Simpan ke Excel
    out_file = args.output
    print(f"\n[*] Menyimpan {len(df)} baris data ke '{out_file}'...")
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        sheet_name = "Data Lestari Ads Bali"
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
