import argparse
import re
import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE_URL = "https://nusakabali.com"
MEDIA_URL = f"{BASE_URL}/media/"
AJAX_URL = f"{BASE_URL}/wp-admin/admin-ajax.php"
OUTPUT_FILE = "data_nusaka.xlsx"
BRAND_NAME = "Nusaka Bali"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

def get_listing_items(bali_only=False):
    """
    Mengambil daftar titik iklan luar ruang dari halaman https://nusakabali.com/media/
    """
    print(f"[*] Mengakses {MEDIA_URL} ...")
    resp = requests.get(MEDIA_URL, headers=HEADERS, timeout=15)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")
    cards = soup.select(".gallery-portfolio")
    print(f"[+] Ditemukan {len(cards)} item pada halaman.")

    items = []
    for card in cards:
        filter_tag = card.get("data-filter", "")
        area_name = filter_tag.replace("portfolio-", "").replace("-", " ").title()

        if bali_only and "bali" not in filter_tag.lower():
            continue

        link = card.find("a")
        data_id = link.get("data-id") if link else None
        title_el = card.find("h3")
        title = title_el.get_text(strip=True) if title_el else ""

        if data_id:
            items.append({
                "id": data_id,
                "title": title,
                "area": area_name,
                "filter": filter_tag
            })

    print(f"[+] Item yang akan diproses: {len(items)}")
    return items

def parse_dms_to_dd(text):
    """Mengubah format koordinat DMS (Derajat Menit Detik) ke Decimal Degrees."""
    cleaned = text.replace("&#8217;", "'").replace("&#8243;", '"').replace("\u2019", "'").replace("\u2033", '"')
    parts = re.findall(r'(\d+)[\D]+(\d+)[\D]+([\d.]+)[\D]*([NSEWnsew])', cleaned)
    lat, lng = None, None
    for p in parts:
        d, m, s, direction = float(p[0]), float(p[1]), float(p[2]), p[3].upper()
        dd = d + (m / 60.0) + (s / 3600.0)
        if direction in ['S', 'W']:
            dd = -dd
        dd = round(dd, 6)
        if direction in ['S', 'N']:
            lat = dd
        elif direction in ['E', 'W']:
            lng = dd
    return lat, lng

def scrape_item_detail(item):
    """
    Mengambil detail titik iklan melalui endpoint AJAX load_portfolio.
    """
    params = {
        "action": "load_portfolio",
        "value": item["id"]
    }
    
    max_retries = 3
    resp = None
    for attempt in range(max_retries):
        try:
            resp = requests.get(AJAX_URL, params=params, headers=HEADERS, timeout=15)
            resp.raise_for_status()
            break
        except Exception as e:
            if attempt == max_retries - 1:
                print(f"[!] Gagal mengambil ID {item['id']}: {e}")
                return None
            time.sleep(1)

    soup = BeautifulSoup(resp.text, "html.parser")

    # Nama / Judul Titik
    modal_title = soup.select_one(".modal-title")
    title = modal_title.get_text(strip=True) if modal_title else item["title"]

    # Tabel Spesifikasi & Informasi
    info = {}
    for tr in soup.select("table tr"):
        tds = tr.find_all("td")
        if len(tds) >= 3:
            k = tds[0].get_text(strip=True).lower()
            v = tds[2].get_text(" ", strip=True)
            info[k] = v
        elif len(tds) == 2:
            k = tds[0].get_text(strip=True).lower()
            v = tds[1].get_text(" ", strip=True)
            info[k] = v

    # 1. Koordinat dari embed Google Maps iframe (paling presisi)
    lat, lng = None, None
    embed_ifr = soup.select_one(".wrapper-map iframe")
    if embed_ifr:
        src = embed_ifr.get("src", "")
        m = re.search(r"!2d([-\d.]+)!3d([-\d.]+)", src)
        if m:
            lng = float(m.group(1))
            lat = float(m.group(2))

    # 2. Alternatif: dari streetview iframe
    if lat is None or lng is None:
        sv_ifr = soup.select_one(".wrapper-streetview iframe")
        if sv_ifr:
            src = sv_ifr.get("src", "")
            m_sv = re.search(r"!2d([-\d.]+).*!1d([-\d.]+)", src)
            if m_sv:
                lng = float(m_sv.group(1))
                lat = float(m_sv.group(2))

    # 3. Alternatif: dari teks koordinat DMS pada tabel
    if lat is None or lng is None:
        raw_coord = info.get("koordinat", "")
        if raw_coord:
            dms_lat, dms_lng = parse_dms_to_dd(raw_coord)
            if lat is None and dms_lat is not None:
                lat = dms_lat
            if lng is None and dms_lng is not None:
                lng = dms_lng

    # 4. Alternatif: dari link Location
    if lat is None or lng is None:
        loc_a = soup.select_one("a.location")
        if loc_a and "href" in loc_a.attrs:
            m_ll = re.search(r"[?&]ll=([-\d.]+),([-\d.]+)", loc_a["href"])
            if m_ll:
                lat = float(m_ll.group(1))
                lng = float(m_ll.group(2))

    # Klasifikasi Jenis Media
    spesifikasi = info.get("spesifikasi", "")
    spesifikasi_lower = (spesifikasi + " " + title).lower()

    if any(k in spesifikasi_lower for k in ["videotron", "led"]):
        jenis = "Videotron"
    elif any(k in spesifikasi_lower for k in ["baliho"]):
        jenis = "Baliho"
    elif any(k in spesifikasi_lower for k in ["billboard", "bpd bali", "frontlite", "frontlight", "backlite"]):
        jenis = "Billboard"
    else:
        jenis = "Billboard"

    # Ekstraksi Ukuran
    ukuran = ""
    m_uk = re.search(r"(\d+[\s]*[Xx*×\u00d7\ufffd][\s]*\d+[\s]*m?)", spesifikasi)
    if m_uk:
        ukuran = m_uk.group(1).replace("\ufffd", "x").replace("×", "x").strip()

    # Alamat
    alamat = title
    if "bali" not in alamat.lower() and item["area"] == "Bali":
        alamat += ", Bali"

    return {
        "Nama Produk": title,
        "Alamat": alamat,
        "Jenis": jenis,
        "Longitude": round(lng, 6) if lng is not None else None,
        "Latitude": round(lat, 6) if lat is not None else None,
        "Brand": BRAND_NAME,
        "Ukuran": ukuran,
        "Kode": info.get("code", ""),
        "Wilayah": item["area"],
        "URL": MEDIA_URL
    }

def main():
    parser = argparse.ArgumentParser(description="Scraper Media Iklan Nusaka Bali (https://nusakabali.com/media/)")
    parser.add_argument(
        "--bali-only",
        action="store_true",
        help="Hanya scrape titik di wilayah Bali (abaikan Kupang dan Labuan Bajo)"
    )
    parser.add_argument(
        "--output",
        type=str,
        default=OUTPUT_FILE,
        help=f"Nama file output Excel (default: {OUTPUT_FILE})"
    )
    args = parser.parse_args()

    print("=" * 60)
    print("        SCRAPER IKLAN LUAR RUANG - NUSAKA BALI        ")
    print(f"        Target: {MEDIA_URL}                          ")
    print("=" * 60)

    start_time = time.time()
    items = get_listing_items(bali_only=args.bali_only)

    if not items:
        print("[!] Tidak ada titik iklan yang ditemukan untuk di-scrape.")
        return

    print(f"\n[*] Mengambil detail untuk {len(items)} titik iklan menggunakan multithreading...")
    scraped_data = []

    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_item = {executor.submit(scrape_item_detail, item): item for item in items}
        completed = 0
        for future in as_completed(future_to_item):
            res = future.result()
            if res:
                scraped_data.append(res)
            completed += 1
            print(f"\r Progress: {completed}/{len(items)} selesai...", end="", flush=True)

    print(f"\n[+] Sukses mengambil data untuk {len(scraped_data)} titik iklan.")

    # Simpan ke DataFrame
    df = pd.DataFrame(scraped_data)

    # Urutan kolom
    columns_order = [
        "Nama Produk",
        "Alamat",
        "Jenis",
        "Longitude",
        "Latitude",
        "Brand",
        "Ukuran",
        "Kode",
        "Wilayah",
        "URL"
    ]
    columns_order = [c for c in columns_order if c in df.columns]
    df = df[columns_order]

    # Simpan ke file Excel
    out_file = args.output
    print(f"[*] Menyimpan data ke file Excel: {out_file} ...")
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data Nusaka Bali")
        
        # Auto-adjust column width
        worksheet = writer.sheets["Data Nusaka Bali"]
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
    print(f"- Wilayah: {df['Wilayah'].value_counts().to_dict()}")

if __name__ == "__main__":
    main()
