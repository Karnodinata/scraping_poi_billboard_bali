import argparse
import re
import time
import requests
import pandas as pd

BASE_URL = "https://billboard.id"
API_URL = f"{BASE_URL}/api/spots"
OUTPUT_FILE = "data_billboard_id.xlsx"
BRAND_NAME = "Billboard.id"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Referer": f"{BASE_URL}/provinsi/bali/"
}

def fetch_spots(province="Bali"):
    """
    Mengambil seluruh data titik iklan dari API backend billboard.id
    berdasarkan provinsi yang dipilih.
    """
    print(f"[*] Mengakses API billboard.id untuk provinsi: {province} ...")
    
    all_spots = []
    page = 1
    limit = 50

    while True:
        params = {
            "province": province,
            "limit": limit,
            "page": page
        }
        
        try:
            resp = requests.get(API_URL, params=params, headers=HEADERS, timeout=15)
            resp.raise_for_status()
            data = resp.json()
        except Exception as e:
            print(f"[!] Gagal mengambil data halaman {page}: {e}")
            break

        spots = data.get("spots", [])
        total = data.get("total", 0)
        total_pages = data.get("totalPages", 1)

        print(f"[+] Halaman {page}/{total_pages}: menerima {len(spots)} titik iklan (Total keseluruhan: {total}).")
        all_spots.extend(spots)

        if page >= total_pages or not spots:
            break

        page += 1

    print(f"[+] Berhasil mengumpulkan {len(all_spots)} titik iklan untuk provinsi {province}.")
    return all_spots

def process_spots(spots):
    """
    Memformat dan membersihkan data hasil ekstraksi untuk disimpan ke Excel.
    """
    rows = []
    for s in spots:
        title = (s.get("title") or "").strip()
        location = (s.get("location") or "").strip()
        area = (s.get("area") or "").strip()
        
        # Alamat gabungan lengkap
        alamat_parts = [p for p in [title, location, area] if p]
        alamat = ", ".join(alamat_parts)

        # URL detail spot
        slug = (s.get("slug") or "").strip()
        spot_url = f"{BASE_URL}/spot/{slug}/" if slug else f"{BASE_URL}/provinsi/bali/"

        # Harga
        price_raw = s.get("selling_price")
        price = None
        if price_raw:
            try:
                clean_p = re.sub(r"[^\d]", "", str(price_raw).split(",")[0])
                if clean_p:
                    price = int(clean_p)
            except Exception:
                price = str(price_raw)

        # Koordinat
        lat = s.get("lat")
        lng = s.get("lng")
        if lat is not None:
            try:
                lat = float(lat)
            except ValueError:
                pass
        if lng is not None:
            try:
                lng = float(lng)
            except ValueError:
                pass

        rows.append({
            "Nama Produk": title,
            "Alamat": alamat,
            "Jenis": s.get("media_type") or "Billboard",
            "Longitude": lng,
            "Latitude": lat,
            "Brand": BRAND_NAME,
            "Ukuran": s.get("media_size") or "",
            "Orientasi": s.get("orientation") or "",
            "Kode": s.get("spot_id") or "",
            "Kabupaten / Kota": location,
            "Harga Sewa (IDR)": price,
            "Status": s.get("spot_status") or "",
            "URL": spot_url
        })

    return rows

def main():
    parser = argparse.ArgumentParser(description="Scraper Titik Billboard & Media Luar Ruang dari billboard.id")
    parser.add_argument(
        "--province",
        type=str,
        default="Bali",
        help="Nama provinsi yang ingin di-scrape (default: Bali)"
    )
    parser.add_argument(
        "--output",
        type=str,
        default=OUTPUT_FILE,
        help=f"Nama file output Excel (default: {OUTPUT_FILE})"
    )
    args = parser.parse_args()

    print("=" * 60)
    print("      SCRAPER TITIK IKLAN LUAR RUANG - BILLBOARD.ID     ")
    print(f"      Provinsi: {args.province}                         ")
    print("=" * 60)

    start_time = time.time()
    raw_spots = fetch_spots(province=args.province)

    if not raw_spots:
        print("[!] Tidak ada data yang ditemukan.")
        return

    formatted_data = process_spots(raw_spots)
    df = pd.DataFrame(formatted_data)

    # Urutan kolom utama
    columns_order = [
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
        "Harga Sewa (IDR)",
        "Status",
        "URL"
    ]
    columns_order = [c for c in columns_order if c in df.columns]
    df = df[columns_order]

    # Simpan ke Excel dengan openpyxl
    out_file = args.output
    print(f"\n[*] Menyimpan data ke file Excel: {out_file} ...")
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        sheet_name = f"Data {args.province}"[:31]
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        # Auto-adjust column width
        worksheet = writer.sheets[sheet_name]
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
    print(f"- Lokasi: {df['Kabupaten / Kota'].value_counts().to_dict()}")

if __name__ == "__main__":
    main()
