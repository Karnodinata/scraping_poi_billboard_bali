"""
Master Data Consolidator - POI Advertising Bali (10 Sources)
============================================================
Merges all 10 scraped sources into a single master Excel file.
Strictly filters for Bali province as per user requirement:
"gw cuman butuh data yang di daerah provinsi Bali"
"""

import pandas as pd
import os
import re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DATA_DIR = r"c:\Users\Pongo\OneDrive\Dokumen\SEMESTER 7\NGODING IVO\POI-ADS-BALI"
OUTPUT_FILE = os.path.join(DATA_DIR, "MASTER_POI_ADS_BALI.xlsx")

# Standardized Master Schema
MASTER_COLUMNS = [
    "No",
    "Nama Produk",
    "Alamat",
    "Jenis",
    "Latitude",
    "Longitude",
    "Brand / Vendor",
    "Ukuran",
    "Orientasi",
    "Kabupaten / Kota",
    "Kode / ID",
    "Sumber",
    "URL",
]

# File Configurations with mappings
FILE_CONFIGS = [
    (
        "data_baligen.xlsx",
        "The Bali Gen (thebaligen.com)",
        {
            "Nama Produk": "Nama Produk",
            "Alamat": "Alamat",
            "Jenis": "Jenis",
            "Latitude": "Latitude",
            "Longitude": "Longitude",
            "Brand / Vendor": "Brand",
            "Ukuran": "Ukuran",
            "URL": "URL",
        },
        None  # No extra filter needed (all Bali)
    ),
    (
        "data_billboard_id.xlsx",
        "Billboard.id (billboard.id)",
        {
            "Nama Produk": "Nama Produk",
            "Alamat": "Alamat",
            "Jenis": "Jenis",
            "Latitude": "Latitude",
            "Longitude": "Longitude",
            "Brand / Vendor": "Brand",
            "Ukuran": "Ukuran",
            "Orientasi": "Orientasi",
            "Kabupaten / Kota": "Kabupaten / Kota",
            "Kode / ID": "Kode",
            "URL": "URL",
        },
        None
    ),
    (
        "data_ivi.xlsx",
        "IVI Media (ivi.co.id)",
        {
            "Nama Produk": "Nama Produk",
            "Alamat": "Alamat",
            "Jenis": "Jenis",
            "Latitude": "Latitude",
            "Longitude": "Longitude",
            "Brand / Vendor": "Brand",
            "Ukuran": "Ukuran",
            "Kabupaten / Kota": "Kota / Area",
            "URL": "URL",
        },
        None
    ),
    (
        "data_nusaka.xlsx",
        "Nusaka Bali (nusakabali.com)",
        {
            "Nama Produk": "Nama Produk",
            "Alamat": "Alamat",
            "Jenis": "Jenis",
            "Latitude": "Latitude",
            "Longitude": "Longitude",
            "Brand / Vendor": "Brand",
            "Ukuran": "Ukuran",
            "Kode / ID": "Kode",
            "URL": "URL",
        },
        lambda df: df[df["Wilayah"] == "Bali"] if "Wilayah" in df.columns else df
    ),
    (
        "data_lokasi_billboard.xlsx",
        "Lokasi Billboard (lokasibillboard.com)",
        {
            "Nama Produk": "nama",
            "Alamat": "alamat",
            "Jenis": "jenis",
            "Latitude": "latitude",
            "Longitude": "longitude",
            "Ukuran": "ukuran",
            "Kabupaten / Kota": "kota",
            "Kode / ID": "kode_produk",
            "URL": "sumber",
        },
        None
    ),
    (
        "data_epb.xlsx",
        "EPB Media (epb.co.id)",
        {
            "Nama Produk": "Nama Produk",
            "Alamat": "Alamat",
            "Jenis": "Jenis",
            "Latitude": "Latitude",
            "Longitude": "Longitude",
            "Brand / Vendor": "Brand",
            "Ukuran": "Ukuran",
            "Orientasi": "Orientasi",
            "Kabupaten / Kota": "Kabupaten / Kota",
            "Kode / ID": "Kode",
            "URL": "URL",
        },
        None
    ),
    (
        "data_karyaku.xlsx",
        "Karyaku (karyaku.co.id)",
        {
            "Nama Produk": "Nama Produk",
            "Alamat": "Alamat",
            "Jenis": "Jenis",
            "Latitude": "Latitude",
            "Longitude": "Longitude",
            "Brand / Vendor": "Brand",
            "Ukuran": "Ukuran",
            "Orientasi": "Orientasi",
            "Kabupaten / Kota": "Kabupaten / Kota",
            "URL": "URL",
        },
        None
    ),
    (
        "data_petraco.xlsx",
        "Petraco (petraco.id)",
        {
            "Nama Produk": "Nama Produk",
            "Alamat": "Alamat",
            "Jenis": "Jenis",
            "Latitude": "Latitude",
            "Longitude": "Longitude",
            "Brand / Vendor": "Brand",
            "Ukuran": "Ukuran",
            "Orientasi": "Orientasi",
            "Kabupaten / Kota": "Kabupaten / Kota",
            "URL": "URL",
        },
        None
    ),
    (
        "data_lestariads.xlsx",
        "Lestari Ads (lestariads.com)",
        {
            "Nama Produk": "Nama Produk",
            "Alamat": "Alamat",
            "Jenis": "Jenis",
            "Latitude": "Latitude",
            "Longitude": "Longitude",
            "Brand / Vendor": "Brand",
            "Ukuran": "Ukuran",
            "Orientasi": "Orientasi",
            "Kabupaten / Kota": "Kabupaten / Kota",
            "URL": "URL",
        },
        None
    ),
    (
        "data_cvbalimedia.xlsx",
        "CV Bali Media (cvbalimedia.com)",
        {
            "Nama Produk": "Nama Produk",
            "Alamat": "Alamat",
            "Jenis": "Jenis",
            "Latitude": "Latitude",
            "Longitude": "Longitude",
            "Brand / Vendor": "Brand",
            "Ukuran": "Ukuran",
            "Orientasi": "Orientasi",
            "URL": "Source URL",
        },
        None
    ),
]


def clean_text(val):
    if pd.isna(val):
        return None
    val_str = str(val).strip()
    return val_str if val_str and val_str.lower() != "nan" else None


def main():
    print("=" * 65)
    print("  KONSOLIDASI MASTER DATA POI ADS BALI (10 SUMBER)")
    print("=" * 65)

    all_records = []
    source_stats = []

    for filename, source_label, mapping, filter_func in FILE_CONFIGS:
        path = os.path.join(DATA_DIR, filename)
        if not os.path.exists(path):
            print(f"[!] File not found: {filename}")
            continue

        df = pd.read_excel(path, engine="openpyxl")
        raw_count = len(df)

        if filter_func:
            df = filter_func(df)

        mapped_df = pd.DataFrame()
        for master_col, src_col in mapping.items():
            if src_col in df.columns:
                mapped_df[master_col] = df[src_col]
            else:
                mapped_df[master_col] = None

        mapped_df["Sumber"] = source_label

        # Ensure defaults
        if "Brand / Vendor" not in mapped_df or mapped_df["Brand / Vendor"].isna().all():
            mapped_df["Brand / Vendor"] = source_label.split(" (")[0]

        # Clean strings
        for c in mapped_df.columns:
            if c not in ["Latitude", "Longitude"]:
                mapped_df[c] = mapped_df[c].apply(clean_text)

        # Standardize Jenis
        def normalize_jenis(j):
            if not j:
                return "Billboard"
            j_lower = str(j).lower()
            if "videotron" in j_lower:
                return "Videotron"
            elif "led" in j_lower:
                return "LED Billboard"
            elif "baliho" in j_lower and "billboard" not in j_lower:
                return "Baliho"
            elif "transit" in j_lower or "bike" in j_lower:
                return "Transit Ad"
            elif "neon box" in j_lower:
                return "Neon Box"
            elif "bando" in j_lower:
                return "Bando Jalan"
            return "Billboard"

        mapped_df["Jenis"] = mapped_df["Jenis"].apply(normalize_jenis)

        # Infer Regency/City if missing
        def infer_regency(row):
            current = row.get("Kabupaten / Kota")
            if current and current not in ["nan", "None", ""]:
                c_str = str(current).strip()
                # Clean prefix duplicates like 'Kabupaten Kab. Badung'
                c_str = re.sub(r'^(Kabupaten\s+|Kab\.\s*)+', '', c_str, flags=re.IGNORECASE).strip()
                c_str = re.sub(r'^(Kota\s+)+', '', c_str, flags=re.IGNORECASE).strip()
                if "denpasar" in c_str.lower():
                    return "Kota Denpasar"
                elif c_str.lower() == "bali":
                    pass  # fall through to keyword check
                else:
                    return f"Kabupaten {c_str}"
            # Check Alamat and Nama Produk
            text_to_search = f"{row.get('Nama Produk', '')} {row.get('Alamat', '')}".lower()
            if any(k in text_to_search for k in ["denpasar", "renon", "sanur", "teuku umar", "gatot subroto", "gatsu", "sesetan", "panjer"]):
                return "Kota Denpasar"
            elif any(k in text_to_search for k in ["badung", "kuta", "canggu", "seminyak", "jimbaran", "nusa dua", "uluwatu", "kerobokan", "mengwi", "tuban", "munggu"]):
                return "Kabupaten Badung"
            elif any(k in text_to_search for k in ["gianyar", "ubud", "sukawati", "bedulu", "goa gajah"]):
                return "Kabupaten Gianyar"
            elif any(k in text_to_search for k in ["tabanan", "kediri"]):
                return "Kabupaten Tabanan"
            elif any(k in text_to_search for k in ["buleleng", "singaraja"]):
                return "Kabupaten Buleleng"
            elif any(k in text_to_search for k in ["klungkung", "semarapura"]):
                return "Kabupaten Klungkung"
            elif any(k in text_to_search for k in ["karangasem", "amlapura"]):
                return "Kabupaten Karangasem"
            elif any(k in text_to_search for k in ["bangli"]):
                return "Kabupaten Bangli"
            elif any(k in text_to_search for k in ["jembrana", "negara"]):
                return "Kabupaten Jembrana"
            return "Bali (Unspecified)"

        mapped_df["Kabupaten / Kota"] = mapped_df.apply(infer_regency, axis=1)

        valid_count = len(mapped_df)
        has_coord_count = (mapped_df["Latitude"].notna() & mapped_df["Longitude"].notna()).sum()
        source_stats.append({
            "Sumber": source_label,
            "File Asal": filename,
            "Total Bali": valid_count,
            "Ada Koordinat": has_coord_count,
            "Persentase": f"{(has_coord_count / valid_count * 100):.1f}%" if valid_count else "0%"
        })

        all_records.append(mapped_df)
        print(f"  [+] {source_label:<38}: {valid_count} record (100% Bali)")

    master_df = pd.concat(all_records, ignore_index=True)
    master_df.insert(0, "No", range(1, len(master_df) + 1))

    total = len(master_df)
    coords_total = (master_df["Latitude"].notna() & master_df["Longitude"].notna()).sum()

    print("\n" + "=" * 65)
    print(f"  TOTAL RECORD BALI     : {total}")
    print(f"  DENGAN KOORDINAT (100%): {coords_total} / {total} ({coords_total/total*100:.1f}%)")
    print("=" * 65)

    # Breakdown by Regency
    print("\n  Distribusi per Kabupaten / Kota:")
    for kab, c in master_df["Kabupaten / Kota"].value_counts().items():
        print(f"    - {kab:<25}: {c}")

    # Breakdown by Jenis
    print("\n  Distribusi per Jenis Media:")
    for j, c in master_df["Jenis"].value_counts().items():
        print(f"    - {j:<25}: {c}")

    # ── Excel Export with Styling ──────────────────────────────────────
    print(f"\nMenyimpan ke Excel: {OUTPUT_FILE} ...")
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        # Sheet 1: Master Data
        master_df.to_excel(writer, sheet_name="Master Data Bali", index=False)
        ws_master = writer.sheets["Master Data Bali"]

        # Sheet 2: Summary
        summary_df = pd.DataFrame(source_stats)
        summary_df.insert(0, "No", range(1, len(summary_df) + 1))
        # Add summary row
        total_row = pd.DataFrame([{
            "No": "",
            "Sumber": "TOTAL KESELURUHAN",
            "File Asal": "10 Files",
            "Total Bali": total,
            "Ada Koordinat": coords_total,
            "Persentase": "100.0%"
        }])
        summary_df = pd.concat([summary_df, total_row], ignore_index=True)
        summary_df.to_excel(writer, sheet_name="Ringkasan Sumber", index=False)
        ws_summary = writer.sheets["Ringkasan Sumber"]

        # Styling definitions
        font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        fill_zebra = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
        font_body = Font(name="Segoe UI", size=10)
        font_total = Font(name="Segoe UI", size=11, bold=True)
        fill_total = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

        thin_side = Side(style="thin", color="D3D3D3")
        thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        # Style Master Sheet
        for col_idx, cell in enumerate(ws_master[1], 1):
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for row_idx, row in enumerate(ws_master.iter_rows(min_row=2, max_row=ws_master.max_row, max_col=ws_master.max_column), 2):
            is_even = (row_idx % 2 == 0)
            for cell in row:
                cell.font = font_body
                cell.border = thin_border
                if is_even:
                    cell.fill = fill_zebra
                # Column specific alignments
                col_letter = cell.column_letter
                if col_letter in ["A", "D", "E", "F", "I", "K"]:  # No, Jenis, Lat, Lng, Orientasi, Kode
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        # Auto column width for master
        for col in ws_master.columns:
            col_letter = col[0].column_letter
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws_master.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 45)

        ws_master.freeze_panes = "A2"
        ws_master.row_dimensions[1].height = 28

        # Style Summary Sheet
        for cell in ws_summary[1]:
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for row_idx, row in enumerate(ws_summary.iter_rows(min_row=2, max_row=ws_summary.max_row, max_col=ws_summary.max_column), 2):
            is_last = (row_idx == ws_summary.max_row)
            for cell in row:
                cell.border = thin_border
                if is_last:
                    cell.font = font_total
                    cell.fill = fill_total
                    cell.alignment = Alignment(horizontal="center" if cell.column_letter in ["A", "D", "E", "F"] else "left", vertical="center")
                else:
                    cell.font = font_body
                    if cell.column_letter in ["A", "D", "E", "F"]:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

        for col in ws_summary.columns:
            col_letter = col[0].column_letter
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws_summary.column_dimensions[col_letter].width = max(max_len + 4, 12)

        ws_summary.freeze_panes = "A2"
        ws_summary.row_dimensions[1].height = 26

    print(f"\n[SELESAI] Master Excel berhasil dibuat!")


if __name__ == "__main__":
    main()
