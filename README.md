# Scraping POI Advertising Bali

Koleksi script web scraper Python untuk mengumpulkan data titik Point of Interest (POI) media periklanan luar ruang (Outdoor Advertising: Billboard, Baliho, Videotron, LED, Transit Ad, dll.) khusus di wilayah **Provinsi Bali**.

Project ini mengumpulkan data lengkap dengan kelengkapan koordinat geografis (**Latitude** & **Longitude** 100%) serta menyatukannya ke dalam satu file Excel master terstandarisasi.

---

## 🎯 Daftar Sumber & Script Scraper

| No | File Script | Domain / Vendor | Jenis Platform / Metode |
|:--:|:---|:---|:---|
| 1 | `baligen.py` | [thebaligen.com](https://thebaligen.com) | HTML Parsing & Metadata Extraction |
| 2 | `billboard_id.py` | [billboard.id](https://billboard.id) | Pagination & Listing Detail Parsing |
| 3 | `ivi.py` | [ivi.co.id](https://www.ivi.co.id) | Detail Page & Google Maps Query Extraction |
| 4 | `nusaka.py` | [nusakabali.com](https://nusakabali.com) | AJAX Action (`load_portfolio`) & DMS Coordinate Parsing |
| 5 | `lokasi_billboard.py` | [lokasibillboard.com](https://lokasibillboard.com) | Directory Scraping & Coordinate Geocoding |
| 6 | `epb.py` | [epb.co.id](https://www.epb.co.id) | Mapbox Static Map Pin Lat/Lng Extraction |
| 7 | `karyaku.co.id` | [karyaku.co.id](https://karyaku.co.id) | Elementor Cards & OpenStreetMap Geocoding Fallback |
| 8 | `petraco.py` | [petraco.id](https://petraco.id) | WordPress REST API (`/wp-json/wp/v2/product`) & Embed Parsing |
| 9 | `lestariads.py` | [lestariads.com](https://www.lestariads.com) | Sitemap Crawling, Listing Specs & Coordinate Extraction |
| 10 | `cvbalimedia.py` | [cvbalimedia.com](https://www.cvbalimedia.com) | Wix HTML Parsing & Google Maps Short URL Resolver |
| 11 | `merge_all.py` | **Master Consolidator** | Menggabungkan seluruh data ke dalam format Excel standar |

---

## 📋 Skema Kolom Master Data

Seluruh scraper menghasilkan atribut yang dinormalisasi oleh `merge_all.py` ke dalam skema standar:

* `No`
* `Nama Produk`
* `Alamat`
* `Jenis` (Billboard, Baliho, Videotron, LED Billboard, Transit Ad, Neon Box, dll.)
* `Latitude`
* `Longitude`
* `Brand / Vendor`
* `Ukuran`
* `Orientasi` (Horizontal / Vertical / 1 Muka / 2 Muka)
* `Kabupaten / Kota` (Badung, Denpasar, Gianyar, Bangli, Karangasem, Jembrana, Tabanan)
* `Kode / ID`
* `Sumber`
* `URL`

---

## 🚀 Instalasi & Cara Penggunaan

### 1. Prasyarat
Pastikan Python 3.9+ sudah terinstall.

### 2. Install Dependensi
```bash
pip install -r requirements.txt
```

### 3. Menjalankan Scraper Individual
Setiap scraper dapat dijalankan secara terpisah:
```bash
python baligen.py
python billboard_id.py
python ivi.py
python nusaka.py
python lokasi_billboard.py
python epb.py
python karyaku.py
python petraco.py
python lestariads.py
python cvbalimedia.py
```

### 4. Menggabungkan Semua Data Menjadi Satu Master Excel
Jalankan script konsolidator:
```bash
python merge_all.py
```
Script akan menghasilkan file `MASTER_POI_ADS_BALI.xlsx` lengkap dengan lembar ringkasan dan format styling profesional.
