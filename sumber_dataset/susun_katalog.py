#!/usr/bin/env python3
"""Susun katalog sumber web untuk setiap file Excel/CSV di folder Dataset repo Inflasi-Aceh.

Keluaran: katalog_sumber.json (dipakai oleh buat_dokumen.py dan ambil_screenshot.py).

Cara pakai:
    python susun_katalog.py /path/ke/Inflasi-Aceh/Dataset

ID tabel dinamis BPS adalah base64 dari "<nomor_var>#2". Status verifikasi:
  - "terverifikasi": nomor tabel dan judulnya muncul berpasangan di indeks mesin pencari
    (halaman bps.go.id / aceh.bps.go.id dengan judul yang sama persis).
  - "pola": nomor tabel disimpulkan dari penomoran berurutan yang sudah terverifikasi
    di beberapa titik dalam seri yang sama; wajib dicek saat membuka halaman.
"""
from __future__ import annotations

import base64
import json
import re
import sys
from pathlib import Path

BPS = "https://www.bps.go.id/id/statistics-table/2/{id}/{slug}.html"
QB_ACEH = "https://aceh.bps.go.id/id/query-builder"
TPT_ACEH = ("https://aceh.bps.go.id/id/statistics-table/3/"
            "V2pOVWJWcHJURGg0U2pONFJYaExhVXB0TUhacVFUMDkjMw==/"
            "tingkat-pengangguran-terbuka--tpt--dan-tingkat-partisipasi-angkatan-kerja--tpak--"
            "menurut-kabupaten-kota-di-provinsi-aceh.html?year={tahun}")

KELOMPOK_IHK = {
    1: "01 Makanan, Minuman dan Tembakau",
    2: "02 Pakaian dan Alas Kaki",
    3: "03 Perumahan, Air, Listrik dan Bahan Bakar Rumah Tangga",
    4: "04 Perlengkapan, Peralatan dan Pemeliharaan Rutin Rumah Tangga",
    5: "05 Kesehatan",
    6: "06 Transportasi",
    7: "07 Informasi, Komunikasi dan Jasa Keuangan",
    8: "08 Rekreasi, Olahraga dan Budaya",
    9: "09 Pendidikan",
    10: "10 Penyediaan Makanan dan Minuman / Restoran",
    11: "11 Perawatan Pribadi dan Jasa Lainnya",
}
# Titik yang sudah terverifikasi lewat mesin pencari (nomor kelompok).
VERIF_IHK = {1, 2, 3, 4, 6, 10, 11}     # var = 2211 + k
VERIF_YTD = {3, 5}                     # var = 2233 + k (M-to-M 02 = 2224 dan 03 = 2225 ikut terverifikasi)
VERIF_YOY = {1, 3, 4, 5, 6, 7, 9, 10, 11}  # var = 2249 + k


def b64(var: int) -> str:
    return base64.b64encode(f"{var}#2".encode()).decode()


def slug(judul: str) -> str:
    s = judul.lower().replace("/", "-")
    s = re.sub(r"[^a-z0-9]+", "-", s)
    return s.strip("-")


def bps(var: int, judul: str, status: str, catatan: str = "") -> dict:
    return {
        "penerbit": "Badan Pusat Statistik (BPS) RI",
        "situs": "www.bps.go.id",
        "jenis": "Tabel Statistik (tabel dinamis)",
        "var_id": var,
        "kode_url": b64(var),
        "url": BPS.format(id=b64(var), slug=slug(judul)),
        # Menurut pengguna, file diunduh lewat fitur Tabel Dinamis (query builder) situs BPS Provinsi Aceh.
        "url_tabel_dinamis": QB_ACEH,
        "status": status,
        "catatan": catatan,
    }


def daftar_tabel() -> list[dict]:
    t: list[dict] = []

    # 1. IHK dan inflasi 150 kabupaten/kota (2022=100)
    t.append({"kode": "IHK150", "judul": "Indeks Harga Konsumen 150 Kabupaten/Kota (2022=100)",
              "pola_file": r"^Indeks Harga Konsumen 150 Kabupaten_Kota",
              **bps(2245, "Indeks Harga Konsumen 150 Kabupaten/Kota (2022=100)", "terverifikasi")})
    for k, nama in KELOMPOK_IHK.items():
        j = f"Indeks Harga Konsumen (2022=100) Menurut Kelompok dan Sub Kelompok {nama}"
        t.append({"kode": f"IHK-K{k:02d}", "judul": j,
                  "pola_file": rf"^Indeks Harga Konsumen \(2022=100\) Menurut Kelompok dan Sub Kelompok {k:02d} ",
                  **bps(2211 + k, j, "terverifikasi" if k in VERIF_IHK else "pola")})
    for k, nama in KELOMPOK_IHK.items():
        j = f"Inflasi Tahun Kalender (Y-to-D) (2022=100) Menurut Kelompok dan Sub Kelompok {nama}"
        t.append({"kode": f"YTD-K{k:02d}", "judul": j,
                  "pola_file": rf"^Inflasi Tahun Kalender \(Y-to-D\) \(2022=100\) Menurut Kelompok dan Sub Kelompok {k:02d} ",
                  **bps(2233 + k, j, "terverifikasi" if k in VERIF_YTD else "pola")})
    for k, nama in KELOMPOK_IHK.items():
        j = f"Inflasi Tahunan (Y-on-Y) (2022=100) Menurut Kelompok dan Sub Kelompok {nama}"
        t.append({"kode": f"YOY-K{k:02d}", "judul": j,
                  "pola_file": rf"^Inflasi Tahunan \(Y-on-Y\) \(2022=100\) Menurut Kelompok dan Sub Kelompok {k:02d} ",
                  **bps(2249 + k, j, "terverifikasi" if k in VERIF_YOY else "pola")})

    # 2. Ketahanan pangan
    j = ("Prevalensi Penduduk Dengan Kerawanan Pangan Sedang Atau Berat, "
         "Berdasarkan Pada Skala Pengalaman Kerawanan Pangan")
    t.append({"kode": "FIES", "judul": j, "pola_file": r"^Prevalensi Penduduk Dengan Kerawanan Pangan",
              **bps(1474, j, "terverifikasi", "Tingkat provinsi (38 provinsi); baris Aceh ada di dalamnya.")})
    j = "Prevalensi Ketidakcukupan Konsumsi Pangan (Persen) Per Kabupaten/kota"
    t.append({"kode": "POU", "judul": j, "pola_file": r"^Prevalensi Ketidakcukupan Konsumsi Pangan",
              **bps(2269, j, "terverifikasi")})

    # 3. Konsumsi per kapita seminggu (Susenas), satuan komoditas
    konsumsi = [  # (kode, kelompok di judul, var, status)
        ("KON-PADI", "Padi-Padian", 2094, "terverifikasi"),
        ("KON-UMBI", "Umbi-Umbian", 2095, "terverifikasi"),
        ("KON-IKAN", "Ikan", 2096, "terverifikasi"),
        ("KON-DAGING", "Daging", 2097, "terverifikasi"),
        ("KON-TELUR", "Telur dan Susu", 2099, "terverifikasi"),
        ("KON-SAYUR", "Sayur-Sayuran", 2100, "terverifikasi"),
        ("KON-KACANG", "Kacang-Kacangan", 2101, "terverifikasi"),
        ("KON-BUAH", "Buah-Buahan", 2102, "terverifikasi"),
        ("KON-MINYAK", "Minyak dan Kelapa", 2103, "terverifikasi"),
        ("KON-MINUMAN", "Bahan Minuman", 2104, "terverifikasi"),
        ("KON-LAINNYA", "Bahan Makanan Lainnya", 2106, "terverifikasi"),
        ("KON-JADI", "Makanan dan Minuman Jadi", 2107, "terverifikasi"),
    ]
    for kode, kel, var, st in konsumsi:
        j = f"Rata-rata Konsumsi Perkapita Seminggu Menurut Kelompok {kel} Per Kabupaten/kota"
        t.append({"kode": kode, "judul": j,
                  "pola_file": rf"^Rata-rata Konsumsi Perkapita Seminggu Menurut Kelompok {re.escape(kel)} Per Kabupaten_kota",
                  **bps(var, j, st)})
    konsumsi25 = [
        ("KON25-IKAN", "Ikan", 2550, "terverifikasi"),
        ("KON25-DAGING", "Daging", 2552, "terverifikasi"),
        ("KON25-TELUR", "Telur dan Susu", 2554, "terverifikasi"),
        ("KON25-SAYUR", "Sayur-Sayuran", 2556, "terverifikasi"),
        ("KON25-BUAH", "Buah-Buahan", 2558, "pola"),
    ]
    for kode, kel, var, st in konsumsi25:
        j = f"Rata-rata Konsumsi Perkapita Seminggu Menurut Kelompok {kel} Per Kabupaten/Kota-Komoditas 2025"
        t.append({"kode": kode, "judul": j,
                  "pola_file": rf"^Rata-rata Konsumsi Perkapita Seminggu Menurut Kelompok {re.escape(kel)} Per Kabupaten_Kota-Komoditas 2025",
                  **bps(var, j, st, "Rincian komoditas versi 2025 (daftar komoditas Susenas 2025)."
                        if st == "terverifikasi" else
                        "Nomor ditarik dari pola pasangan konsumsi genap/pengeluaran ganjil (2550, 2552, 2554, 2556); "
                        "kepastian paling rendah, cek judul saat halaman dibuka.")})

    # 4. Pengeluaran per kapita seminggu (Susenas), rupiah/kapita/minggu
    pengeluaran = [
        ("PEN-PADI", "Padi-Padian", 2109, "terverifikasi"),
        ("PEN-UMBI", "Umbi-Umbian", 2110, "terverifikasi"),
        ("PEN-IKAN", "Ikan", 2111, "terverifikasi"),
        ("PEN-DAGING", "Daging", 2113, "terverifikasi"),
        ("PEN-TELUR", "Telur dan Susu", 2115, "terverifikasi"),
        ("PEN-SAYUR", "Sayur-Sayuran", 2116, "terverifikasi"),
        ("PEN-KACANG", "Kacang-Kacangan", 2117, "pola"),
        ("PEN-BUAH", "Buah-Buahan", 2118, "pola"),
        ("PEN-MINYAK", "Minyak dan Kelapa", 2119, "terverifikasi"),
        ("PEN-MINUMAN", "Bahan Minuman", 2120, "terverifikasi"),
        ("PEN-BUMBU", "Bumbu-Bumbuan", 2122, "terverifikasi"),
        ("PEN-JADI", "Makanan Minuman Jadi", 2123, "pola"),
    ]
    for kode, kel, var, st in pengeluaran:
        j = f"Rata-rata Pengeluaran Perkapita Seminggu Menurut Kelompok {kel} Per Kabupaten/kota"
        t.append({"kode": kode, "judul": j,
                  "pola_file": rf"^Rata-rata Pengeluaran Perkapita Seminggu +Menurut Kelompok {re.escape(kel)} Per Kabupaten_kota",
                  **bps(var, j, st, "" if st == "terverifikasi" else
                        "Nomor ditarik dari urutan seri konsumsi (sayur, kacang, buah, minyak) yang "
                        "berpadanan dengan seri pengeluaran 2116 sampai 2119; cek judul saat halaman dibuka.")})
    pengeluaran25 = [
        ("PEN25-IKAN", "Ikan", 2551, "pola"),
        ("PEN25-DAGING", "Daging", 2553, "pola"),
        ("PEN25-TELUR", "Telur dan Susu", 2555, "terverifikasi"),
        ("PEN25-SAYUR", "Sayur-Sayuran", 2557, "pola"),
        ("PEN25-BUAH", "Buah-Buahan", 2559, "pola"),
    ]
    for kode, kel, var, st in pengeluaran25:
        j = f"Rata-rata Pengeluaran Perkapita Seminggu Menurut Kelompok {kel} Per Kabupaten/Kota-Komoditas 2025"
        t.append({"kode": kode, "judul": j,
                  "pola_file": rf"^Rata-rata Pengeluaran Perkapita Seminggu +Menurut Kelompok {re.escape(kel)} Per Kabupaten_Kota-Komoditas 2025",
                  **bps(var, j, st, "" if st == "terverifikasi" else
                        "Nomor ditarik dari pola pasangan konsumsi genap/pengeluaran ganjil "
                        "(2554 konsumsi telur, 2555 pengeluaran telur); cek judul saat halaman dibuka.")})

    # 5. TPT dan TPAK Aceh (BPS Provinsi Aceh), satu URL per tahun lewat parameter ?year=
    t.append({"kode": "TPT-ACEH",
              "judul": ("Tingkat Pengangguran Terbuka (TPT) dan Tingkat Partisipasi Angkatan Kerja (TPAK) "
                        "Menurut Kabupaten/Kota di Provinsi Aceh"),
              "pola_file": r"^Tingkat Pengangguran Terbuka \(TPT\) dan Tingkat Partisipasi Angkatan Kerja",
              "penerbit": "BPS Provinsi Aceh", "situs": "aceh.bps.go.id",
              "jenis": "Tabel Statistik (tabel per tahun, dipilih dengan parameter ?year=)",
              "var_id": None, "kode_url": "V2pOVWJWcHJURGg0U2pONFJYaExhVXB0TUhacVFUMDkjMw==",
              "url_tabel_dinamis": QB_ACEH,
              "url": TPT_ACEH.format(tahun=2025), "url_per_tahun": TPT_ACEH, "status": "terverifikasi",
              "catatan": "Halaman yang sama melayani tahun lain dengan mengganti nilai ?year= (terindeks untuk 2019, 2022, 2025)."})

    # 5b. Tabel statis BPS Provinsi Aceh (.xls): nama file adalah slug judul tabel BPS.
    t.append({"kode": "KONPEN-ACEH",
              "judul": ("Rata-rata Konsumsi dan Pengeluaran Perkapita Seminggu Menurut Komoditi Makanan "
                        "dan Golongan Pengeluaran per Kapita Seminggu di Provinsi Aceh, 2018-2025"),
              "pola_file": r"^rata-rata-konsumsi-dan-pengeluaran-perkapita-seminggu--menurut-komoditi-makanan",
              "penerbit": "BPS Provinsi Aceh", "situs": "aceh.bps.go.id",
              "jenis": "Tabel Statistik statis (statistics-table/1), unduhan .xls",
              "var_id": None, "kode_url": "",
              "url": "https://aceh.bps.go.id/id/statistics-table",
              "status": "belum_ditemukan",
              "catatan": ("Nama file sama dengan slug judul tabel statis BPS, dan isi lembar 'indo' berjudul "
                          "'Rata-rata Konsumsi Perkapita Seminggu Menurut Komoditi Makanan dan Golongan Pengeluaran "
                          "per Kapita Seminggu, 2018-2025, PROVINSI ACEH' (metadata file: disimpan 23 Desember 2025). "
                          "Kode tabel statis BPS terenkripsi dan halaman ini tidak terindeks mesin pencari, sehingga "
                          "URL persisnya belum dapat dipastikan dari sesi ini. Telusuri dari menu Tabel Statistik "
                          "aceh.bps.go.id, subjek Konsumsi dan Pengeluaran, dengan judul di atas.")})

    # 6. DJPK Kemenkeu
    t.append({"kode": "DJPK-APBD", "judul": "Portal Data DJPK: Ringkasan APBD per Pemerintah Daerah",
              "pola_file": r"^apbd_kabupaten_kota\.xlsx$",
              "penerbit": "Direktorat Jenderal Perimbangan Keuangan (DJPK), Kementerian Keuangan",
              "situs": "djpk.kemenkeu.go.id", "jenis": "Halaman portal data (ditarik dengan skrip tarik_apbd_djpk.py)",
              "var_id": None, "kode_url": "",
              "url": "https://djpk.kemenkeu.go.id/portal/data/apbd?periode=12&tahun=2025&provinsi=01&pemda=12",
              "url_dasar": "https://djpk.kemenkeu.go.id/portal/data/apbd",
              "url_unduh_csv": "https://djpk.kemenkeu.go.id/portal/csv_apbd?type=apbd&periode=12&tahun=2026&provinsi=--&pemda=--",
              "status": "terverifikasi",
              "catatan": ("URL persis setiap baris tersimpan di kolom url_sumber (5.353 URL unik, 2016 sampai 2026, "
                          "37 provinsi). Contoh pada screenshot: Kota Banda Aceh (provinsi=01, pemda=12), TA 2025.")})
    t.append({"kode": "DJPK-TKDD", "judul": "Portal Data DJPK: Transfer ke Daerah dan Dana Desa (TKDD) per Pemerintah Daerah",
              "pola_file": r"^tkdd_kabupaten_kota\.xlsx$",
              "penerbit": "Direktorat Jenderal Perimbangan Keuangan (DJPK), Kementerian Keuangan",
              "situs": "djpk.kemenkeu.go.id", "jenis": "Halaman portal data (ditarik dengan skrip tarik_tkdd_djpk.py)",
              "var_id": None, "kode_url": "",
              "url": "https://djpk.kemenkeu.go.id/portal/data/tkdd?tahun=2025&provinsi=01&pemda=12",
              "url_dasar": "https://djpk.kemenkeu.go.id/portal/data/tkdd",
              "status": "terverifikasi",
              "catatan": ("URL persis setiap baris tersimpan di kolom url_sumber (5.080 URL unik, 2017 sampai 2026). "
                          "Contoh pada screenshot: Kota Banda Aceh, TA 2025.")})
    t.append({"kode": "DJPK-LOG", "judul": "Log proses penarikan APBD (bukan dataset unduhan)",
              "pola_file": r"^apbd_(errors|no_data)\.csv$",
              "penerbit": "Dihasilkan skrip tarik_apbd_djpk.py", "situs": "djpk.kemenkeu.go.id",
              "jenis": "Catatan galat dan kombinasi tanpa data", "var_id": None, "kode_url": "",
              "url": "https://djpk.kemenkeu.go.id/portal/data/apbd?periode=12&tahun=2016&provinsi=26&pemda=03",
              "status": "terverifikasi",
              "catatan": ("apbd_errors.csv berisi 186 kombinasi tahun|periode|provinsi|pemda yang dijawab server dengan "
                          "HTTP 500; apbd_no_data.csv berisi 49 kombinasi yang halamannya tidak memuat tabel. "
                          "URL yang dicoba tertulis lengkap di kolom error. Contoh pada screenshot adalah baris pertama apbd_errors.csv.")})
    # Sumber yang nomornya semula disimpulkan dari pola, lalu dikonfirmasi lewat judul halaman yang tampil
    # saat ambil_screenshot.py dijalankan di komputer pengguna (24-11-2026).
    for x in t:
        if x["status"] == "pola" and x["kode"] in KONFIRMASI_BROWSER:
            x["status"] = "terverifikasi"
            x["catatan"] = ("Nomor tabel semula disimpulkan dari pola penomoran; judul halaman yang tampil "
                            "saat dibuka di browser pada 24-11-2026 sama persis dengan judul tabel ini.")
    return t


KONFIRMASI_BROWSER = {
    "IHK-K05", "IHK-K07", "IHK-K08", "IHK-K09",
    "YTD-K02", "YTD-K04", "YTD-K06", "YTD-K07", "YTD-K08", "YTD-K09", "YTD-K10", "YTD-K11",
    "YOY-K02", "YOY-K08", "KON25-BUAH", "PEN-KACANG", "PEN-BUAH", "PEN-JADI",
    "PEN25-IKAN", "PEN25-DAGING", "PEN25-SAYUR", "PEN25-BUAH",
}


def main() -> None:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else "Dataset")
    files = sorted(p for p in root.rglob("*") if p.suffix.lower() in {".xlsx", ".xls", ".csv"})
    tabel = daftar_tabel()
    for x in tabel:
        x["file"] = []
    tanpa_pasangan = []
    for f in files:
        rel = f.relative_to(root).as_posix()
        cocok = [x for x in tabel if re.search(x["pola_file"], f.name)]
        if len(cocok) != 1:
            tanpa_pasangan.append((rel, [c["kode"] for c in cocok]))
            continue
        m = re.search(r", ?(\d{4})\.xlsx$", f.name)
        cocok[0]["file"].append({"path": rel, "tahun": m.group(1) if m else ""})
    if tanpa_pasangan:
        for r, c in tanpa_pasangan:
            print("TIDAK TERPETAKAN / GANDA:", r, c)
        sys.exit(1)
    kosong = [x["kode"] for x in tabel if not x["file"]]
    tabel = [x for x in tabel if x["file"]]
    out = Path(__file__).with_name("katalog_sumber.json")
    out.write_text(json.dumps({"jumlah_file": len(files), "tabel": tabel}, ensure_ascii=False, indent=1))
    n = sum(len(x["file"]) for x in tabel)
    print(f"{len(files)} file Excel/CSV, {n} terpetakan ke {len(tabel)} sumber. "
          f"Sumber tanpa file (dibuang): {kosong}")


if __name__ == "__main__":
    main()
