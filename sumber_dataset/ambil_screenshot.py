#!/usr/bin/env python3
"""Ambil screenshot halaman sumber setiap dataset, lalu susun ulang dokumen Word.

Dijalankan di komputer sendiri (bukan server cloud), karena BPS dan DJPK menolak akses dari pusat data.
Browser dibuka dalam mode terlihat (headed) dengan profil tetap, sehingga verifikasi Cloudflare cukup
diselesaikan manual satu kali.

    pip install playwright python-docx
    python -m playwright install chromium
    python ambil_screenshot.py            # semua sumber yang belum punya screenshot
    python ambil_screenshot.py --ulang    # timpa semua screenshot
    python ambil_screenshot.py IHK150 TPT-ACEH   # hanya kode tertentu
"""
from __future__ import annotations

import json
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
SS = HERE / "ss"
PROFIL = HERE / ".profil_browser"
TANDA_TANTANGAN = ("Just a moment", "Performing security verification", "Verify you are human",
                   "Web Page Blocked")


def tunggu_lolos(page, batas_detik=300):
    """Tunggu sampai halaman bukan lagi halaman verifikasi/blokir. Penyelesaian verifikasi oleh pengguna."""
    awal = time.time()
    while time.time() - awal < batas_detik:
        isi = (page.title() or "") + " " + page.inner_text("body")[:500]
        if not any(t in isi for t in TANDA_TANTANGAN):
            return True
        print("   ... menunggu verifikasi diselesaikan di jendela browser")
        time.sleep(3)
    return False


def main() -> None:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    ulang = "--ulang" in sys.argv
    tabel = json.loads((HERE / "katalog_sumber.json").read_text())["tabel"]
    SS.mkdir(exist_ok=True)
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(str(PROFIL), headless=False,
                                                   viewport={"width": 1440, "height": 900},
                                                   device_scale_factor=1.5, locale="id-ID")
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        for t in tabel:
            kode, url = t["kode"], t["url"]
            png = SS / f"{kode}.png"
            if args and kode not in args:
                continue
            if png.exists() and not ulang:
                continue
            print(f"[{kode}] {url}")
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=90_000)
                if not tunggu_lolos(page):
                    print("   dilewati: verifikasi tidak selesai")
                    continue
                page.wait_for_load_state("networkidle", timeout=60_000)
                # Tabel dinamis BPS dirender di sisi klien; beri waktu agar sel terisi.
                for sel in ("table", "#tabel", "main"):
                    try:
                        page.wait_for_selector(sel, timeout=15_000)
                        break
                    except Exception:
                        pass
                time.sleep(3)
                page.screenshot(path=str(png), full_page=False)
                judul = page.title()
                (SS / f"{kode}.txt").write_text(
                    f"Diakses {datetime.now():%d-%m-%Y %H:%M}. Judul halaman: {judul}. URL: {page.url}")
                print(f"   tersimpan -> {png.name} | {judul}")
            except Exception as e:  # lanjut ke sumber berikutnya
                print("   gagal:", e)
        ctx.close()
    subprocess.run([sys.executable, str(HERE / "buat_dokumen.py")], check=True)


if __name__ == "__main__":
    main()
