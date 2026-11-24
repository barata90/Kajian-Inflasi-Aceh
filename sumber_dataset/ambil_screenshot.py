#!/usr/bin/env python3
"""Ambil screenshot halaman sumber setiap dataset, lalu susun ulang dokumen Word.

Dijalankan di komputer sendiri (bukan server cloud), karena BPS dan DJPK menolak akses dari pusat data.
Browser dibuka dalam mode terlihat (headed) dengan profil tetap, sehingga verifikasi Cloudflare cukup
diselesaikan manual satu kali.

    pip install playwright python-docx
    python -m playwright install chromium
    python ambil_screenshot.py                      # semua sumber yang belum punya screenshot
    python ambil_screenshot.py --ulang              # timpa semua screenshot
    python ambil_screenshot.py IHK150 TPT-ACEH      # hanya kode tertentu
    python ambil_screenshot.py --manual KONPEN-ACEH # buka halaman awal, kita navigasi sendiri ke tabelnya,
                                                    # lalu tekan Enter di terminal; URL akhir ikut direkam

Setiap pengambilan menulis ss/<kode>.json (URL akhir, judul halaman, kecocokan judul, waktu akses)
dan ss/<kode>.txt (keterangan gambar). Ringkasan semua sumber ditulis ke ss/laporan_screenshot.csv.
"""
from __future__ import annotations

import json
import re
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


def norm(s: str) -> str:
    s = re.split(r" - (Tabel Statistik|Statistical Data|Badan Pusat Statistik)", s or "")[0]
    return re.sub(r"[^a-z0-9]+", "", s.lower())


def judul_cocok(judul_halaman: str, judul_katalog: str) -> bool:
    a, b = norm(judul_halaman), norm(judul_katalog)
    return bool(a) and (a == b or a.startswith(b) or b.startswith(a))


def tunggu_lolos(page, batas_detik=300) -> bool:
    """Tunggu sampai halaman bukan lagi halaman verifikasi/blokir. Verifikasi diselesaikan oleh pengguna."""
    awal = time.time()
    while time.time() - awal < batas_detik:
        try:
            isi = (page.title() or "") + " " + page.inner_text("body")[:500]
        except Exception:
            isi = ""
        if isi.strip() and not any(t in isi for t in TANDA_TANTANGAN):
            return True
        print("   ... menunggu halaman/verifikasi selesai di jendela browser")
        time.sleep(3)
    return False


class Browser:
    """Konteks Chromium yang dibuka ulang otomatis bila jendelanya tertutup."""

    def __init__(self, pw):
        self.pw, self.ctx, self.page = pw, None, None

    def buka(self):
        if self.ctx is not None:
            try:
                self.ctx.close()
            except Exception:
                pass
        self.ctx = self.pw.chromium.launch_persistent_context(
            str(PROFIL), headless=False, viewport={"width": 1440, "height": 900},
            device_scale_factor=1.5, locale="id-ID")
        self.page = self.ctx.pages[0] if self.ctx.pages else self.ctx.new_page()

    def halaman(self):
        try:
            if self.page is None or self.page.is_closed():
                raise RuntimeError
            self.page.title()
        except Exception:
            print("   (browser tertutup, membuka ulang)")
            self.buka()
        return self.page


def bersihkan_layar(page) -> None:
    """Tutup pop-up pengumuman (mis. Maklumat Pelayanan BPS Aceh) dan menu yang terbuka karena kursor."""
    for _ in range(3):
        ditutup = False
        for nama in ("Tutup", "Close", "Tutup Pengumuman"):
            try:
                tombol = page.get_by_role("button", name=nama, exact=True)
                for i in range(tombol.count()):
                    if tombol.nth(i).is_visible():
                        tombol.nth(i).click(timeout=2_000)
                        ditutup = True
            except Exception:
                pass
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
        if not ditutup:
            break
        time.sleep(1)
    try:
        vp = page.viewport_size or {"width": 1440, "height": 900}
        page.mouse.move(vp["width"] - 10, vp["height"] - 10)
    except Exception:
        pass
    time.sleep(1.5)


def simpan(t: dict, page, mode: str, png: Path) -> dict | None:
    judul = page.title()
    if not judul.strip() or page.url in ("", "about:blank"):
        print("   TIDAK disimpan: halaman kosong (judul/URL kosong). Ulangi untuk kode ini.")
        return None
    bersihkan_layar(page)
    page.screenshot(path=str(png), full_page=False)
    info = {
        "kode": t["kode"], "mode": mode, "url_katalog": t["url"], "url_akhir": page.url,
        "judul_halaman": judul, "judul_katalog": t["judul"],
        "judul_cocok": judul_cocok(judul, t.get("judul_halaman_harapan") or t["judul"]),
        "waktu_akses": datetime.now().strftime("%d-%m-%Y %H:%M"),
    }
    (SS / f"{t['kode']}.json").write_text(json.dumps(info, ensure_ascii=False, indent=1))
    (SS / f"{t['kode']}.txt").write_text(
        f"Diakses {info['waktu_akses']}. Judul halaman: {judul}. URL: {page.url}")
    tanda = "cocok" if info["judul_cocok"] else "JUDUL TIDAK COCOK, periksa"
    print(f"   tersimpan -> {png.name} | {judul} | {tanda}")
    return info


def otomatis(br: Browser, t: dict, png: Path):
    for coba in (1, 2):
        page = br.halaman()
        try:
            page.goto(t["url"], wait_until="domcontentloaded", timeout=90_000)
            if not tunggu_lolos(page):
                print("   dilewati: verifikasi/halaman tidak selesai")
                return None
            try:
                page.wait_for_load_state("networkidle", timeout=60_000)
            except Exception:
                pass
            for sel in ("table", "main"):
                try:
                    page.wait_for_selector(sel, timeout=15_000)
                    break
                except Exception:
                    pass
            time.sleep(3)
            return simpan(t, page, "otomatis", png)
        except Exception as e:
            print(f"   gagal (percobaan {coba}): {str(e).splitlines()[0]}")
    return None


def manual(br: Browser, t: dict, png: Path):
    page = br.halaman()
    awal = t.get("url_tabel_dinamis") or t["url"]
    if t["kode"] == "KONPEN-ACEH":
        awal = t["url"]
    page.goto(awal, wait_until="domcontentloaded", timeout=90_000)
    print("\n   MODE MANUAL. Di jendela browser, buka tabel berjudul:")
    print(f"     {t['judul']}")
    tahun = [f['tahun'] for f in t['file'] if f['tahun']]
    if tahun:
        print(f"   Tahun pada file: {', '.join(tahun)}")
    if t.get("catatan"):
        print(f"   Catatan: {t['catatan'][:300]}")
    input("   Setelah tabel tampil penuh di layar, tekan Enter di sini untuk memotret... ")
    # Tabel sering terbuka di tab baru: potret tab terbuka yang paling akhir, bukan tab awal.
    try:
        terbuka = [p for p in br.ctx.pages if not p.is_closed()]
    except Exception:
        terbuka = []
    if not terbuka:
        print("   Jendela browser sudah tertutup sebelum Enter ditekan; tidak ada yang dipotret.")
        return None
    hal = terbuka[-1]
    hal.bring_to_front()
    return simpan(t, hal, "manual", png)


def main() -> None:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    ulang, mode_manual = "--ulang" in sys.argv, "--manual" in sys.argv
    tabel = json.loads((HERE / "katalog_sumber.json").read_text())["tabel"]
    SS.mkdir(exist_ok=True)
    with sync_playwright() as pw:
        br = Browser(pw)
        br.buka()
        for t in tabel:
            kode = t["kode"]
            png = SS / f"{kode}.png"
            if args and kode not in args:
                continue
            if png.exists() and not (ulang or mode_manual):
                continue
            print(f"[{kode}] {t['url']}")
            try:
                (manual if mode_manual else otomatis)(br, t, png)
            except Exception as e:
                print("   gagal:", str(e).splitlines()[0])
        try:
            br.ctx.close()
        except Exception:
            pass

    # Ringkasan ss/laporan_screenshot.csv ditulis oleh buat_dokumen.py.
    subprocess.run([sys.executable, str(HERE / "buat_dokumen.py")], check=True)


if __name__ == "__main__":
    main()
