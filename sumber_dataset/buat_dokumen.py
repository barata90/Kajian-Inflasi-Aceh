#!/usr/bin/env python3
"""Bangun dokumen Word berisi URL sumber web setiap file Excel/CSV di folder Dataset.

Masukan : katalog_sumber.json (hasil susun_katalog.py) dan folder ss/ berisi <kode>.png.
Keluaran: Sumber_Web_Dataset_Inflasi_Aceh.docx

Screenshot yang ada di ss/ disisipkan otomatis; yang belum ada diberi kotak penanda.
"""
from __future__ import annotations

import json
import re
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = Path(__file__).resolve().parent
KATALOG = HERE / "katalog_sumber.json"
SS_DIR = HERE / "ss"
OUT = HERE / "Sumber_Web_Dataset_Inflasi_Aceh.docx"
REPO = "https://github.com/barata90/Inflasi-Aceh/tree/main/Dataset"

LABEL_STATUS = {
    "terverifikasi": "Terverifikasi",
    "pola": "Disimpulkan dari pola nomor tabel (wajib dicek)",
    "belum_ditemukan": "URL persis belum ditemukan",
}
WARNA_STATUS = {"terverifikasi": "1E7B34", "pola": "B26A00", "belum_ditemukan": "B3261E"}
GRUP = [
    ("A. Indeks Harga Konsumen dan Inflasi 150 Kabupaten/Kota (BPS RI)", ("IHK", "YTD", "YOY")),
    ("B. Ketahanan Pangan (BPS RI)", ("FIES", "POU")),
    ("C. Konsumsi per Kapita Seminggu (BPS RI, Susenas)", ("KON-", "KON25")),
    ("D. Pengeluaran per Kapita Seminggu (BPS RI, Susenas)", ("PEN-", "PEN25")),
    ("E. Tabel Khusus Provinsi Aceh (BPS)", ("TPT", "KONPEN")),
    ("F. Keuangan Daerah (DJPK Kementerian Keuangan)", ("DJPK",)),
]


def judul_cocok(judul_halaman: str, judul_harapan: str) -> bool:
    def norm(s: str) -> str:
        s = re.split(r" - (Tabel Statistik|Statistical Data|Badan Pusat Statistik)", s or "")[0]
        return re.sub(r"[^a-z0-9]+", "", s.lower())
    a, b = norm(judul_halaman), norm(judul_harapan)
    return bool(a) and (a == b or a.startswith(b) or b.startswith(a))


def hyperlink(par, url: str, teks: str | None = None, ukuran: int = 9):
    r_id = par.part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                              is_external=True)
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), r_id)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    for tag, val in (("w:color", "0563C1"), ("w:u", "single"), ("w:sz", str(ukuran * 2))):
        e = OxmlElement(tag)
        e.set(qn("w:val"), val)
        rpr.append(e)
    run.append(rpr)
    t = OxmlElement("w:t")
    t.text = teks or url
    t.set(qn("xml:space"), "preserve")
    run.append(t)
    h.append(run)
    par._p.append(h)


def shade(cell, hex_fill: str):
    tcpr = cell._tc.get_or_add_tcPr()
    s = OxmlElement("w:shd")
    s.set(qn("w:val"), "clear")
    s.set(qn("w:color"), "auto")
    s.set(qn("w:fill"), hex_fill)
    tcpr.append(s)


def kotak(par, warna="999999"):
    ppr = par._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    for sisi in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{sisi}")
        e.set(qn("w:val"), "dashed")
        e.set(qn("w:sz"), "6")
        e.set(qn("w:space"), "6")
        e.set(qn("w:color"), warna)
        bdr.append(e)
    ppr.append(bdr)


def teks_sel(cell, teks: str, tebal=False, ukuran=9, warna: str | None = None, miring=False):
    p = cell.paragraphs[0]
    r = p.add_run(teks)
    r.bold, r.italic, r.font.size = tebal, miring, Pt(ukuran)
    if warna:
        r.font.color.rgb = RGBColor.from_string(warna)
    return p


def lebar_kolom(tabel, lebar_cm):
    tabel.autofit = False
    grid = tabel._tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), lebar_cm):
        gc.set(qn("w:w"), str(int(w * 567)))
    for col, w in zip(tabel.columns, lebar_cm):
        col.width = Cm(w)
    for row in tabel.rows:
        for c, w in zip(row.cells, lebar_cm):
            c.width = Cm(w)


def grup_dari(kode: str) -> int:
    for i, (_, awalan) in enumerate(GRUP):
        if kode.startswith(awalan):
            return i
    return len(GRUP) - 1


def main() -> None:
    data = json.loads(KATALOG.read_text())
    tabel = data["tabel"]
    n_file = sum(len(t["file"]) for t in tabel)
    ada_ss = {t["kode"] for t in tabel if (SS_DIR / f"{t['kode']}.png").exists()}
    # Hasil mode manual ambil_screenshot.py: URL akhir yang dibuka pengguna menjadi URL persis sumber.
    for t in tabel:
        j = SS_DIR / f"{t['kode']}.json"
        if j.exists():
            info = json.loads(j.read_text())
            # Hitung ulang kecocokan judul memakai judul halaman yang diharapkan bila ada
            # (portal DJPK memakai judul seragam "Portal Data SIKD" di semua halaman).
            info["judul_cocok"] = judul_cocok(info.get("judul_halaman", ""),
                                              t.get("judul_halaman_harapan") or t["judul"])
            t["_info"] = info
            if (info.get("mode") == "manual" and str(info.get("url_akhir", "")).startswith("http")
                    and info.get("judul_halaman", "").strip()):
                t["url"] = info["url_akhir"]
                t["status"] = "terverifikasi"
                t["catatan"] = (t.get("catatan", "") + " URL di atas direkam dari browser saat tabel dibuka "
                                f"secara manual pada {info.get('waktu_akses', '')}; judul halaman: "
                                f"{info.get('judul_halaman', '')}.").strip()

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.orientation = WD_ORIENT.PORTRAIT
    for side in ("left_margin", "right_margin"):
        setattr(sec, side, Cm(2))
    sec.top_margin = sec.bottom_margin = Cm(2)
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "Calibri", Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")

    doc.add_heading("Sumber Web Dataset Kajian Inflasi Aceh", 0)
    p = doc.add_paragraph()
    p.add_run("Folder sumber: ")
    hyperlink(p, REPO, ukuran=10)
    p.add_run(f"\nDisusun {date.today():%d-%m-%Y}. Mencakup {n_file} file Excel/CSV "
              f"({len(tabel)} tabel sumber).")

    doc.add_heading("Ringkasan", 1)
    from collections import Counter
    c = Counter(t["status"] for t in tabel)
    fc = Counter()
    for t in tabel:
        fc[t["status"]] += len(t["file"])
    doc.add_paragraph(
        f"Seluruh {n_file} file Excel dan CSV di folder Dataset (254 .xlsx, 1 .xls, 2 .csv) sudah dipasangkan "
        f"ke tabel sumbernya. File yang hanya berbeda tahun berasal dari satu halaman tabel yang sama, "
        f"sehingga {n_file} file tersebut bermuara ke {len(tabel)} halaman sumber: tabel statistik BPS RI "
        f"(www.bps.go.id), tabel BPS Provinsi Aceh (aceh.bps.go.id), dan Portal Data DJPK Kementerian Keuangan "
        f"(djpk.kemenkeu.go.id). Pemasangan dilakukan dari judul tabel yang tercetak di baris pertama setiap file, "
        f"lalu dicocokkan dengan judul halaman BPS yang terindeks mesin pencari dan dengan judul halaman yang "
        f"tampil saat setiap URL dibuka di browser. Untuk file DJPK, URL persis per baris sudah tersimpan di "
        f"kolom url_sumber di dalam file itu sendiri.")
    doc.add_paragraph(
        "File BPS diunduh lewat fitur Tabel Dinamis (query builder) situs BPS Provinsi Aceh. Halaman itu dipakai "
        "untuk memilih tabel dan tahun, sedangkan tautan permanen tiap tabel adalah halaman Tabel Statistik "
        "pada baris URL. Keduanya dicantumkan di rincian per sumber.")
    rincian = [f"{c['terverifikasi']} sumber ({fc['terverifikasi']} file) terverifikasi"]
    if c["pola"]:
        rincian.append(f"{c['pola']} sumber ({fc['pola']} file) disimpulkan dari pola penomoran tabel BPS")
    if c["belum_ditemukan"]:
        rincian.append(f"{c['belum_ditemukan']} sumber ({fc['belum_ditemukan']} file) belum ditemukan URL persisnya")
    doc.add_paragraph(
        "Status URL: " + (", ".join(rincian[:-1]) + ", dan " + rincian[-1] if len(rincian) > 1 else rincian[0])
        + f". Tangkapan layar (screenshot) tersedia untuk {len(ada_ss)} dari {len(tabel)} sumber; "
          "judul halaman yang tampil saat pemotretan tercantum di keterangan setiap gambar.")

    doc.add_heading("Cara membaca", 2)
    butir = [
        ("Terverifikasi", "nomor tabel di URL cocok dengan judul tabel, dibuktikan oleh hasil mesin pencari "
                          "dan/atau oleh judul halaman yang tampil saat URL dibuka di browser."),
    ]
    if c["pola"]:
        butir.append(("Disimpulkan dari pola", "BPS memberi nomor berurutan untuk satu seri tabel. Nomor yang "
                                               "belum terindeks diisi mengikuti urutan itu, sehingga judul "
                                               "halaman tetap perlu dicek saat dibuka."))
    butir += [
        ("Kode di URL", "bagian seperti MjIxMiMy adalah base64 dari \"2212#2\" (nomor tabel #2 = tabel dinamis). "
                        "Teks setelah kode hanya hiasan; BPS membuka tabel berdasarkan kode."),
        ("Memilih tahun", "tabel dinamis BPS menampilkan tahun terbaru. File berakhiran \", 2018\" dan seterusnya "
                          "diperoleh dengan memilih tahun itu pada pemilih tahun di halaman yang sama, atau di "
                          "Tabel Dinamis. Tabel TPT Aceh memakai parameter ?year= di URL."),
    ]
    for judul, isi in butir:
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(judul + ": ").bold = True
        p.add_run(isi)

    # Daftar ringkas
    doc.add_heading("Daftar sumber", 1)
    t = doc.add_table(rows=1, cols=5)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for cell, h in zip(t.rows[0].cells, ["No", "Judul tabel sumber", "URL", "Status", "File"]):
        teks_sel(cell, h, tebal=True)
        shade(cell, "DCE6F1")
    for i, s in enumerate(tabel, 1):
        r = t.add_row().cells
        teks_sel(r[0], str(i))
        teks_sel(r[1], s["judul"], ukuran=8.5)
        hyperlink(r[2].paragraphs[0], s["url"], s["situs"] + (f" (#{s['var_id']})" if s.get("var_id") else ""), 8)
        teks_sel(r[3], LABEL_STATUS[s["status"]].split(" (")[0], ukuran=8, warna=WARNA_STATUS[s["status"]])
        teks_sel(r[4], str(len(s["file"])))
    lebar_kolom(t, [0.9, 8.2, 4.4, 2.6, 1.0])

    # Rincian per sumber
    doc.add_page_break()
    doc.add_heading("Rincian per sumber", 1)
    grup_aktif = -1
    for i, s in enumerate(tabel, 1):
        g = grup_dari(s["kode"])
        if g != grup_aktif:
            grup_aktif = g
            doc.add_heading(GRUP[g][0], 2)
        doc.add_heading(f"{i}. {s['judul']}", 3)
        m = doc.add_table(rows=0, cols=2)
        m.style = "Table Grid"
        baris = [
            ("Penerbit", s["penerbit"]),
            ("Jenis halaman", s["jenis"]),
            ("URL", None),
            ("Status URL", LABEL_STATUS[s["status"]]),
        ]
        if s.get("url_tabel_dinamis"):
            baris.append(("Tabel Dinamis (tempat unduh)", s["url_tabel_dinamis"]))
            baris.append(("Langkah di Tabel Dinamis",
                          "Pilih Tabel/Indikator dengan judul di atas, centang tahun yang tercantum di baris "
                          "\"Tahun yang diunduh\", tampilkan, lalu unduh sebagai Excel."))
        if s.get("url_skrip"):
            baris.append(("Skrip penarik data", s["url_skrip"]))
        if s.get("url_per_tahun"):
            baris.append(("URL per tahun", s["url_per_tahun"].replace("{tahun}", "<tahun>")))
        if s.get("url_dasar"):
            baris.append(("URL dasar", s["url_dasar"]))
        if s.get("url_unduh_csv"):
            baris.append(("Tautan unduh CSV di portal", s["url_unduh_csv"]))
        tahun = [f["tahun"] for f in s["file"] if f["tahun"]]
        if tahun:
            baris.append(("Tahun yang diunduh", ", ".join(tahun)))
        if s.get("catatan"):
            baris.append(("Catatan", s["catatan"]))
        for k, v in baris:
            r = m.add_row().cells
            teks_sel(r[0], k, tebal=True)
            shade(r[0], "F2F2F2")
            if k == "URL":
                hyperlink(r[1].paragraphs[0], s["url"])
            elif k.startswith("URL") or k.startswith("Tautan") or k.startswith("Skrip") or k.startswith("Tabel Dinamis"):
                hyperlink(r[1].paragraphs[0], v.replace("<tahun>", "2025") if "<tahun>" in v else v, v)
            elif k == "Status URL":
                teks_sel(r[1], v, tebal=True, warna=WARNA_STATUS[s["status"]])
            else:
                teks_sel(r[1], v)
        r = m.add_row().cells
        teks_sel(r[0], f"File di repo ({len(s['file'])})", tebal=True)
        shade(r[0], "F2F2F2")
        p = r[1].paragraphs[0]
        for j, f in enumerate(s["file"]):
            if j:
                p.add_run("\n")
            p.add_run("Dataset/" + f["path"]).font.size = Pt(8)
        lebar_kolom(m, [3.4, 13.6])

        cap = doc.add_paragraph()
        cap.paragraph_format.space_before = Pt(6)
        png = SS_DIR / f"{s['kode']}.png"
        if png.exists():
            doc.add_picture(str(png), width=Cm(16.5))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            meta = SS_DIR / f"{s['kode']}.txt"
            ket = meta.read_text().strip() if meta.exists() else ""
            cap = doc.add_paragraph()
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            rr = cap.add_run(f"Tampilan halaman sumber {s['kode']}. {ket}")
            rr.italic, rr.font.size = True, Pt(8.5)
            if s.get("_info") and s["_info"].get("judul_cocok") is False:
                rw = cap.add_run(" Peringatan: judul halaman tidak sama dengan judul tabel, periksa URL ini.")
                rw.bold, rw.font.size = True, Pt(8.5)
                rw.font.color.rgb = RGBColor.from_string("B3261E")
        else:
            kotak(cap)
            rr = cap.add_run(f"Screenshot {s['kode']} belum tersedia. ")
            rr.bold, rr.font.size = True, Pt(9)
            rr = cap.add_run("Situs sumber menolak akses dari server cloud tempat dokumen ini disusun. "
                             "Jalankan ambil_screenshot.py dari komputer Anda; gambar akan masuk ke slot ini.")
            rr.font.size = Pt(9)
        doc.add_paragraph()

    doc.add_heading("Mengisi screenshot", 1)
    for langkah in [
        "Pasang Python 3 lalu jalankan: pip install playwright python-docx, kemudian: python -m playwright install chromium",
        "Dari folder sumber_dataset, jalankan: python ambil_screenshot.py",
        "Jendela Chromium terbuka. Bila muncul verifikasi Cloudflare BPS, selesaikan secara manual; skrip menunggu "
        "sampai tabel tampil, menyimpan ss/<kode>.png, lalu menyusun ulang dokumen ini.",
        "Skrip mencatat judul halaman yang tampil dan kecocokannya dengan judul tabel ke ss/<kode>.json, serta "
        "ringkasan semua sumber ke ss/laporan_screenshot.csv.",
        "Untuk tabel yang URL-nya belum diketahui, jalankan: python ambil_screenshot.py --manual <KODE>. "
        "Buka tabelnya sendiri di jendela browser, lalu tekan Enter di terminal; URL akhir ikut direkam dan "
        "dipakai sebagai URL sumber di dokumen ini.",
    ]:
        doc.add_paragraph(langkah, style="List Number")

    doc.save(OUT)
    tulis_daftar_link(tabel)
    print(f"Tersimpan: {OUT} | screenshot tersisip: {len(ada_ss)}/{len(tabel)}")


def tulis_daftar_link(tabel: list[dict]) -> None:
    """Daftar semua link dalam satu lembar Excel, berdampingan dengan dokumen Word."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
    except ImportError:
        print("openpyxl tidak terpasang; Daftar_Link_Sumber_Dataset.xlsx dilewati")
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Link sumber"
    kepala = ["No", "Kode", "Judul tabel sumber", "Penerbit", "URL sumber", "Tabel Dinamis (tempat unduh)",
              "Skrip penarik", "Status URL", "Jumlah file", "Tahun", "File di repo (Dataset/...)",
              "Screenshot", "Judul halaman saat dipotret"]
    ws.append(kepala)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DCE6F1")
    for i, t in enumerate(tabel, 1):
        info = t.get("_info", {})
        judul_hal = info.get("judul_halaman", "")
        if not judul_hal:
            meta = SS_DIR / f"{t['kode']}.txt"
            if meta.exists():
                m = re.search(r"Judul halaman: (.*?)\. URL:", meta.read_text())
                judul_hal = m.group(1) if m else ""
        ws.append([i, t["kode"], t["judul"], t["penerbit"], t["url"], t.get("url_tabel_dinamis", ""),
                   t.get("url_skrip", ""), LABEL_STATUS[t["status"]], len(t["file"]),
                   ", ".join(f["tahun"] for f in t["file"] if f["tahun"]),
                   "\n".join(f["path"] for f in t["file"]),
                   f"ss/{t['kode']}.png" if (SS_DIR / f"{t['kode']}.png").exists() else "belum ada",
                   judul_hal])
        for kol in (5, 6, 7):
            sel = ws.cell(row=i + 1, column=kol)
            if sel.value:
                sel.hyperlink = sel.value
                sel.font = Font(color="0563C1", underline="single")
    lebar = [5, 13, 60, 28, 70, 40, 40, 22, 9, 30, 70, 20, 60]
    for j, w in enumerate(lebar, 1):
        ws.column_dimensions[ws.cell(row=1, column=j).column_letter].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=c.column in (3, 11, 13))
    ws.freeze_panes = "C2"
    out = HERE / "Daftar_Link_Sumber_Dataset.xlsx"
    wb.save(out)
    print(f"Tersimpan: {out}")


if __name__ == "__main__":
    main()
