#!/usr/bin/env python3
"""
build_data.py  --  Ubah file Excel "REKAPITULASI HASIL LAB KIMIA AIR BULAN ..." menjadi data.js

Cara pakai:
    python build_data.py                # baca semua *.xlsx di folder ./excel  -> tulis data.js
    python build_data.py folder_excel   # baca semua *.xlsx di folder tertentu
    python build_data.py folder_excel hasil/data.js

Nama file harus mengandung nama bulan (JANUARI ... DESEMBER). Tahun default: 2026
(ubah lewat variabel TAHUN di bawah atau env LAB_YEAR).

Dependensi:  pip install openpyxl
"""
import sys, os, re, json, glob, datetime, warnings
from collections import Counter, defaultdict

import openpyxl

warnings.filterwarnings("ignore")

TAHUN = int(os.environ.get("LAB_YEAR", "2026"))
BULAN = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI",
         "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"]
MATRIKS = ["AIR PERMUKAAN", "AIR LIMBAH", "AIR BERSIH", "AIR MINUM",
           "AIR PILIHAN", "MAKANAN", "NARKOBA"]

# ---------------------------------------------------------------------------
# Baku mutu untuk penentuan MS/TMS otomatis (dipakai HANYA bila file Excel
# tidak punya baris "INTERPRESTASI HASIL"). Format: (batas_bawah, batas_atas)
# Ubah angka di sini kalau acuan baku mutu berubah.
# ---------------------------------------------------------------------------
ALIAS = {
    "residu terlarut (tds)": "TDS", "tds": "TDS",
    "residu tersuspensi (tss)": "TSS", "tss": "TSS",
    "kromium valensi 6 (cr 6+)": "Cr6", "kromium heksavalen (cr6+)": "Cr6",
    "besi (fe)": "Fe", "mangan (mn)": "Mn", "nitrit (no2)": "NO2",
    "fluorida (f)": "F", "aluminium (al)": "Al", "nitrat (no3)": "NO3",
    "cadmium (cd)": "Cd", "timbal (pb)": "Pb", "kekeruhan": "Kekeruhan",
    "warna": "Warna", "ph": "pH", "bod5": "BOD5", "cod": "COD",
    "nh3": "NH3", "minyak & lemak": "Minyak", "minyak dan lemak": "Minyak",
}
RULES = {
    "AIR MINUM": {"TDS": (None, 300), "Kekeruhan": (None, 3), "Warna": (None, 10),
                  "pH": (6.5, 8.5), "Fe": (None, 0.2), "Mn": (None, 0.1),
                  "NO2": (None, 3), "Cr6": (None, 0.01), "F": (None, 1.5),
                  "Al": (None, 0.2), "Cd": (None, 0.003), "Pb": (None, 0.01),
                  "NO3": (None, 20)},
    # Air Bersih: dikalibrasi terhadap kolom interpretasi asli (akurasi ~99%)
    "AIR BERSIH": {"TDS": (None, 300), "Kekeruhan": (None, 3), "Warna": (None, 10),
                   "pH": (6.5, 8.5), "Fe": (None, 0.2), "Mn": (None, 0.1),
                   "NO2": (None, 3), "NO3": (None, 20), "Cr6": (None, 0.01)},
    "AIR LIMBAH": {"TSS": (None, 30), "pH": (6.0, 9.0), "BOD5": (None, 30),
                   "COD": (None, 100), "NH3": (None, 10), "Minyak": (None, 5)},
    "AIR PERMUKAAN": {"TDS": (None, 1000), "TSS": (None, 50), "pH": (6.0, 9.0),
                      "Fe": (None, 0.3), "Mn": (None, 0.1), "BOD5": (None, 3),
                      "COD": (None, 25)},
    # Air Pilihan: dikalibrasi terhadap interpretasi asli (akurasi ~95%)
    "AIR PILIHAN": {"TDS": (None, 500), "TSS": (None, 50), "Kekeruhan": (None, 3),
                    "Warna": (None, 10), "pH": (6.5, 8.5), "Fe": (None, 0.2),
                    "Mn": (None, 0.1), "NO2": (None, 3), "NO3": (None, 20),
                    "NH3": (None, 1.5), "BOD5": (None, 30), "COD": (None, 100)},
}

# Urutan penting: pola yang lebih spesifik harus di atas.
LOKASI_RULES = [
    (r"\bPLG\b|PALEMBANG", "Palembang"),
    (r"MUBA\b|MUSI BANYUASIN", "Musi Banyuasin"),
    (r"BANYUASIN", "Banyuasin"),
    (r"OGAN ILIR|^OI$|INDRALAYA", "Ogan Ilir"),
    (r"OKUT|OKU TIMUR|MARTAPURA", "OKU Timur"),
    (r"OKUS|OKU SELATAN|MUARADUA", "OKU Selatan"),
    (r"BATURAJA|^OKU$|OGAN KOMERING ULU", "OKU"),
    (r"^OKI$|OGAN KOMERING ILIR|KAYUAGUNG", "OKI"),
    (r"MUARA ENIM|^ME$|TANJUNG ENIM|GUNUNG MEGANG", "Muara Enim"),
    (r"LAHAT", "Lahat"),
    (r"EMPAT LAWANG", "Empat Lawang"),
    (r"MURATARA|MUSI RAWAS UTARA", "Musi Rawas Utara"),
    (r"MUSI RAWAS|^MURA$", "Musi Rawas"),
    (r"LUBUK ?LINGGAU", "Lubuklinggau"),
    (r"PRABUMULIH", "Prabumulih"),
    (r"PAGAR ?ALAM", "Pagar Alam"),
    (r"^PALI$|PENUKAL", "PALI"),
]
# Lokasi di luar Sumatera Selatan (dikelompokkan sebagai "Luar Sumsel" di peta)
LUAR_SUMSEL = {"Pangkal Pinang", "Bangka Barat", "Bengkulu", "Lampung"}

META_KEYS = {
    "NO SPPC": "sppc", "LOKASI": "lok", "TANGGAL MASUK": "in",
    "TANGGAL KELUAR": "out", "STATUS": "stat", "NO SAMPEL": "samp",
    "INTERPRESTASI HASIL": "interp", "INTERPRETASI HASIL": "interp",
}


def key_of(name):
    """Nama parameter -> kunci aturan. 'Amoniak (NH3)' -> 'NH3'."""
    n = name.strip()
    a = ALIAS.get(n.lower())
    if a:
        return a
    m = re.search(r"\(([^)]+)\)\s*$", n)
    return m.group(1).replace(" ", "") if m else n


def norm_text(v):
    if v is None:
        return None
    s = str(v).strip()
    return s if s else None


def clean_loc(v):
    s = norm_text(v)
    if not s:
        return None
    up = re.sub(r"\s+", " ", s.upper())
    for pat, name in LOKASI_RULES:
        if re.search(pat, up):
            return name
    return s.title()


def clean_date(v):
    """Kembalikan 'YYYY-MM-DD' atau None bila tidak valid."""
    if v is None:
        return None
    if isinstance(v, datetime.datetime):
        d = v.date()
    elif isinstance(v, datetime.date):
        d = v
    elif isinstance(v, (int, float)):
        try:
            d = (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(v))).date()
        except Exception:
            return None
    else:
        s = str(v).strip()
        if not s:
            return None
        m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$", s)
        if not m:
            m2 = re.match(r"^(\d{1,2})/(\d{2})(\d{4})$", s)  # contoh salah ketik 14/102026
            m = m2
        if not m:
            m3 = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
            if m3:
                try:
                    d = datetime.date(int(m3[1]), int(m3[2]), int(m3[3]))
                except ValueError:
                    return None
                return d.isoformat() if 2000 <= d.year <= 2100 else None
            return None
        try:
            d = datetime.date(int(m[3]), int(m[2]), int(m[1]))
        except ValueError:
            return None
    return d.isoformat() if 2000 <= d.year <= 2100 else None


def norm_id(v):
    """Nomor sampel: 2192.0 -> '2192' ; 'U. 4331' -> 'U.4331'."""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    t = norm_text(v)
    if not t:
        return None
    return re.sub(r"^U\.\s+", "U.", t)


def resolve_dates(d_in, d_out):
    """Perbaiki salah ketik tahun; tandai yang tidak bisa diperbaiki.
    return (in, out, tat, dq)  dq: 0=ok, 1=diperbaiki otomatis, 2=tidak valid."""
    if not d_in or not d_out:
        return d_in, d_out, None, 0
    a = datetime.date.fromisoformat(d_in)
    b = datetime.date.fromisoformat(d_out)
    if 0 <= (b - a).days <= 60:
        return d_in, d_out, (b - a).days, 0

    def with_year(d, y):
        try:
            return d.replace(year=y)
        except ValueError:
            return None
    cands = []
    for ya in (TAHUN, a.year):
        for yb in (TAHUN, TAHUN + 1, b.year):
            a2, b2 = with_year(a, ya), with_year(b, yb)
            if a2 and b2:
                cands.append((a2, b2))
    for a2, b2 in cands:
        if 0 <= (b2 - a2).days <= 60 and (a2, b2) != (a, b):
            return a2.isoformat(), b2.isoformat(), (b2 - a2).days, 1
    return d_in, d_out, None, 2


def parse_numeric(v):
    """Angka dari sel hasil. '< 0,0035' -> 0 ; 'Negatif' -> None."""
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", ".")
    if s.startswith("<"):
        return 0.0
    m = re.search(r"[-+]?\d*\.?\d+", s)
    return float(m.group()) if m else None


def fmt_value(v):
    if isinstance(v, float):
        if v == int(v) and abs(v) < 1e12:
            return str(int(v))
        return ("%.6g" % v)
    return str(v).strip()


def month_from_filename(fn):
    up = os.path.basename(fn).upper()
    for i, b in enumerate(BULAN):
        if b in up:
            return i + 1
    return None


def read_sheet(ws, maxrow=120):
    return [list(r) for r in ws.iter_rows(min_row=1, max_row=maxrow, values_only=True)]


def parse_sheet(rows, matrix, period):
    # --- cari baris header 'NO SPPC' di kolom A-F
    meta = {}
    for ri, r in enumerate(rows):
        for ci in range(0, min(6, len(r))):
            t = norm_text(r[ci])
            if t:
                key = META_KEYS.get(re.sub(r"\s+", " ", t.upper()))
                if key and key not in meta:
                    meta[key] = (ri, ci)
    if "samp" not in meta or "sppc" not in meta:
        return []
    label_col = meta["samp"][1]
    first_data_col = label_col + 1
    meta_end = max(v[0] for v in meta.values())

    # --- batas akhir parameter
    end_row = len(rows)
    for ri in range(meta_end + 1, len(rows)):
        r = rows[ri]
        txt = " ".join(str(c) for c in r[:4] if c)
        if "Kepala Instalasi" in txt:
            end_row = ri
            break

    params = []
    for ri in range(meta_end + 1, end_row):
        r = rows[ri]
        name = norm_text(r[1]) if len(r) > 1 else None
        if not name:
            continue
        if re.match(r"^[A-C]\.\s", name):  # judul seksi "A. Fisika"
            continue
        analis = norm_text(r[2]) if len(r) > 2 else None
        params.append((ri, name, analis or "-"))

    maxcols = max(len(r) for r in rows)

    def cell(key, c):
        if key not in meta:
            return None
        row = rows[meta[key][0]]
        return row[c] if c < len(row) else None

    out = []
    cur = dict(sppc="-", lok=None, in_=None, out=None)
    for c in range(first_data_col, maxcols):
        samp = norm_id(cell("samp", c))
        if not samp:
            continue
        sp = norm_text(cell("sppc", c))
        if sp:
            # SPPC baru -> mulai kelompok baru
            sp = norm_id(cell("sppc", c))
            cur = dict(sppc=sp, lok=clean_loc(cell("lok", c)),
                       in_=clean_date(cell("in", c)), out=clean_date(cell("out", c)))
        else:
            l = clean_loc(cell("lok", c))
            if l: cur["lok"] = l
            d1 = clean_date(cell("in", c))
            if d1: cur["in_"] = d1
            d2 = clean_date(cell("out", c))
            if d2: cur["out"] = d2

        status = norm_text(cell("stat", c))
        status = "SUDAH" if (status and "SUDAH" in status.upper()) else "BELUM"

        interp = norm_text(cell("interp", c))
        interp = interp.upper() if interp else None
        if interp not in ("MS", "TMS"):
            interp = None

        # --- parameter
        pd_list, flagged = [], []
        rules = {k: v for k, v in RULES.get(matrix, {}).items()}
        positif = False
        for ri, name, analis in params:
            row = rows[ri]
            raw = row[c] if c < len(row) else None
            if raw is None or (isinstance(raw, str) and not raw.strip()):
                continue
            val = fmt_value(raw)
            oos = 0
            if matrix in ("MAKANAN", "NARKOBA"):
                if re.search(r"positif", val, re.I) and not re.search(r"negatif", val, re.I):
                    oos = 1
                    positif = True
            else:
                rule = rules.get(key_of(name))
                if rule:
                    num = parse_numeric(raw)
                    lo, hi = rule
                    if num is not None and not str(raw).strip().startswith("<"):
                        if (lo is not None and num < lo) or (hi is not None and num > hi):
                            oos = 1
            if oos:
                flagged.append(name)
            pd_list.append([name, val, analis, oos])

        auto = ("TMS" if flagged else "MS") if pd_list else "N/A"
        hasil = interp or auto
        hsrc = "file" if interp else "auto"

        # sel TOTAL di kolom akhir: nomor sampel angka murni, tanpa tanggal
        if samp.isdigit() and not cur["in_"] and not cur["out"]:
            continue

        d_in, d_out, tat, dq = resolve_dates(cur["in_"], cur["out"])

        out.append({
            "b": period, "m": matrix, "sppc": cur["sppc"], "s": samp,
            "l": cur["lok"] or "Tidak Diketahui",
            "in": d_in, "out": d_out, "tat": tat, "dq": dq,
            "dr": ([cur["in_"], cur["out"]] if dq else None), "st": status,
            "h": hasil, "hs": hsrc, "auto": auto, "fl": flagged, "pd": pd_list,
        })
    return out


def parse_workbook(path, month=None):
    month = month or month_from_filename(path)
    if not month:
        raise ValueError(f"Bulan tidak terdeteksi dari nama file: {path}")
    period = f"{TAHUN}-{month:02d}"
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    res = []
    for ws in wb.worksheets:
        name = ws.title.strip().upper()
        if name not in MATRIKS:
            continue
        res.extend(parse_sheet(read_sheet(ws), name, period))
    return res


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else "excel"
    outfile = sys.argv[2] if len(sys.argv) > 2 else "data.js"
    files = sorted(glob.glob(os.path.join(folder, "*.xlsx")) +
                   glob.glob(os.path.join(folder, "*.xlsm")))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    if not files:
        sys.exit(f"Tidak ada file .xlsx di folder '{folder}'")
    allrec = []
    for f in sorted(files, key=lambda x: month_from_filename(x) or 99):
        try:
            recs = parse_workbook(f)
        except Exception as e:
            print(f"[GAGAL] {os.path.basename(f)}: {e}")
            continue
        c = Counter(r["m"] for r in recs)
        print(f"[OK] {os.path.basename(f)[:55]:55s} {len(recs):4d} sampel  {dict(c)}")
        allrec.extend(recs)
    payload = {
        "generated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "year": TAHUN,
        # konfigurasi dipakai juga oleh fitur upload Excel di browser (index.html)
        "cfg": {"year": TAHUN, "rules": RULES, "alias": ALIAS,
                "loc": LOKASI_RULES, "luar": sorted(LUAR_SUMSEL)},
        "records": allrec,
    }
    with open(outfile, "w", encoding="utf-8") as fh:
        fh.write("window.LAB_DATA=" + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + ";")
    print(f"\nTotal {len(allrec)} sampel -> {outfile} ({os.path.getsize(outfile)/1024:.0f} KB)")


if __name__ == "__main__":
    main()
