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
LUAR_SUMSEL = {"Pangkal Pinang", "Bangka Barat", "Bangka Tengah", "Bengkulu", "Lampung",
               "Lampung Timur", "Metro"}

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
            try:   # format Amerika m/d/yyyy, mis. 2/18/2026
                d = datetime.date(int(m[3]), int(m[1]), int(m[2]))
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
            if isinstance(raw, (datetime.datetime, datetime.date)):
                # hasil "19.2" yang terbaca Excel sebagai tanggal: jangan dinilai sebagai angka
                pd_list.append([name, f"{raw.day}.{raw.month} (cek)", analis, 0])
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


# ===========================================================================
#  LAB UDARA
#  File Excel udara berisi satu sheet per bulan (Januari ... Desember).
#  Taruh di folder ./excel_udara . File yang namanya memuat kata "PROGRAM"
#  dicatat sebagai sumber "Udara - Program", selain itu "Udara".
# ===========================================================================
BULAN_ID = {b.lower(): i + 1 for i, b in enumerate(BULAN)}
SEKSI_UDARA = {
    "UDARA AMBIENT": "UDARA AMBIEN", "UDARA AMBIEN": "UDARA AMBIEN",
    "UDARA INDOOR": "UDARA INDOOR",
    "UDARA EMISI TIDAK BERGERAK": "EMISI TIDAK BERGERAK",
    "UDARA EMISI BERGERAK": "EMISI BERGERAK",
    "PARAMETER FISIK": "PARAMETER FISIK",
}
MATRIKS_UDARA = ["UDARA AMBIEN", "UDARA INDOOR", "EMISI TIDAK BERGERAK",
                 "EMISI BERGERAK", "PARAMETER FISIK"]
# Penanggung jawab (analis) per parameter Lab Udara.
# Kunci = nama parameter (atau rumus di dalam kurung) huruf kecil tanpa spasi/tanda baca.
ANALIS_UDARA = {
    # Fisika
    "suhu": "Idan", "kelembaban": "M Firdaus", "arahangin": "Badri",
    "kecangin": "Badri", "kecepatanangin": "Badri", "kebisingan": "Badri",
    "pencahayaan": "Merry", "getaran": "Idan", "kecepatanlajualir": "M Firdaus",
    "lajuventilasi": "M Firdaus", "tekananudara": "M Firdaus",
    # Kimia
    "h2s": "Badri", "nh3": "Merry", "so2": "Badri", "co": "Merry", "no2": "Idan",
    "nox": "Idan", "tsp": "Idan", "pb": "Badri", "hc": "Merry", "hcco": "Merry",
    "pm10": "Badri", "pm25": "M Firdaus", "o3": "M Firdaus", "opasitas": "Idan",
    "partikulat": "Merry",
    # ISBB, iklim kerja, CO2: belum ada penanggung jawab di daftar -> "-"
}
META_UDARA = [
    ("LOKASI", "lok"), ("TANGGAL PENGAMBILAN", "tk"), ("TANGGAL DITERIMA", "in"),
    ("DEADLINE", "dl"), ("STATUS VERIFIKASI", "st"), ("TANGGAL VERIFIKASI", "st"),
    ("KESIMPULAN", "ks"),
]


def ud_key(name):
    n = name.strip().lower()
    m = re.search(r"\(([^)]+)\)\s*$", n)
    if m:
        n = m.group(1)
    return re.sub(r"[^a-z0-9]", "", n)


def ud_analis(name):
    return ANALIS_UDARA.get(ud_key(name), "-")


def id_date(v):
    """Tanggal Indonesia -> 'YYYY-MM-DD'. '19-21 Agustus 2026' -> tanggal terakhir (21)."""
    if v is None or isinstance(v, (datetime.datetime, datetime.date, int, float)):
        return clean_date(v)
    s = re.sub(r"\s+", " ", str(v).strip().lower())
    m = re.search(r"(\d{1,2}) ?([a-z]+) ?(\d{4})$", s)
    if m and m[2] in BULAN_ID:
        try:
            d = datetime.date(int(m[3]), BULAN_ID[m[2]], int(m[1]))
            return d.isoformat()
        except ValueError:
            return None
    return clean_date(v)


def ud_value(raw):
    """Sel hasil yang terbaca Excel sebagai tanggal (salah ketik/format) -> kembalikan ke angka.
    30/01/1900 12:00 -> 30.5 ; 31/03/2026 (ketik '31.3') -> '31.3 (cek)' ; format yyyy.m -> '2114.9 (cek)'"""
    if isinstance(raw, datetime.datetime):
        if raw.year == 1900:
            base = datetime.datetime(1899, 12, 31) if raw < datetime.datetime(1900, 3, 1) \
                else datetime.datetime(1899, 12, 30)
            return fmt_value(round((raw - base).total_seconds() / 86400, 4))
        if raw.day == 1 and raw.year != TAHUN:
            return f"{raw.year}.{raw.month} (cek)"
        return f"{raw.day}.{raw.month} (cek)"
    return fmt_value(raw)


def ud_id(v):
    """No sampel/No Lab yang terformat tanggal 'yyyy.m' (mis. 6176.2) -> '6176.2'."""
    if isinstance(v, datetime.datetime):
        return f"{v.year}.{v.month}" if v.day == 1 else v.date().isoformat()
    return norm_id(v)


def parse_udara_sheet(rows, period, lab):
    meta = {}
    for ri, r in enumerate(rows[:15]):
        t = norm_text(r[0]) if r else None
        if not t:
            continue
        up = re.sub(r"\s+", " ", t.upper())
        for pref, key in META_UDARA:
            if up.startswith(pref) and key not in meta:
                meta[key] = ri
    if not meta:
        return []
    samp_row = min(meta.values()) - 1
    sppc_row = samp_row - 1
    if sppc_row < 0:
        return []

    # parameter per seksi
    params, sec, ket_row = [], None, None
    for ri in range(samp_row + 1, len(rows)):
        r = rows[ri]
        a, b, c = (norm_text(r[i]) if len(r) > i else None for i in range(3))
        if a and "KETERANGAN" in a.upper():
            ket_row = ri
            break
        if a and b and len(a) <= 2:
            sec = SEKSI_UDARA.get(re.sub(r"\s+", " ", b.upper()))
            continue
        if sec and c:
            params.append((ri, sec, c, ud_analis(c)))

    maxcols = max(len(r) for r in rows)

    def cell(ri, c):
        if ri is None or ri >= len(rows):
            return None
        row = rows[ri]
        return row[c] if c < len(row) else None

    def m_(key, c):
        return cell(meta.get(key), c)

    # --- kelompokkan kolom per No Lab (SPPC); info kelompok diambil dari kolom mana pun
    groups, g = [], None
    for c in range(3, maxcols):
        samp = ud_id(cell(samp_row, c))
        sp = ud_id(cell(sppc_row, c))
        if (samp and samp.startswith("#")) or (sp and sp.startswith("#")):
            continue
        if sp or g is None:
            g = dict(sppc=sp or "-", cols=[], lok=None, tk=None, in_=None, dl=None, st=None)
            groups.append(g)
        if not samp:
            continue
        g["cols"].append((c, samp))
        for k, raw in (("lok", m_("lok", c)), ("tk", m_("tk", c)), ("in_", m_("in", c)),
                       ("dl", m_("dl", c)), ("st", m_("st", c))):
            if g[k] is None and norm_text(raw):
                g[k] = raw

    out = []
    for g in groups:
        has_info = any(g[k] for k in ("lok", "tk", "in_", "st"))
        if not has_info:      # kolom TOTAL / rekap di ujung kanan
            continue
        g_lok, g_in, g_dl = clean_loc(g["lok"]), id_date(g["in_"]), id_date(g["dl"])
        g_tk = g["tk"]
        if isinstance(g_tk, (datetime.datetime, datetime.date)):
            g_tk = clean_date(g_tk)
        else:
            g_tk = norm_text(g_tk)
        for c, samp in g["cols"]:
            cur = dict(sppc=g["sppc"], lok=clean_loc(m_("lok", c)) or g_lok, tk=g_tk,
                       in_=g_in, dl=g_dl, st=norm_text(m_("st", c)) or norm_text(g["st"]))
            pd_list, secs = [], Counter()
            for ri, s, name, an in params:
                raw = cell(ri, c)
                if raw is None or (isinstance(raw, str) and not raw.strip()):
                    continue
                secs[s] += 1
                pd_list.append([name, ud_value(raw), an, s])
            matrix = secs.most_common(1)[0][0] if secs else "UDARA (BELUM ADA HASIL)"
            pd_out = [[name if s == matrix else f"{name} [{s.title()}]", val, an, 0]
                      for name, val, an, s in pd_list]

            ks = norm_text(m_("ks", c))
            ks = ks.upper() if ks else None
            if ks not in ("MS", "TMS"):
                ks = None
            status = "SUDAH" if "SUDAH" in (cur["st"] or "").upper() else "BELUM"
            ket = " / ".join(x for x in (norm_text(cell(ket_row, c)),
                                         norm_text(cell(ket_row + 1 if ket_row else None, c))) if x)
            out.append({
                "lab": lab, "b": period, "m": matrix, "sppc": cur["sppc"], "s": samp,
                "l": cur["lok"] or "Tidak Diketahui",
                "in": cur["in_"], "out": None, "tat": None, "dq": 0, "dr": None,
                "dl": cur["dl"], "tk": cur["tk"], "ket": ket or None, "st": status,
                "h": ks or "N/A", "hs": "file" if ks else "auto", "auto": "N/A",
                "fl": [], "pd": pd_out,
            })
    return out


def parse_udara_workbook(path):
    lab = "UDARA-P" if "PROGRAM" in os.path.basename(path).upper() else "UDARA"
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    res = []
    for ws in wb.worksheets:
        mi = BULAN_ID.get(ws.title.strip().lower())
        if not mi:
            continue
        rows = [list(r) for r in ws.iter_rows(min_row=1, max_row=80, values_only=True)]
        res.extend(parse_udara_sheet(rows, f"{TAHUN}-{mi:02d}", lab))
    return res


LAB_NAMA = {"KIMIA": "Kimia Air & Makanan", "UDARA": "Udara",
            "UDARA-P": "Udara (Program)"}


def load_existing(path):
    """Ambil records dari data.js yang sudah ada (dipakai bila folder Excel kosong)."""
    if not os.path.exists(path):
        return []
    s = open(path, encoding="utf-8").read().strip()
    s = s[s.index("=") + 1:].rstrip(";")
    recs = json.loads(s).get("records", [])
    for r in recs:
        r.setdefault("lab", "KIMIA")
    return recs


def main():
    """python build_data.py [folder_kimia] [data.js] [folder_udara]"""
    folder = sys.argv[1] if len(sys.argv) > 1 else "excel"
    outfile = sys.argv[2] if len(sys.argv) > 2 else "data.js"
    folder_ud = sys.argv[3] if len(sys.argv) > 3 else "excel_udara"
    lama = load_existing(outfile)

    def xl(d):
        f = sorted(glob.glob(os.path.join(d, "*.xlsx")) + glob.glob(os.path.join(d, "*.xlsm")))
        return [x for x in f if not os.path.basename(x).startswith("~$")]

    allrec = []
    # ---------- Lab Kimia Air & Makanan
    files = xl(folder)
    if files:
        for f in sorted(files, key=lambda x: month_from_filename(x) or 99):
            try:
                recs = parse_workbook(f)
            except Exception as e:
                print(f"[GAGAL] {os.path.basename(f)}: {e}")
                continue
            for r in recs:
                r["lab"] = "KIMIA"
            c = Counter(r["m"] for r in recs)
            print(f"[OK] {os.path.basename(f)[:55]:55s} {len(recs):4d} sampel  {dict(c)}")
            allrec.extend(recs)
    else:
        keep = [r for r in lama if r.get("lab", "KIMIA") == "KIMIA"]
        print(f"[INFO] folder '{folder}' kosong -> {len(keep)} sampel Kimia Air diambil dari {outfile} lama")
        allrec.extend(keep)

    # ---------- Lab Udara
    files = xl(folder_ud)
    if files:
        for f in files:
            try:
                recs = parse_udara_workbook(f)
            except Exception as e:
                print(f"[GAGAL] {os.path.basename(f)}: {e}")
                continue
            c = Counter(r["m"] for r in recs)
            print(f"[OK] {os.path.basename(f)[:55]:55s} {len(recs):4d} sampel  {dict(c)}")
            allrec.extend(recs)
    else:
        keep = [r for r in lama if r.get("lab", "KIMIA").startswith("UDARA")]
        print(f"[INFO] folder '{folder_ud}' kosong -> {len(keep)} sampel Udara diambil dari {outfile} lama")
        allrec.extend(keep)

    if not allrec:
        sys.exit("Tidak ada data. Taruh file Excel di folder 'excel' dan/atau 'excel_udara'.")
    allrec.sort(key=lambda r: (r["b"], r.get("lab", "")))
    payload = {
        "generated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "year": TAHUN,
        # konfigurasi dipakai juga oleh fitur upload Excel di browser (index.html)
        "cfg": {"year": TAHUN, "rules": RULES, "alias": ALIAS,
                "loc": LOKASI_RULES, "luar": sorted(LUAR_SUMSEL),
                "labs": LAB_NAMA, "mxu": MATRIKS_UDARA, "seksi": SEKSI_UDARA,
                "analisU": ANALIS_UDARA},
        "records": allrec,
    }
    with open(outfile, "w", encoding="utf-8") as fh:
        fh.write("window.LAB_DATA=" + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + ";")
    c = Counter(LAB_NAMA.get(r.get("lab"), r.get("lab")) for r in allrec)
    print(f"\nTotal {len(allrec)} sampel {dict(c)} -> {outfile} ({os.path.getsize(outfile)/1024:.0f} KB)")


if __name__ == "__main__":
    main()
