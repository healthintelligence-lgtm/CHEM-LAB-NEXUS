# Dasbor Laboratorium Kimia Air & Makanan

Isi folder:
- `index.html`  : dashboard (satu halaman, tanpa build)
- `data.js`     : data hasil olahan 11 file Excel (Jan–Nov 2026, 3.279 sampel)
- `build_data.py` : skrip pengubah Excel -> data.js
- `excel/`      : taruh file Excel bulanan di sini (jangan ikut di-upload ke GitHub)

## Cara update data tiap bulan

**Cara A (disarankan, hasil sama persis dengan yang sekarang)**
1. Taruh file Excel bulan baru di folder `excel/` (nama file harus memuat nama bulan, mis. `12__REKAPITULASI_..._DESEMBER.xlsx`).
2. `pip install openpyxl` (sekali saja), lalu jalankan `python build_data.py`.
3. Commit/upload `data.js` yang baru ke GitHub -> Vercel otomatis deploy ulang.

**Cara B (tanpa Python)**
1. Buka dashboard, klik **Tambah data bulan (.xlsx)**, pilih file bulan baru. Data langsung tampil.
2. Klik **Unduh data.js**, ganti file `data.js` di repositori, lalu deploy ulang.
   (Tanpa langkah 2, data tambahan hanya ada di browser itu sampai halaman dimuat ulang.)

Bila bulan yang sama diunggah lagi, datanya menggantikan yang lama.

## Deploy gratis ke Vercel
1. Buat akun di github.com dan vercel.com (login pakai GitHub).
2. Di GitHub: **New repository** -> **uploading an existing file** -> upload `index.html`, `data.js`, `build_data.py`, `README.md` -> Commit.
3. Di Vercel: **Add New > Project** -> pilih repositori tadi -> **Deploy** (Framework: Other, tanpa pengaturan lain).
4. Selesai. Tiap `data.js` diganti di GitHub, Vercel deploy otomatis.

Catatan: URL Vercel bisa dibuka siapa saja yang punya link. Karena memuat nomor sampel dan lokasi, pertimbangkan repositori private dan fitur proteksi password Vercel bila perlu.

## Cara dashboard menilai MS/TMS
- Bila file Excel punya baris **INTERPRESTASI HASIL** (mulai September), nilai itu dipakai apa adanya.
- Bila tidak ada (Jan–Agu), hasil diestimasi dari baku mutu di `RULES` dalam `build_data.py` dan ditandai **≈** di dashboard. Batas Air Bersih dan Air Pilihan sudah dikalibrasi terhadap interpretasi asli (cocok ±99% dan ±95%). Ubah angkanya di `RULES` bila acuan baku mutu berbeda, lalu jalankan ulang `build_data.py`.
- Makanan: TMS bila ada parameter Positif.
