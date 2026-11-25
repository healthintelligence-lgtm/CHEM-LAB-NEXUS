# Dasbor Laboratorium Kimia Air, Makanan & Udara

Isi folder:
- `index.html`  : dashboard (satu halaman, tanpa build)
- `data.js`     : data gabungan Jan–Nov 2026, 4.086 sampel
    - Lab Kimia Air & Makanan : 3.279 sampel (11 file bulanan)
    - Lab Udara               :   755 sampel (Draft_Hasil_Sampel_Udara.xlsx)
    - Lab Udara (Program)     :    52 sampel (Draft_Hasil_Sampel_Udara_Program.xlsx)
- `build_data.py` : skrip pengubah Excel -> data.js
- `apps-script/Code.gs` : skrip Google Apps Script agar dashboard membaca Google Sheets langsung (lihat bagian paling bawah)
- `excel/`        : file Excel bulanan Lab Kimia Air (jangan ikut di-upload ke GitHub)
- `excel_udara/`  : file Excel Lab Udara, satu file berisi sheet per bulan (Januari ... Desember).
                    File yang namanya memuat kata **PROGRAM** masuk sebagai "Udara (Program)".

Filter di dashboard: **Lab** -> Periode -> Matriks -> Kab/kota -> Analis. Pilihan matriks, lokasi
dan analis menyesuaikan lab yang dipilih.

## Cara update data tiap bulan

**Cara A (disarankan, hasil sama persis dengan yang sekarang)**
1. Taruh file Excel bulan baru di folder `excel/` (nama file harus memuat nama bulan, mis. `12__REKAPITULASI_..._DESEMBER.xlsx`).
   File Udara terbaru (yang sheet bulannya sudah bertambah) taruh di `excel_udara/`, timpa file lama.
2. `pip install openpyxl` (sekali saja), lalu jalankan `python build_data.py`.
   Bila salah satu folder kosong, data lab tersebut diambil dari `data.js` yang sudah ada.
3. Commit/upload `data.js` yang baru ke GitHub -> Vercel otomatis deploy ulang.

**Cara B (tanpa Python)**
1. Buka dashboard, klik **Tambah data (.xlsx)**, pilih file Kimia bulan baru atau file Udara. Data langsung tampil.
   (File Udara dikenali otomatis dari sheet bernama bulan; data lab+bulan yang sama diganti.)
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

## Lab Udara
- **Analis** per parameter mengikuti daftar penanggung jawab (`ANALIS_UDARA` di `build_data.py`):
  Suhu, Getaran, NO2, NOx, TSP, Opasitas = Idan · Kelembaban, O3, PM2,5, Laju ventilasi = M Firdaus ·
  Arah & kec. angin, Kebisingan, SO2, H2S, Pb, PM10 = Badri · Pencahayaan, CO, NH3, HC, Partikulat = Merry.
  ISBB dan CO2 belum punya penanggung jawab (tampil "-"). Ubah di `ANALIS_UDARA` bila ada perubahan.
- **Matriks** = seksi di file: Udara ambien, Udara indoor, Emisi tidak bergerak, Emisi bergerak, Parameter fisik.
- **MS/TMS** diambil apa adanya dari baris "Kesimpulan Sampel". Tidak ada estimasi otomatis untuk udara
  (baku mutu udara bergantung waktu pengukuran/peruntukan). Sampel yang sudah ada hasilnya tetapi
  kesimpulannya kosong/"Proses" ditampilkan sebagai **Proses**.
- **TAT** tidak dihitung untuk udara karena file hanya memuat tanggal diterima dan deadline (bukan tanggal keluar).
- **Lokasi** Jan–Jun kosong karena sheet bulan itu belum punya baris "Lokasi Sampling".
- Sel yang terbaca Excel sebagai tanggal (mis. Suhu diketik "31.3" jadi 31-Mar) dikembalikan ke angka
  dan diberi tanda "(cek)" di rincian sampel.

## Sambungkan ke Google Sheets (tanpa upload manual dan tanpa deploy ulang)

Alurnya: **Google Sheets (kamu isi) -> Apps Script (membaca & mengolah) -> dashboard di Vercel (menampilkan)**.
Setelah dipasang sekali, cukup isi data di Google Sheets. Dashboard memuat data terbaru otomatis
(paling lambat 30 menit, atau langsung dengan tombol **Segarkan data**). `data.js` tetap ada sebagai cadangan.

### A. Siapkan Google Sheets
1. Di Google Drive buat satu folder (mis. "Data Lab Dashboard"), lalu upload semua file Excel (Kimia bulanan dan Udara).
2. Ubah jadi Google Sheets: otomatis lewat langkah B3 di bawah, atau manual (klik kanan file -> Buka dengan -> Google Sheets -> File -> Simpan sebagai Google Sheets).
3. Aturan nama dan isi file (dikenali otomatis):
   - **Kimia**: nama file memuat nama bulan (JANUARI ... DESEMBER), sheet bernama AIR MINUM, AIR BERSIH, dst.
   - **Udara**: file yang punya sheet bernama bulan (Januari ... Desember). Nama file memuat **PROGRAM** = "Udara (Program)".
   - Bila nama file memuat tahun (mis. 2027) dipakai tahun itu, selain itu `TAHUN` di Code.gs.
4. Bulan baru: untuk Kimia, **File -> Buat salinan** dari bulan lalu, ganti nama (mis. ...DESEMBER), kosongkan datanya, simpan di folder yang sama. Untuk Udara, tambah sheet bulan di file yang sudah ada.

### B. Pasang Apps Script (sekali saja)
1. Buka script.google.com -> **Proyek baru**. Hapus isi `Code.gs`, tempel seluruh isi `apps-script/Code.gs`.
2. Ganti `ISI_ID_FOLDER_DRIVE` dengan ID folder (URL folder: `drive.google.com/drive/folders/<ID>`).
3. *Opsional (kalau file masih .xlsx)*: klik **Layanan (+)** -> tambah **Drive API**, pilih fungsi `konversiXlsx` -> **Jalankan**. Ulangi bila belum semua terkonversi.
4. Pilih fungsi `perbarui` -> **Jalankan**, izinkan akses (bila muncul "belum diverifikasi": Lanjutan -> buka proyek). Proses pertama 1-3 menit.
5. Pilih fungsi `pasangTrigger` -> **Jalankan** (membaca ulang otomatis tiap 30 menit; hanya file yang berubah yang diproses ulang).
6. **Terapkan -> Deployment baru** -> jenis **Aplikasi web** -> Jalankan sebagai: **Saya**, Yang memiliki akses: **Siapa saja** -> Terapkan -> salin URL berakhiran `/exec`.

### C. Sambungkan dashboard
1. Buka `index.html` dengan editor teks, cari `const API='';` lalu isi: `const API='https://script.google.com/macros/s/xxxx/exec';`
2. Upload `index.html` ke GitHub (timpa) -> Vercel deploy **sekali ini saja**. Selesai.

Catatan:
- Bila `Code.gs` diubah kemudian: **Terapkan -> Kelola deployment -> Edit -> Versi baru** (URL tidak berubah).
- Google Sheets tetap privat. Yang bisa diakses lewat URL `/exec` hanya JSON hasil olahan, jadi jangan bagikan URL itu sembarangan.
- Bila Google Sheets gagal dibaca, dashboard otomatis memakai `data.js` dan menampilkan pesannya di bawah judul.
- Pembaca sheet di `Code.gs` adalah salinan dari `index.html` / `build_data.py`. Bila aturan baku mutu di `build_data.py` diubah, tempel ulang bagian `CFG` di `Code.gs` dan naikkan `VERSI`.
