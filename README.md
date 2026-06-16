# QC Master Pro - LGJA Edition

> Otomasi cerdas untuk menyusun foto QC lapangan ke dalam laporan Excel secara instan. Dibangun dengan Streamlit, bisa diakses dari browser HP maupun PC.

**Live App:** [https://eteazzi.streamlit.app/](https://eteazzi.streamlit.app/)

---

## Tentang

**QC Master Pro** adalah platform web multi-tool yang dirancang untuk mempercepat proses Quality Control (QC) lapangan. Dari upload foto sampai jadi PDF siap kirim, semua bisa dilakukan dari satu tempat.

**6 Tool Tersedia:**

| Tool | Fungsi |
|---|---|
| **QC Image Inserter** | Susun foto QC otomatis ke template Excel |
| **Image Compress** | Batch kompres foto untuk kurangi ukuran file |
| **Watermark** | Tambah watermark teks (lokasi, tanggal, nama toko) |
| **Batch Rename** | Rename ratusan foto sekaligus dengan format konsisten |
| **PDF Converter** | Ubah file Excel menjadi PDF siap kirim |
| **Traffic Log** | Pantau aktivitas dan statistik penggunaan sistem |

---

## Demo & Monitoring

| Link | URL |
|---|---|
| **Live App** | [https://eteazzi.streamlit.app/](https://eteazzi.streamlit.app/) |
| **Traffic Log** | [https://jsonblob.com/019e8740-72c4-7731-8328-0e2c67465233](https://jsonblob.com/019e8740-72c4-7731-8328-0e2c67465233) |

---

## Prasyarat

- **Python 3.9** atau lebih baru
- **pip** (Python package manager)
- Browser (Chrome, Safari, Firefox)
- Bisa diakses dari HP jika deploy di jaringan lokal

---

## Instalasi

### 1. Clone Repository

```bash
git clone https://github.com/LegendofGJA/QC-excel-inserter-exe.git
cd QC-excel-inserter-exe
```

### 2. Buat Virtual Environment

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Linux / macOS
source venv/bin/activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Jalankan Aplikasi

```bash
streamlit run app.py
```

Aplikasi akan terbuka di browser secara otomatis.

---

## Struktur Proyek

```
QC-excel-inserter-exe/
├── .streamlit/
│   └── config.toml               # Konfigurasi tema & server
├── app.py                         # Landing page (6 tools)
├── style.py                       # CSS theme & shared utilities
├── requirements.txt
├── README.md
├── presets/                       # Template Excel master
│   ├── QC_LGJA.xlsx              # Layout 6x12, foto 6.4x8.30 cm
│   ├── QC_Sultan.xlsx            # Layout 15x6, foto 3.2x4.10 cm
│   └── QC_Vano.xlsx              # Layout 15x6, foto 3.2x4.10 cm
└── pages/
    ├── 1_📸_QC_Image_Inserter.py  # Tool utama
    ├── 2_🗜️_Image_Compress.py     # Batch kompres
    ├── 3_💧_Watermark.py           # Tambah watermark
    ├── 4_✏️_Batch_Rename.py        # Rename massal
    ├── 5_📄_PDF_Converter.py       # Excel ke PDF
    └── 6_📊_Traffic.py            # Dashboard log
```

---

## Tool 1 — QC Image Inserter

Tool utama untuk menyusun foto QC ke dalam template Excel secara otomatis.

### Alur Kerja

```
Isi Nama → Info Lokasi + Template → Pilih Sheet → Upload Foto → PROSES → Download
```

### Langkah Penggunaan

**1. Isi Nama Pengguna**

Masukkan nama Anda di kolom yang tersedia. Nama ini akan digunakan untuk logging dan pengisian cell otomatis.

**2. Informasi Lokasi QC**

- **Nama Toko / Area**: Contoh `Batavia PIK`
- **Tanggal QC**: Contoh `10 Mar 26`

**3. Pilih Template Excel**

| Mode | Cara |
|---|---|
| **Upload Manual** | Upload file `.xlsx` dari perangkat |
| **File Preset** | Pilih dari dropdown (`QC_LGJA`, `QC_Sultan`, `QC_Vano`) |

Setelah template dipilih, pilih **Target Sheet** tempat foto akan dimasukkan.

**4. Atur Cell Auto-Write (Opsional)**

Mode **Auto** mengisi cell secara otomatis:

| Cell | Isi |
|---|---|
| **B6** | Nama Toko |
| **B7** | Tanggal QC (dikonversi ke bahasa Inggris) |
| **D6** | Nama Pengguna |

Mode **Manual**: Anda mengisi sendiri cell tersebut di Excel setelah download.

**5. Atur Layout**

| Layout | Rows | Kolom | Ukuran Foto |
|---|---|---|---|
| **LGJA** | 2, 4, 6, 8, 10, 12 | 1-12 | 6.4 x 8.30 cm |
| **Sultan** | 2, 4, 6, ..., 30 | 1-6 | 3.2 x 4.10 cm |
| **Vano** | 2, 4, 6, ..., 30 | 1-6 | 3.2 x 4.10 cm |
| **Custom** | Isi manual | Isi manual | Isi manual |

**6. Upload Foto**

- Pilih semua foto sekaligus dari galeri HP atau file manager
- Format didukung: `.jpg`, `.jpeg`, `.png`, `.webp`
- Foto dikompres otomatis ke JPEG 1280px sebelum masuk Excel
- Orientasi foto dikoreksi otomatis dari metadata EXIF

**7. Download**

Klik **MULAI EXPORT DAN PROSES DATA**, lalu download file Excel hasilnya.

---

## Tool 2 — Image Compress

Kompres batch foto untuk mengurangi ukuran file tanpa kehilangan kualitas yang signifikan.

### Pengaturan

| Parameter | Range | Default |
|---|---|---|
| **Kualitas JPEG** | 10 - 100 | 80 |
| **Max Dimensi (px)** | 320 - 4096 | 1280 |

### Output

- Statistik: jumlah foto, ukuran awal, ukuran setelah kompres, persen penghematan
- Tabel detail per file
- Download semua foto dalam satu file ZIP

---

## Tool 3 — Watermark

Tambahkan watermark teks pada foto untuk bukti otentik QC lapangan.

### Pengaturan

| Parameter | Opsi |
|---|---|
| **Teks** | 2 baris (Baris 1 wajib, Baris 2 opsional) |
| **Posisi** | Bawah Kanan/Kiri/Tengah, Tengah, Atas Kanan/Kiri/Tengah |
| **Ukuran Font** | Kecil (36px), Sedang (56px), Besar (80px), Sangat Besar (120px), Custom |
| **Warna** | Putih, Kuning, Merah, Hijau, Biru |
| **Transparansi** | 50 - 255 |

### Fitur

- Preview real-time di foto pertama sebelum proses semua
- Background semi-transparan di belakang teks agar tetap terbaca
- Output: ZIP berisi semua foto dengan watermark

---

## Tool 4 — Batch Rename

Rename ratusan foto sekaligus dengan format penamaan yang konsisten.

### Pengaturan

| Parameter | Contoh |
|---|---|
| **Prefix** | `QC_BataviaPIK` |
| **Angka Awal** | `1` |
| **Digit Padding** | 3 → `001`, `002`, `003` |
| **Format** | Pertahankan asli atau semua ke `.jpg` |

### Fitur

- Preview nama file secara real-time
- Atur urutan foto dengan tombol panah (naik/turun)
- Thumbnail foto ditampilkan di setiap baris
- Output: ZIP berisi semua foto yang sudah di-rename

---

## Tool 5 — PDF Converter

Ubah file Excel (`.xlsx`) menjadi PDF siap kirim ke klien atau atasan.

### Fitur

- Konversi tabel data Excel ke PDF Landscape A4
- Ekstrak gambar floating/over-cell dari sheet dan masukkan ke PDF
- Header tabel otomatis diulang di setiap halaman baru
- Penomoran halaman otomatis (Page X/Y)
- Font Unicode (DejaVu) dengan fallback Helvetica
- Konversi gambar RGBA/transparan ke JPEG dengan background putih

---

## Tool 6 — Traffic Log

Dashboard untuk memantau aktivitas penggunaan sistem.

**Live Traffic Log:** [https://jsonblob.com/019e8740-72c4-7731-8328-0e2c67465233](https://jsonblob.com/019e8740-72c4-7731-8328-0e2c67465233)

### Data yang Dicatat

| Field | Contoh |
|---|---|
| Nama Pengguna | `gabriel`, `Yoga QC`, `Yoga saputra` |
| Nama Toko | `Gandaria city` |
| Tanggal QC | `10 Mar 26` |
| Timestamp | `04 June 26 / 13:27` |
| Template | `QC_LGJA.xlsx, LGJA` |
| Opsi Layout | `LGJA`, `Sultan`, `Custom` |

### Statistik

- Total aktivitas
- Jumlah pengguna unik
- Jumlah toko terdata
- Aktivitas terakhir

---

## Konfigurasi Server

File `.streamlit/config.toml` sudah dikonfigurasi untuk:

| Config | Nilai | Keterangan |
|---|---|---|
| `maxUploadSize` | 500 MB | Mendukung upload ratusan foto sekaligus |
| `headless` | true | Cocok untuk deploy di server |
| `serverAddress` | `0.0.0.0` | Bisa diakses dari device lain di jaringan |
| `serverPort` | 8501 | Port default Streamlit |
| `primaryColor` | `#E5322D` | Merah konsisten dengan tema |

### Akses dari HP di Jaringan Lokal

1. Jalankan `streamlit run app.py` di PC/laptop
2. Cek IP lokal PC (contoh: `192.168.1.100`)
3. Buka browser HP, akses `http://192.168.1.100:8501`

---

## Konversi Tanggal Indonesia ke English

Saat mode Auto Cell Write aktif, tanggal QC otomatis dikonversi:

| Indonesia | English |
|---|---|
| Januari / Jan | January |
| Februari / Feb | February |
| Maret / Mar | March |
| April / Apr | April |
| Mei / May | May |
| Juni / Jun | June |
| Juli / Jul | July |
| Agustus / Agu / Aug | August |
| September / Sep | September |
| Oktober / Okt / Oct | October |
| November / Nov | November |
| Desember / Des / Dec | December |

---

## Panduan Upload dari HP

### Tips untuk Upload Stabil

- Gunakan **Gallery bawaan HP** atau **File Manager**
- **Google Photos** sering memutus koneksi saat upload banyak foto
- Jika muncul "reconnecting", tunggu beberapa detik lalu coba lagi
- Pilih semua foto sekaligus, jangan satu per satu
- Gunakan tombol **Hapus Semua Foto (Reset)** jika perlu mulai ulang

---

## Troubleshooting

| Masalah | Solusi |
|---|---|
| `ModuleNotFoundError` | Jalankan `pip install -r requirements.txt` |
| Upload foto gagal / reconnecting | Gunakan Gallery bawaan HP, bukan Google Photos |
| Sheet tidak ditemukan | Pastikan nama sheet di template sesuai dengan pilihan dropdown |
| Gambar tidak muncul di PDF | Pastikan file Excel mengandung embedded images |
| PDF error font | Pastikan font DejaVu terinstall di sistem (Linux) |
| Traffic log kosong | Periksa koneksi internet (traffic log memanggil API eksternal) |
| Aplikasi tidak bisa diakses dari HP | Pastikan `serverAddress = "0.0.0.0"` di config.toml dan firewall mengizinkan port 8501 |
| Upload timeout | Naikkan `websocketPingTimeout` di `.streamlit/config.toml` |

---

## Tech Stack

- **Streamlit** — Web framework untuk Python
- **openpyxl** — Manipulasi file Excel
- **Pillow (PIL)** — Pemrosesan dan kompresi gambar
- **fpdf2** — Generate PDF dari data Excel
- **pandas** — Data processing dan tabel
- **requests** — HTTP client untuk traffic logging

---

## Credits

**Build by AI & GJorma**

---

## License

MIT License — silakan digunakan dan dimodifikasi sesuai kebutuhan.
