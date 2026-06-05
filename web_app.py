import streamlit as st
import os
import re
import shutil
import requests
from datetime import datetime, timedelta, timezone
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PILImage, ExifTags
from io import BytesIO

# ==========================================
# HARUS PALING ATAS SEBELUM COMMAND ST LAIN
# ==========================================
st.set_page_config(page_title="QC Master Pro - LGJA", page_icon="🚀", layout="centered")

# ==========================================
# MODERN DARK MODE CSS
# ==========================================
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

    /* ===== GLOBAL ===== */
    html, body, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        background: #0a0a0f !important;
        color: #e0e0e6 !important;
        font-family: 'Space Grotesk', sans-serif !important;
    }

    [data-testid="stAppViewContainer"] > .main {
        background: #0a0a0f !important;
    }

    /* Grain overlay */
    [data-testid="stAppViewContainer"]::before {
        content: '';
        position: fixed;
        inset: 0;
        opacity: 0.025;
        pointer-events: none;
        z-index: 9999;
        background-image: url("data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='noise'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23noise)'/%3E%3C/svg%3E");
    }

    /* ===== SCROLLBAR ===== */
    ::-webkit-scrollbar { width: 6px; }
    ::-webkit-scrollbar-track { background: #0a0a0f; }
    ::-webkit-scrollbar-thumb { background: #2a2a3a; border-radius: 3px; }
    ::-webkit-scrollbar-thumb:hover { background: #3a3a5a; }

    /* ===== TITLE ===== */
    [data-testid="stMarkdownContainer"] h1 {
        font-family: 'Space Grotesk', sans-serif !important;
        font-weight: 700 !important;
        font-size: 2.4rem !important;
        background: linear-gradient(135deg, #a78bfa, #7c3aed, #c084fc);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        letter-spacing: -0.02em;
        padding-bottom: 0.3rem !important;
        border-bottom: 2px solid rgba(124, 58, 237, 0.2);
    }

    /* ===== CAPTION ===== */
    [data-testid="stCaption"] {
        color: #6b6b80 !important;
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 0.78rem !important;
        letter-spacing: 0.04em;
    }

    /* ===== SUBHEADERS ===== */
    [data-testid="stMarkdownContainer"] h3 {
        font-family: 'Space Grotesk', sans-serif !important;
        font-weight: 600 !important;
        font-size: 1.15rem !important;
        color: #c4b5fd !important;
        letter-spacing: 0.01em;
        padding-top: 0.5rem !important;
    }

    [data-testid="stMarkdownContainer"] h5 {
        font-family: 'Space Grotesk', sans-serif !important;
        font-weight: 500 !important;
        color: #8b8ba0 !important;
    }

    /* ===== DIVIDER ===== */
    hr {
        border: none !important;
        height: 1px !important;
        background: linear-gradient(90deg, transparent, #2a2a4a, transparent) !important;
        margin: 1.5rem 0 !important;
    }

    /* ===== INPUTS ===== */
    [data-testid="stTextInput"] input,
    [data-testid="stNumberInput"] input {
        background: #131320 !important;
        border: 1px solid #2a2a40 !important;
        border-radius: 10px !important;
        color: #e8e8f0 !important;
        font-family: 'Space Grotesk', sans-serif !important;
        padding: 0.7rem 1rem !important;
        transition: border-color 0.25s ease, box-shadow 0.25s ease !important;
    }

    [data-testid="stTextInput"] input:focus,
    [data-testid="stNumberInput"] input:focus {
        border-color: #7c3aed !important;
        box-shadow: 0 0 0 3px rgba(124, 58, 237, 0.15) !important;
        outline: none !important;
    }

    [data-testid="stTextInput"] input::placeholder {
        color: #4a4a60 !important;
        font-style: italic;
    }

    /* ===== LABELS ===== */
    [data-testid="stWidgetLabel"] label,
    label {
        color: #a0a0b8 !important;
        font-family: 'Space Grotesk', sans-serif !important;
        font-weight: 500 !important;
        font-size: 0.85rem !important;
        letter-spacing: 0.02em;
    }

    /* ===== SELECTBOX ===== */
    [data-testid="stSelectbox"] > div > div {
        background: #131320 !important;
        border: 1px solid #2a2a40 !important;
        border-radius: 10px !important;
        color: #e8e8f0 !important;
        transition: border-color 0.25s ease !important;
    }

    [data-testid="stSelectbox"] > div > div:hover {
        border-color: #7c3aed !important;
    }

    /* ===== RADIO BUTTONS ===== */
    [data-testid="stRadio"] > div {
        gap: 0.5rem !important;
    }

    [data-testid="stRadio"] label {
        background: #131320 !important;
        border: 1px solid #2a2a40 !important;
        border-radius: 8px !important;
        padding: 0.45rem 0.9rem !important;
        transition: all 0.2s ease !important;
    }

    [data-testid="stRadio"] label:hover {
        border-color: #7c3aed !important;
        background: #1a1a2e !important;
    }

    /* ===== FILE UPLOADER ===== */
    [data-testid="stFileUploader"] {
        border: 2px dashed #2a2a4a !important;
        border-radius: 14px !important;
        background: rgba(124, 58, 237, 0.03) !important;
        padding: 1rem !important;
        transition: border-color 0.3s ease, background 0.3s ease !important;
    }

    [data-testid="stFileUploader"]:hover {
        border-color: #7c3aed !important;
        background: rgba(124, 58, 237, 0.06) !important;
    }

    [data-testid="stFileUploader"] button {
        background: #1e1e35 !important;
        color: #c4b5fd !important;
        border: 1px solid #3a3a5a !important;
        border-radius: 8px !important;
        font-family: 'Space Grotesk', sans-serif !important;
        font-weight: 500 !important;
        transition: all 0.2s ease !important;
    }

    [data-testid="stFileUploader"] button:hover {
        background: #7c3aed !important;
        color: #fff !important;
        border-color: #7c3aed !important;
    }

    /* ===== PRIMARY BUTTONS ===== */
    [data-testid="stButton"] button[kind="primary"],
    [data-testid="stDownloadButton"] a,
    [data-testid="stDownloadButton"] button {
        background: linear-gradient(135deg, #7c3aed, #6d28d9) !important;
        color: #fff !important;
        border: none !important;
        border-radius: 12px !important;
        font-family: 'Space Grotesk', sans-serif !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        padding: 0.75rem 1.5rem !important;
        letter-spacing: 0.02em;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 20px rgba(124, 58, 237, 0.3) !important;
    }

    [data-testid="stButton"] button[kind="primary"]:hover,
    [data-testid="stDownloadButton"] a:hover,
    [data-testid="stDownloadButton"] button:hover {
        background: linear-gradient(135deg, #8b5cf6, #7c3aed) !important;
        box-shadow: 0 6px 30px rgba(124, 58, 237, 0.5) !important;
        transform: translateY(-1px);
    }

    /* ===== SECONDARY BUTTONS ===== */
    [data-testid="stButton"] button:not([kind="primary"]) {
        background: #1a1a2e !important;
        color: #a78bfa !important;
        border: 1px solid #2a2a4a !important;
        border-radius: 10px !important;
        font-family: 'Space Grotesk', sans-serif !important;
        font-weight: 500 !important;
        transition: all 0.2s ease !important;
    }

    [data-testid="stButton"] button:not([kind="primary"]):hover {
        background: #1e1e38 !important;
        border-color: #7c3aed !important;
        color: #c4b5fd !important;
    }

    /* ===== ALERTS ===== */
    [data-testid="stAlert"] {
        border-radius: 10px !important;
        border: none !important;
        font-family: 'Space Grotesk', sans-serif !important;
        font-size: 0.88rem !important;
    }

    [data-testid="stAlert"][kind="success"] {
        background: rgba(16, 185, 129, 0.1) !important;
        color: #6ee7b7 !important;
        border-left: 3px solid #10b981 !important;
    }

    [data-testid="stAlert"][kind="warning"] {
        background: rgba(245, 158, 11, 0.1) !important;
        color: #fcd34d !important;
        border-left: 3px solid #f59e0b !important;
    }

    [data-testid="stAlert"][kind="error"] {
        background: rgba(239, 68, 68, 0.1) !important;
        color: #fca5a5 !important;
        border-left: 3px solid #ef4444 !important;
    }

    [data-testid="stAlert"][kind="info"] {
        background: rgba(99, 102, 241, 0.08) !important;
        color: #a5b4fc !important;
        border-left: 3px solid #6366f1 !important;
    }

    /* ===== TOAST ===== */
    [data-testid="stToast"] {
        background: #1a1a2e !important;
        color: #a78bfa !important;
        border: 1px solid #2a2a4a !important;
        border-radius: 10px !important;
        font-family: 'Space Grotesk', sans-serif !important;
    }

    /* ===== SPINNER ===== */
    [data-testid="stSpinner"] {
        color: #a78bfa !important;
    }

    /* ===== COLUMNS GAP ===== */
    [data-testid="stHorizontalBlock"] {
        gap: 1.2rem !important;
    }

    /* ===== SIDEBAR ===== */
    [data-testid="stSidebar"] {
        background: #0d0d18 !important;
        border-right: 1px solid #1a1a2e !important;
    }

    /* ===== NUMBER INPUT STEPPERS ===== */
    [data-testid="stNumberInput"] button {
        background: #1e1e35 !important;
        color: #a78bfa !important;
        border: 1px solid #2a2a4a !important;
    }

    [data-testid="stNumberInput"] button:hover {
        background: #7c3aed !important;
        color: #fff !important;
    }

    /* ===== CONTAINER ===== */
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 2rem !important;
        max-width: 780px !important;
    }

    /* ===== ENTRANCE ANIMATION ===== */
    [data-testid="stVerticalBlock"] > div {
        animation: fadeSlideUp 0.5s ease forwards;
        opacity: 0;
    }

    @keyframes fadeSlideUp {
        from { opacity: 0; transform: translateY(16px); }
        to   { opacity: 1; transform: translateY(0); }
    }

    [data-testid="stVerticalBlock"] > div:nth-child(1)  { animation-delay: 0.00s; }
    [data-testid="stVerticalBlock"] > div:nth-child(2)  { animation-delay: 0.05s; }
    [data-testid="stVerticalBlock"] > div:nth-child(3)  { animation-delay: 0.10s; }
    [data-testid="stVerticalBlock"] > div:nth-child(4)  { animation-delay: 0.15s; }
    [data-testid="stVerticalBlock"] > div:nth-child(5)  { animation-delay: 0.20s; }
    [data-testid="stVerticalBlock"] > div:nth-child(6)  { animation-delay: 0.25s; }
    [data-testid="stVerticalBlock"] > div:nth-child(7)  { animation-delay: 0.30s; }
    [data-testid="stVerticalBlock"] > div:nth-child(8)  { animation-delay: 0.35s; }
    [data-testid="stVerticalBlock"] > div:nth-child(9)  { animation-delay: 0.40s; }
    [data-testid="stVerticalBlock"] > div:nth-child(10) { animation-delay: 0.45s; }
    [data-testid="stVerticalBlock"] > div:nth-child(11) { animation-delay: 0.50s; }
    [data-testid="stVerticalBlock"] > div:nth-child(12) { animation-delay: 0.55s; }
    [data-testid="stVerticalBlock"] > div:nth-child(13) { animation-delay: 0.60s; }
    [data-testid="stVerticalBlock"] > div:nth-child(14) { animation-delay: 0.65s; }
    [data-testid="stVerticalBlock"] > div:nth-child(15) { animation-delay: 0.70s; }

    /* ===== DOWNLOAD GLOW ===== */
    [data-testid="stDownloadButton"] a,
    [data-testid="stDownloadButton"] button {
        animation: pulseGlow 2s ease-in-out infinite;
    }

    @keyframes pulseGlow {
        0%, 100% { box-shadow: 0 4px 20px rgba(124, 58, 237, 0.3); }
        50%      { box-shadow: 0 4px 30px rgba(124, 58, 237, 0.6); }
    }

    /* ===== TABLE ===== */
    [data-testid="stDataFrame"] {
        border-radius: 10px;
        overflow: hidden;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# FUNGSI-FUNGSI UTAMA
# ==========================================

def log_traffic(user, toko, tgl_qc, layout):
    JSONBLOB_ID = "019e8740-72c4-7731-8328-0e2c67465233"
    API_URL = f"https://jsonblob.com/api/jsonBlob/{JSONBLOB_ID}"
    try:
        response = requests.get(API_URL)
        try:
            data = response.json()
        except:
            data = []
        if not isinstance(data, list):
            data = []
        tz_jkt = timezone(timedelta(hours=7))
        timestamp = datetime.now(tz_jkt).strftime("%d %B %y / %H:%M")
        data.append({
            "Nama Pengguna": user,
            "Nama Toko": toko,
            "Tanggal QC": tgl_qc,
            "Timestamp": timestamp,
            "Ukuran Layout": layout,
        })
        headers = {"Content-Type": "application/json", "Accept": "application/json"}
        put_response = requests.put(API_URL, json=data, headers=headers)
        if put_response.status_code in [200, 201]:
            st.toast("Traffic berhasil dicatat di server!")
        else:
            st.error(f"Gagal mencatat traffic! Status: {put_response.status_code}")
    except Exception as e:
        st.error(f"Sistem Traffic Error Koneksi: {e}")


def correct_orientation(img):
    """Memperbaiki rotasi gambar berdasarkan EXIF Orientation"""
    try:
        if hasattr(img, "_getexif") and img._getexif() is not None:
            exif = img._getexif()
            orientation = exif.get(ExifTags.Base.Orientation)
            if orientation == 3:
                img = img.rotate(180, expand=True)
            elif orientation == 6:
                img = img.rotate(270, expand=True)
            elif orientation == 8:
                img = img.rotate(90, expand=True)
    except:
        pass
    return img


def extract_datetime(filename, uploaded_file_obj):
    match = re.search(r"(\d{4}-\d{2}-\d{2}) at (\d{2}\.\d{2}\.\d{2})", filename)
    if match:
        try:
            return datetime.strptime(
                f"{match.group(1)} {match.group(2).replace('.', ':')}", "%Y-%m-%d %H:%M:%S"
            )
        except:
            pass
    return datetime.min


# ==========================================
# TAMPILAN ANTARMUKA WEB-APP
# ==========================================

st.title("QC Master Pro - LGJA")
st.caption("Advanced Dynamic Automation System | Build by AI & GJorma")
st.markdown("---")

# Wajib Isi Nama Pengguna
user_name = st.text_input(
    "Nama Pengguna (Wajib diisi untuk memproses)", placeholder="Masukkan nama Anda..."
)
if not user_name:
    st.warning("Silakan isi Nama Pengguna di atas agar sistem bisa digunakan.")
    st.stop()

# 1. Penamaan File Output
st.subheader("1. Informasi Lokasi QC (Untuk Nama File)")
col1, col2 = st.columns(2)
with col1:
    nama_toko = st.text_input("Nama Toko / Area", placeholder="Contoh: Batavia PIK")
with col2:
    tanggal_qc = st.text_input("Tanggal QC", placeholder="Contoh: 10 Mar 26")

st.markdown("---")

# 2. Upload Template Master
st.subheader("2. Upload Excel Template Master")
mode_template = st.radio("Pilih sumber template Excel:", ["Upload Manual", "File Preset (GitHub)"])

excel_file = None
if mode_template == "Upload Manual":
    excel_file = st.file_uploader("Pilih file Excel (.xlsx)", type=["xlsx"])
else:
    preset_pilihan = st.selectbox(
        "Pilih file preset:", ["QC_LGJA.xlsx", "QC_Sultan.xlsx", "QC_Vano.xlsx"]
    )
    if os.path.exists(preset_pilihan):
        with open(preset_pilihan, "rb") as f:
            excel_file = BytesIO(f.read())
    else:
        st.error(f"File '{preset_pilihan}' belum tersedia di folder server/GitHub Anda.")

selected_sheet = None
if excel_file:
    try:
        excel_file.seek(0)
        wb_scan = load_workbook(excel_file, read_only=True)
        sheet_names = wb_scan.sheetnames
        wb_scan.close()
        selected_sheet = st.selectbox("Target Sheet:", sheet_names)
    except Exception as e:
        st.error(f"Gagal membaca struktur Excel: {e}")

# Pemilihan Layout & Ukuran Gambar
st.markdown("##### Pengaturan Layout & Ukuran Gambar")
layout_option = st.selectbox("Pilih Opsi Layout:", ["LGJA", "Sultan", "Vano", "Custom"])

if layout_option == "LGJA":
    ROWS = [2, 4, 6, 8, 10, 12]
    COLS = list(range(1, 13))
    COL_W = 41
    ROW_H = 246
    IMAGE_WIDTH_CM = 6.4
    IMAGE_HEIGHT_CM = 8.30
elif layout_option in ["Sultan", "Vano"]:
    ROWS = [2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30]
    COLS = [1, 2, 3, 4, 5, 6]
    COL_W = 20.43
    ROW_H = 123.75
    IMAGE_WIDTH_CM = 3.2
    IMAGE_HEIGHT_CM = 4.10
else:  # Custom
    c_row, c_col = st.columns(2)
    r_in = c_row.text_input("Rows (pisahkan dengan koma)", "2,4,6,8,10,12")
    c_in = c_col.text_input("Columns (angka 1=A, 2=B, pisahkan dengan koma)", "1,2,3,4,5,6")
    try:
        ROWS = [int(x.strip()) for x in r_in.split(",")]
        COLS = [int(x.strip()) for x in c_in.split(",")]
    except:
        st.error("Format Rows/Columns salah! Gunakan angka dipisah koma.")
        st.stop()
    c_w, c_h = st.columns(2)
    COL_W = c_w.number_input("Col Width", value=20.43)
    ROW_H = c_h.number_input("Row Height", value=123.75)
    i_w, i_h = st.columns(2)
    IMAGE_WIDTH_CM = i_w.number_input("Image Width (cm)", value=3.2)
    IMAGE_HEIGHT_CM = i_h.number_input("Image Height (cm)", value=4.10)

st.markdown("---")

# 3. Upload Foto-Foto QC
st.subheader("3. Upload Foto QC Lapangan")
st.info("Tekan Tahan / Pilih Banyak Foto Sekaligus dari Galeri HP Anda.")

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

if st.button("Delete All Foto (Reset Uploader)"):
    st.session_state.uploader_key += 1
    st.rerun()

uploaded_photos = st.file_uploader(
    "Pilih semua foto sekaligus dari Galeri",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
    key=f"photos_{st.session_state.uploader_key}",
)

if uploaded_photos:
    st.success(f"Photo : {len(uploaded_photos)} (total foto dipilih)")
else:
    st.caption("Belum ada foto yang dipilih")

st.markdown("---")

# 4. Tombol Eksekusi
st.subheader("4. Eksekusi Proses Data")
if st.button("MULAI EXPORT DAN PROSES DATA", type="primary", use_container_width=True):
    if not nama_toko or not tanggal_qc:
        st.warning("Silakan isi Nama Toko dan Tanggal QC terlebih dahulu agar nama file rapi!")
    elif not excel_file:
        st.warning("Silakan upload atau pilih Excel Template Master terlebih dahulu!")
    elif not uploaded_photos:
        st.warning("Silakan pilih foto QC yang akan diproses!")
    elif not selected_sheet:
        st.error("Target sheet tidak valid.")
    else:
        with st.spinner("Sedang memproses dan mengompres foto... Mohon tunggu..."):
            try:
                log_traffic(user_name, nama_toko, tanggal_qc, layout_option)

                excel_file.seek(0)
                wb = load_workbook(excel_file)
                ws = wb[selected_sheet]

                for c in COLS:
                    ws.column_dimensions[chr(64 + c)].width = COL_W
                for r in ROWS:
                    ws.row_dimensions[r].height = ROW_H

                sorted_photos = sorted(
                    uploaded_photos,
                    key=lambda x: extract_datetime(x.name, x),
                    reverse=True,
                )

                all_cells = [f"{chr(64 + col)}{row}" for row in ROWS for col in COLS]
                success_count = 0

                temp_dir = "temp_web_photos"
                os.makedirs(temp_dir, exist_ok=True)

                for i in range(min(len(sorted_photos), len(all_cells))):
                    photo = sorted_photos[i]
                    temp_path = os.path.join(temp_dir, f"compressed_img_{i}.jpg")

                    with PILImage.open(photo) as img_pil:
                        img_pil = correct_orientation(img_pil)

                        if img_pil.mode in ("RGBA", "P"):
                            img_pil = img_pil.convert("RGB")

                        img_pil.thumbnail((1280, 1280), PILImage.Resampling.LANCZOS)
                        img_pil.save(
                            temp_path,
                            format="JPEG",
                            quality=82,
                            optimize=True,
                            subsampling=0,
                        )

                    img_excel = ExcelImage(temp_path)
                    img_excel.width = int(IMAGE_WIDTH_CM * 37.8)
                    img_excel.height = int(IMAGE_HEIGHT_CM * 37.8)

                    ws.add_image(img_excel, all_cells[i])
                    success_count += 1

                output = BytesIO()
                wb.save(output)
                wb.close()
                output.seek(0)

                shutil.rmtree(temp_dir, ignore_errors=True)

                final_filename = f"{nama_toko.strip()} {tanggal_qc.strip()}.xlsx"

                st.success(
                    f"Berhasil menyusun & mengompres {success_count} foto! File siap diunduh."
                )

                st.download_button(
                    label="DOWNLOAD FILE",
                    data=output,
                    file_name=final_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"Terjadi kesalahan sistem saat memproses Excel: {e}")
