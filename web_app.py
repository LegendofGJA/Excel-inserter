import streamlit as st
import os
import re
import shutil
import requests
from datetime import datetime, timedelta, timezone
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from PIL import Image as PILImage, ExifTags
from io import BytesIO

# ==========================================
# KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(
    page_title="QC Master Pro - LGJA",
    page_icon="🚀",
    layout="centered"
)

# ==========================================
# CUSTOM CSS — DARK AMOLED
# ==========================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;500;700&family=JetBrains+Mono:wght@400;500&display=swap');
html, body, [class*="css"] { font-family: 'Syne', sans-serif !important; }
.stApp { background-color: #0A0A0A !important; }
#MainMenu, header, footer { visibility: hidden; }
.qc-header {
    background: #111111; border: 1px solid #222222; border-radius: 12px;
    padding: 20px 24px; margin-bottom: 20px; display: flex; align-items: center; gap: 16px;
}
.qc-badge { background: #1D9E75; color: white; font-family: 'JetBrains Mono', monospace; font-size: 10px;
    font-weight: 500; letter-spacing: 0.08em; padding: 4px 10px; border-radius: 4px; text-transform: uppercase; }
.qc-title { font-size: 26px; font-weight: 700; color: #F0F0F0; margin: 0; letter-spacing: -0.02em; }
.qc-subtitle { font-size: 12px; color: #888888; font-family: 'JetBrains Mono', monospace; margin-top: 4px; }
.qc-logo { font-size: 36px; margin-left: auto; }
.step-header {
    display: flex; align-items: center; gap: 12px; padding: 14px 18px;
    background: #1A1A1A; border-bottom: 1px solid #222222; border-radius: 12px 12px 0 0;
}
.step-num {
    width: 28px; height: 28px; background: #1D9E75; color: white; border-radius: 6px;
    display: inline-flex; align-items: center; justify-content: center;
    font-size: 12px; font-weight: 700; font-family: 'JetBrains Mono', monospace;
}
.field-label {
    font-size: 11px; font-weight: 500; color: #AAAAAA; letter-spacing: 0.06em;
    text-transform: uppercase; font-family: 'JetBrains Mono', monospace; margin-bottom: 4px; display: block;
}
.photo-stat {
    background: #111111; border: 1px solid #222222; border-radius: 12px;
    padding: 16px 20px; margin-top: 8px;
}
.photo-stat-num { font-size: 32px; font-weight: 700; color: #1D9E75; font-family: 'JetBrains Mono', monospace; }
.photo-stat-label { font-size: 13px; color: #AAAAAA; font-family: 'JetBrains Mono', monospace; }
.block-container { padding-top: 20px !important; padding-bottom: 40px !important; max-width: 680px !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# FUNGSI TRAFFIC (UPDATED)
# ==========================================
def log_traffic(user, toko, tgl_qc, preset, target_sheet, layout):
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
            "Date QC & TS": f"{tgl_qc} & {timestamp}",
            "Template": f'File Preset "{preset}", Target Sheet "{target_sheet}", Layout "{layout}"'
        })
       
        headers = {'Content-Type': 'application/json', 'Accept': 'application/json'}
        put_response = requests.put(API_URL, json=data, headers=headers)
        if put_response.status_code in [200, 201]:
            st.toast("✅ Traffic berhasil dicatat di server!")
        else:
            st.error(f"⚠️ Gagal mencatat traffic! Status: {put_response.status_code}")
    except Exception as e:
        st.error(f"⚠️ Sistem Traffic Error Koneksi: {e}")

# ==========================================
# FUNGSI LAIN
# ==========================================
def correct_orientation(img):
    try:
        if hasattr(img, '_getexif') and img._getexif() is not None:
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
    match = re.search(r'(\d{4}-\d{2}-\d{2}) at (\d{2}\.\d{2}\.\d{2})', filename)
    if match:
        try:
            return datetime.strptime(f"{match.group(1)} {match.group(2).replace('.', ':')}", "%Y-%m-%d %H:%M:%S")
        except:
            pass
    return datetime.min

# ==========================================
# FUNGSI GAMBAR DI TENGAH CELL (SUDAH DIPERBAIKI)
# ==========================================
def add_image_to_center(ws, img_path, cell, img_width_px, img_height_px):
    img = ExcelImage(img_path)
    img.width = img_width_px
    img.height = img_height_px

    # Parsing cell (contoh: "A5" → col='A', row=5)
    col_letter = cell[0].upper()
    row_num = int(cell[1:])

    # Convert column letter to 0-based index
    col_idx = ord(col_letter) - ord('A')

    anchor = OneCellAnchor(
        _from=AnchorMarker(
            col=col_idx,           # Harus integer
            row=row_num - 1,       # openpyxl pakai 0-based row
            colOff=0,
            rowOff=0
        )
    )

    # Hitung ukuran cell untuk centering
    try:
        cell_width_px = (ws.column_dimensions[col_letter].width or 20) * 7.5
        cell_height_px = (ws.row_dimensions[row_num].height or 100) * 1.34
    except:
        cell_width_px = 300
        cell_height_px = 300

    x_offset = int((cell_width_px - img_width_px) / 2)
    y_offset = int((cell_height_px - img_height_px) / 2)

    anchor._from.colOff = x_offset * 9525   # EMU unit
    anchor._from.rowOff = y_offset * 9525

    img.anchor = anchor
    ws.add_image(img)

# ==========================================
# HEADER & STEPS
# ==========================================
st.markdown("""
<div class="qc-header">
    <div>
        <span class="qc-badge">v2.0 LGJA</span>
        <div class="qc-title">QC Master Pro</div>
        <div class="qc-subtitle">Advanced Dynamic Automation System · Build by AI & GJorma</div>
    </div>
    <div class="qc-logo">🚀</div>
</div>
""", unsafe_allow_html=True)

# STEP 0 — Identitas
st.markdown("""<div class="step-header"><span class="step-num">👤</span><span class="step-title">Identitas Pengguna</span></div>""", unsafe_allow_html=True)
st.markdown('<span class="field-label">Nama Pengguna (wajib diisi)</span>', unsafe_allow_html=True)
user_name = st.text_input("Nama Pengguna", placeholder="Masukkan nama Anda...", label_visibility="collapsed")
if not user_name:
    st.warning("⚠️ Silakan isi Nama Pengguna di atas agar sistem bisa digunakan.")
    st.stop()
st.markdown("---")

# STEP 1 — Informasi Lokasi
st.markdown("""<div class="step-header"><span class="step-num">01</span><span class="step-title">Informasi Lokasi QC — Untuk Nama File</span></div>""", unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    st.markdown('<span class="field-label">Nama Toko / Area</span>', unsafe_allow_html=True)
    nama_toko = st.text_input("Nama Toko", placeholder="Contoh: Batavia PIK", label_visibility="collapsed")
with col2:
    st.markdown('<span class="field-label">Tanggal QC</span>', unsafe_allow_html=True)
    tanggal_qc = st.text_input("Tanggal QC", placeholder="Contoh: 10 Mar 26", label_visibility="collapsed")
st.markdown("---")

# STEP 2 — Template Excel
st.markdown("""<div class="step-header"><span class="step-num">02</span><span class="step-title">Template Excel Master</span></div>""", unsafe_allow_html=True)
st.markdown('<span class="field-label">Sumber Template</span>', unsafe_allow_html=True)
mode_template = st.radio("Sumber Template", ["📤 Upload Manual", "📦 File Preset (GitHub)"], label_visibility="collapsed", horizontal=True)

excel_file = None
preset_pilihan = None

if "Upload" in mode_template:
    excel_file = st.file_uploader("Excel", type=["xlsx"], label_visibility="collapsed")
else:
    preset_pilihan = st.selectbox("Preset", ["QC_LGJA.xlsx", "QC_Sultan.xlsx", "QC_Vano.xlsx"], label_visibility="collapsed")
    if os.path.exists(preset_pilihan):
        with open(preset_pilihan, "rb") as f:
            excel_file = BytesIO(f.read())

selected_sheet = None
if excel_file:
    try:
        excel_file.seek(0)
        wb_scan = load_workbook(excel_file, read_only=True)
        sheet_names = wb_scan.sheetnames
        wb_scan.close()
        st.markdown('<span class="field-label">Target Sheet</span>', unsafe_allow_html=True)
        selected_sheet = st.selectbox("Sheet", sheet_names, label_visibility="collapsed")
    except Exception as e:
        st.error(f"Gagal membaca struktur Excel: {e}")
st.markdown("---")

# STEP 3 — Layout
st.markdown("""<div class="step-header"><span class="step-num">03</span><span class="step-title">Pengaturan Layout & Ukuran Gambar</span></div>""", unsafe_allow_html=True)
st.markdown('<span class="field-label">Pilih Opsi Layout</span>', unsafe_allow_html=True)
layout_option = st.selectbox("Layout", ["LGJA", "Sultan", "Vano", "Custom"], label_visibility="collapsed")

if layout_option == "LGJA":
    ROWS = [2, 4, 6, 8, 10, 12]
    COLS = list(range(1, 13))
    COL_W, ROW_H = 41, 246
    IMAGE_WIDTH_CM, IMAGE_HEIGHT_CM = 6.4, 8.30
elif layout_option in ["Sultan", "Vano"]:
    ROWS = [2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30]
    COLS = [1, 2, 3, 4, 5, 6]
    COL_W, ROW_H = 20.43, 123.75
    IMAGE_WIDTH_CM, IMAGE_HEIGHT_CM = 3.2, 4.10
else:
    col_r, col_c = st.columns(2)
    with col_r:
        r_in = st.text_input("Rows", "2,4,6,8,10,12", label_visibility="collapsed")
    with col_c:
        c_in = st.text_input("Cols", "1,2,3,4,5,6", label_visibility="collapsed")
    try:
        ROWS = [int(x.strip()) for x in r_in.split(",")]
        COLS = [int(x.strip()) for x in c_in.split(",")]
    except:
        st.error("Format Rows/Columns salah!")
        st.stop()
    col_w, col_h = st.columns(2)
    with col_w: COL_W = st.number_input("ColW", value=20.43, label_visibility="collapsed")
    with col_h: ROW_H = st.number_input("RowH", value=123.75, label_visibility="collapsed")
    col_iw, col_ih = st.columns(2)
    with col_iw: IMAGE_WIDTH_CM = st.number_input("ImgW", value=3.2, label_visibility="collapsed")
    with col_ih: IMAGE_HEIGHT_CM = st.number_input("ImgH", value=4.32, label_visibility="collapsed")
st.markdown("---")

# STEP 4 — Upload Foto
st.markdown("""<div class="step-header"><span class="step-num">04</span><span class="step-title">Upload Foto QC Lapangan</span></div>""", unsafe_allow_html=True)
st.info("💡 Tekan Tahan / Pilih Banyak Foto Sekaligus dari Galeri HP Anda.")
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if st.button("🗑️ Reset — Hapus Semua Foto"):
    st.session_state.uploader_key += 1
    st.rerun()

uploaded_photos = st.file_uploader(
    "Foto QC",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
    key=f"photos_{st.session_state.uploader_key}",
    label_visibility="collapsed"
)

if uploaded_photos:
    st.markdown(f"""
    <div class="photo-stat">
        <div><div class="photo-stat-num">{len(uploaded_photos)}</div>
        <div class="photo-stat-label">foto dipilih</div></div>
    </div>
    """, unsafe_allow_html=True)
st.markdown("---")

# ==========================================
# STEP 5 — EKSEKUSI
# ==========================================
st.markdown("""<div class="step-header"><span class="step-num">🚀</span><span class="step-title">Eksekusi — Proses & Export Data</span></div>""", unsafe_allow_html=True)

if st.button("⚡ MULAI EXPORT DAN PROSES DATA", type="primary", use_container_width=True):
    if not nama_toko or not tanggal_qc:
        st.warning("⚠️ Silakan isi Nama Toko dan Tanggal QC terlebih dahulu.")
    elif not excel_file:
        st.warning("⚠️ Silakan upload atau pilih Excel Template Master terlebih dahulu.")
    elif not uploaded_photos:
        st.warning("⚠️ Silakan pilih foto QC yang akan diproses.")
    elif not selected_sheet:
        st.error("⚠️ Target sheet tidak valid.")
    else:
        with st.spinner("⏳ Sedang memproses dan mengompres foto... Mohon tunggu..."):
            try:
                # Traffic Log
                preset_name = preset_pilihan if preset_pilihan else "Manual Upload"
                log_traffic(user_name, nama_toko, tanggal_qc, preset_name, selected_sheet, layout_option)

                excel_file.seek(0)
                wb = load_workbook(excel_file)
                ws = wb[selected_sheet]

                # Set column width & row height
                for c in COLS:
                    ws.column_dimensions[chr(64 + c)].width = COL_W
                for r in ROWS:
                    ws.row_dimensions[r].height = ROW_H

                sorted_photos = sorted(uploaded_photos, key=lambda x: extract_datetime(x.name, x), reverse=True)
                all_cells = [f"{chr(64 + col)}{row}" for row in ROWS for col in COLS]

                success_count = 0
                temp_dir = "temp_web_photos"
                os.makedirs(temp_dir, exist_ok=True)

                img_width_px = int(IMAGE_WIDTH_CM * 37.8)
                img_height_px = int(IMAGE_HEIGHT_CM * 37.8)

                for i in range(min(len(sorted_photos), len(all_cells))):
                    photo = sorted_photos[i]
                    temp_path = os.path.join(temp_dir, f"compressed_img_{i}.jpg")
                   
                    with PILImage.open(photo) as img_pil:
                        img_pil = correct_orientation(img_pil)
                        if img_pil.mode in ("RGBA", "P"):
                            img_pil = img_pil.convert("RGB")
                        img_pil.thumbnail((1280, 1280), PILImage.Resampling.LANCZOS)
                        img_pil.save(temp_path, format="JPEG", quality=82, optimize=True, subsampling=0)

                    # Tambah gambar di tengah cell
                    add_image_to_center(ws, temp_path, all_cells[i], img_width_px, img_height_px)
                    success_count += 1

                output = BytesIO()
                wb.save(output)
                wb.close()
                output.seek(0)
                shutil.rmtree(temp_dir, ignore_errors=True)

                final_filename = f"{nama_toko.strip()} {tanggal_qc.strip()}.xlsx"
                st.success(f"✅ Berhasil menyusun & mengompres **{success_count} foto**! Gambar sudah di tengah cell.")
                st.download_button(
                    label="📥 DOWNLOAD FILE HASIL",
                    data=output,
                    file_name=final_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")
