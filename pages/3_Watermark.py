import streamlit as st
import zipfile
from PIL import Image as PILImage, ImageDraw, ImageFont
from io import BytesIO
from style import inject_css, inject_sidebar_brand, inject_footer, correct_orientation

st.set_page_config(page_title="Watermark", page_icon="💧", layout="wide")
inject_css()
inject_sidebar_brand()

st.markdown(
    """
    <div class="page-head">
        <div class="page-head-icon">💧</div>
        <h2>Photo Watermark</h2>
        <p>Tambah watermark teks di foto untuk bukti otentik QC lapangan</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# WATERMARK SETTINGS
st.markdown("##### Pengaturan Watermark")
col1, col2 = st.columns(2)

with col1:
    wm_text_line1 = st.text_input("Baris 1 (Utama)", placeholder="Contoh: QC - Batavia PIK")
    wm_text_line2 = st.text_input("Baris 2 (Opsional)", placeholder="Contoh: 10 Mar 2026")

with col2:
    wm_position = st.selectbox("Posisi Watermark:", [
        "Bawah Kanan", "Bawah Kiri", "Bawah Tengah",
        "Atas Kanan", "Atas Kiri", "Atas Tengah",
    ])
    wm_opacity = st.slider("Transparansi", 50, 255, 180, step=10)

col3, col4 = st.columns(2)
with col3:
    wm_font_size = st.slider("Ukuran Font", 16, 120, 48, step=4)
with col4:
    wm_color = st.selectbox("Warna Teks:", ["Putih", "Kuning", "Merah", "Hijau", "Biru"])

COLOR_MAP = {
    "Putih": (255, 255, 255),
    "Kuning": (255, 255, 0),
    "Merah": (255, 60, 60),
    "Hijau": (0, 255, 120),
    "Biru": (80, 160, 255),
}

st.markdown("---")

# UPLOAD
uploaded_images = st.file_uploader(
    "Upload foto untuk ditambahkan watermark",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
)


def add_watermark(img, line1, line2, position, color_rgb, opacity, font_size):
    """Tambahkan watermark multi-line ke gambar"""
    img = img.copy()
    if img.mode != "RGBA":
        img = img.convert("RGBA")

    txt_layer = PILImage.new("RGBA", img.size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(txt_layer)

    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", font_size)
    except Exception:
        try:
            font = ImageFont.truetype("arial.ttf", font_size)
        except Exception:
            font = ImageFont.load_default()

    # Hitung ukuran teks
    lines = [l for l in [line1, line2] if l.strip()]
    line_heights = []
    line_widths = []
    padding = int(font_size * 0.6)
    margin = int(font_size * 1.2)

    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=font)
        line_widths.append(bbox[2] - bbox[0])
        line_heights.append(bbox[3] - bbox[1])

    total_text_height = sum(line_heights) + (len(lines) - 1) * padding
    max_text_width = max(line_widths) if line_widths else 0

    w, h = img.size

    # Hitung posisi
    positions_map = {
        "Bawah Kanan":  (w - max_text_width - margin, h - total_text_height - margin),
        "Bawah Kiri":   (margin, h - total_text_height - margin),
        "Bawah Tengah": ((w - max_text_width) // 2, h - total_text_height - margin),
        "Atas Kanan":   (w - max_text_width - margin, margin),
        "Atas Kiri":    (margin, margin),
        "Atas Tengah":  ((w - max_text_width) // 2, margin),
    }
    x, y = positions_map.get(position, positions_map["Bawah Kanan"])

    # Background bar
    bg_padding = int(font_size * 0.4)
    bg_coords = [
        x - bg_padding,
        y - bg_padding,
        x + max_text_width + bg_padding,
        y + total_text_height + bg_padding,
    ]
    draw.rectangle(bg_coords, fill=(0, 0, 0, int(opacity * 0.6)))

    # Gambar teks
    current_y = y
    for i, line in enumerate(lines):
        draw.text((x, current_y), line, font=font, fill=(*color_rgb, opacity))
        current_y += line_heights[i] + padding

    return PILImage.alpha_composite(img, txt_layer).convert("RGB")


if uploaded_images:
    st.info(f"{len(uploaded_images)} foto dipilih.")

    if st.button("TAMBAHKAN WATERMARK", type="primary", use_container_width=True):
        if not wm_text_line1.strip():
            st.warning("Isi minimal Baris 1 watermark!")
        else:
            with st.spinner("Menambahkan watermark..."):
                zip_buffer = BytesIO()
                color_rgb = COLOR_MAP.get(wm_color, (255, 255, 255))

                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for img_file in uploaded_images:
                        img_file.seek(0)
                        with PILImage.open(img_file) as img_pil:
                            img_pil = correct_orientation(img_pil)
                            result = add_watermark(
                                img_pil, wm_text_line1, wm_text_line2,
                                wm_position, color_rgb, wm_opacity, wm_font_size,
                            )
                            buf = BytesIO()
                            result.save(buf, format="JPEG", quality=90, optimize=True)
                            name = img_file.name.rsplit(".", 1)[0] + "_watermarked.jpg"
                            zf.writestr(name, buf.getvalue())

                zip_buffer.seek(0)
                st.success(f"Watermark berhasil ditambahkan ke {len(uploaded_images)} foto!")

                st.download_button(
                    label="DOWNLOAD SEMUA (ZIP)",
                    data=zip_buffer,
                    file_name="watermarked_photos.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True,
                )
else:
    st.caption("Upload foto di atas untuk mulai menambahkan watermark.")

inject_footer()
