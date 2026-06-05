import streamlit as st
import zipfile
import pandas as pd
from PIL import Image as PILImage
from io import BytesIO
from style import inject_css, inject_sidebar_brand, inject_footer, correct_orientation

st.set_page_config(page_title="Image Compress", page_icon="🗜️", layout="wide")
inject_css()
inject_sidebar_brand()

# PAGE HEADER
st.markdown(
    """
    <div class="page-head">
        <div class="page-head-icon">🗜️</div>
        <h2>Image Compress</h2>
        <p>Kompres batch foto untuk kurangi ukuran file secara signifikan</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# UPLOAD
uploaded_images = st.file_uploader(
    "Upload foto yang ingin dikompres",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
)

# SETTINGS
st.markdown("##### Pengaturan Kompresi")
col1, col2 = st.columns(2)
with col1:
    quality = st.slider("Kualitas JPEG", min_value=10, max_value=100, value=80, step=5)
with col2:
    max_dim = st.number_input("Max Dimensi (px)", min_value=320, max_value=4096, value=1280, step=64)

if uploaded_images:
    st.info(f"{len(uploaded_images)} foto dipilih. Klik tombol di bawah untuk mulai kompres.")

    if st.button("MULAI KOMPRES", type="primary", use_container_width=True):
        with st.spinner("Mengompres foto..."):
            zip_buffer = BytesIO()
            total_original = 0
            total_compressed = 0
            results = []

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for img_file in uploaded_images:
                    img_file.seek(0)
                    original_size = len(img_file.getvalue())
                    total_original += original_size

                    with PILImage.open(img_file) as img_pil:
                        img_pil = correct_orientation(img_pil)
                        if img_pil.mode in ("RGBA", "P"):
                            img_pil = img_pil.convert("RGB")
                        img_pil.thumbnail((max_dim, max_dim), PILImage.Resampling.LANCZOS)

                        buf = BytesIO()
                        img_pil.save(buf, format="JPEG", quality=quality, optimize=True, subsampling=0)
                        compressed_data = buf.getvalue()
                        total_compressed += len(compressed_data)

                        name = img_file.name.rsplit(".", 1)[0] + "_compressed.jpg"
                        zf.writestr(name, compressed_data)
                        results.append({
                            "File": img_file.name,
                            "Original (KB)": round(original_size / 1024, 1),
                            "Compressed (KB)": round(len(compressed_data) / 1024, 1),
                        })

            zip_buffer.seek(0)
            saving_pct = (
                round((1 - total_compressed / total_original) * 100, 1)
                if total_original > 0 else 0
            )

            # STATS
            st.markdown(
                f"""
                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="stat-card-value">{len(results)}</div>
                        <div class="stat-card-label">Foto Diproses</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-card-value">{round(total_original/1024/1024, 1)} MB</div>
                        <div class="stat-card-label">Ukuran Awal</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-card-value">{round(total_compressed/1024/1024, 1)} MB</div>
                        <div class="stat-card-label">Setelah Kompres</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-card-value">{saving_pct}%</div>
                        <div class="stat-card-label">Penghematan</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            # DETAIL TABLE
            df = pd.DataFrame(results)
            st.dataframe(df, use_container_width=True, hide_index=True)

            st.download_button(
                label="DOWNLOAD SEMUA (ZIP)",
                data=zip_buffer,
                file_name="compressed_images.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )
else:
    st.caption("Upload foto di atas untuk mulai kompres.")

inject_footer()
