import streamlit as st
import zipfile
import os
from style import inject_css, inject_sidebar_brand, inject_footer

st.set_page_config(page_title="Batch Rename", page_icon="✏️", layout="wide")
inject_css()
inject_sidebar_brand()

st.markdown(
    """
    <div class="page-head">
        <div class="page-head-icon">✏️</div>
        <h2>Batch Rename</h2>
        <p>Rename ratusan foto sekaligus dengan format penamaan yang konsisten</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# SETTINGS
st.markdown("##### Pengaturan Penamaan")
col1, col2 = st.columns(2)

with col1:
    prefix = st.text_input("Prefix Nama", placeholder="Contoh: QC_BataviaPIK")
    start_num = st.number_input("Mulai dari angka", min_value=1, value=1, step=1)

with col2:
    pad_zeros = st.select_slider("Digit angka (padding)", options=[1, 2, 3, 4], value=3)
    file_format = st.selectbox("Format output:", ["Pertahankan Format Asli", "Semua ke JPG"])

# PREVIEW
if prefix.strip():
    example = f"{prefix.strip()}_{str(start_num).zfill(pad_zeros)}.jpg"
    st.markdown(
        f'<p style="color:#50506a; font-size:0.82rem;">'
        f'Preview: <span style="color:#E5322D; font-family:JetBrains Mono,monospace;">'
        f'{example}</span></p>',
        unsafe_allow_html=True,
    )

st.markdown("---")

# UPLOAD
uploaded_images = st.file_uploader(
    "Upload foto yang ingin di-rename",
    type=["jpg", "jpeg", "png", "webp", "heic"],
    accept_multiple_files=True,
)

if uploaded_images:
    st.info(f"{len(uploaded_images)} foto dipilih.")

    if st.button("RENAME & DOWNLOAD", type="primary", use_container_width=True):
        if not prefix.strip():
            st.warning("Isi Prefix Nama terlebih dahulu!")
        else:
            with st.spinner("Merename foto..."):
                zip_buffer = BytesIO()
                ext = ".jpg" if file_format == "Semua ke JPG" else None

                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for idx, img_file in enumerate(uploaded_images):
                        img_file.seek(0)
                        data = img_file.getvalue()

                        if ext is None:
                            original_ext = os.path.splitext(img_file.name)[1].lower()
                            if not original_ext:
                                original_ext = ".jpg"
                            ext_use = original_ext
                        else:
                            ext_use = ext

                        new_name = f"{prefix.strip()}_{str(start_num + idx).zfill(pad_zeros)}{ext_use}"
                        zf.writestr(new_name, data)

                zip_buffer.seek(0)
                st.success(f"{len(uploaded_images)} foto berhasil di-rename!")

                st.download_button(
                    label="DOWNLOAD SEMUA (ZIP)",
                    data=zip_buffer,
                    file_name=f"{prefix.strip()}_renamed.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True,
                )
else:
    st.caption("Upload foto di atas untuk mulai rename.")

inject_footer()
