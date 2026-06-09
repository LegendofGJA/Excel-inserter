import streamlit as st
import zipfile
import os
from io import BytesIO
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

# ARROW REORDER
if uploaded_images:
    # Initialize or reset order when file count changes
    if ("rename_order" not in st.session_state
            or len(st.session_state.rename_order) != len(uploaded_images)):
        st.session_state.rename_order = list(range(len(uploaded_images)))

    st.markdown("##### Atur Urutan Foto")
    st.caption("Klik tombol panah untuk menggeser posisi foto")

    order = st.session_state.rename_order

    for display_i, file_idx in enumerate(order):
        fname = uploaded_images[file_idx].name
        col_num, col_thumb, col_name, col_up, col_down = st.columns(
            [0.5, 1, 3, 0.5, 0.5]
        )

        with col_num:
            st.markdown(
                f'<p style="text-align:center; font-weight:700; color:#E5322D; '
                f'font-size:0.9rem; margin-top:10px;">{display_i + 1}</p>',
                unsafe_allow_html=True,
            )

        with col_thumb:
            uploaded_images[file_idx].seek(0)
            st.image(uploaded_images[file_idx], width=80)

        with col_name:
            new_preview = f"{prefix.strip()}_{str(start_num + display_i).zfill(pad_zeros)}"
            ext_preview = os.path.splitext(fname)[1].lower() if file_format == "Pertahankan Format Asli" else ".jpg"
            st.markdown(
                f'<p style="margin-top:10px; font-size:0.85rem; color:#eaeaf2;">'
                f'{fname}</p>'
                f'<p style="font-size:0.72rem; color:#E5322D; font-family:JetBrains Mono,monospace;">'
                f'{new_preview}{ext_preview}</p>',
                unsafe_allow_html=True,
            )

        with col_up:
            if display_i > 0:
                if st.button("▲", key=f"up_{display_i}"):
                    order[display_i], order[display_i - 1] = order[display_i - 1], order[display_i]
                    st.session_state.rename_order = order
                    st.rerun()

        with col_down:
            if display_i < len(order) - 1:
                if st.button("▼", key=f"down_{display_i}"):
                    order[display_i], order[display_i + 1] = order[display_i + 1], order[display_i]
                    st.session_state.rename_order = order
                    st.rerun()

    st.markdown("---")
    st.info(f"{len(uploaded_images)} foto siap di-rename sesuai urutan di atas.")

    if st.button("RENAME & DOWNLOAD", type="primary", use_container_width=True):
        if not prefix.strip():
            st.warning("Isi Prefix Nama terlebih dahulu!")
        else:
            with st.spinner("Merename foto..."):
                zip_buffer = BytesIO()
                ext_force = ".jpg" if file_format == "Semua ke JPG" else None
                order = st.session_state.rename_order

                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for display_i, file_idx in enumerate(order):
                        img_file = uploaded_images[file_idx]
                        img_file.seek(0)
                        data = img_file.getvalue()

                        if ext_force is None:
                            original_ext = os.path.splitext(img_file.name)[1].lower()
                            if not original_ext:
                                original_ext = ".jpg"
                            ext_use = original_ext
                        else:
                            ext_use = ext_force

                        new_name = f"{prefix.strip()}_{str(start_num + display_i).zfill(pad_zeros)}{ext_use}"
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
