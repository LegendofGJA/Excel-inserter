import streamlit as st
from style import inject_css, inject_sidebar_brand, inject_footer

st.set_page_config(page_title="QC Master Pro", page_icon="📸", layout="wide")
inject_css()
inject_sidebar_brand()

# HERO
st.markdown(
    """
    <section class="ilp-hero">
        <div class="ilp-hero-icon">📋</div>
        <h1>QC Master Pro</h1>
        <p>Otomasi cerdas untuk menyusun foto QC lapangan ke dalam laporan Excel secara instan.</p>
    </section>
    """,
    unsafe_allow_html=True,
)

# TOOLS DEFINITION
tools = [
    ("📸", "QC Image Inserter", "Upload foto QC dan susun otomatis ke template Excel master", "TOOL UTAMA", "pages/1_QC_Image_Inserter.py"),
    ("🗜️", "Image Compress", "Kompres batch foto untuk kurangi ukuran file tanpa hilang kualitas", "UTILITAS", "pages/2_Image_Compress.py"),
    ("💧", "Watermark", "Tambah watermark teks di foto seperti lokasi, tanggal, dan nama toko", "UTILITAS", "pages/3_Watermark.py"),
    ("✏️", "Batch Rename", "Rename ratusan foto sekaligus dengan format penamaan yang konsisten", "UTILITAS", "pages/4_Batch_Rename.py"),
    ("📄", "PDF Converter", "Ubah file Excel hasil export menjadi PDF siap kirim ke klien", "KONVERSI", "pages/5_PDF_Converter.py"),
    ("📊", "Traffic Log", "Pantau aktivitas pengguna dan statistik penggunaan sistem", "ANALITIK", "pages/6_Traffic.py"),
]

# ROW 1
row1 = st.columns(3)
for col, (icon, title, desc, tag, page) in zip(row1, tools[:3]):
    with col:
        st.markdown(
            f"""<div class="tool-card">
                <div class="tool-card-icon">{icon}</div>
                <h3>{title}</h3>
                <p>{desc}</p>
                <span class="tool-card-tag">{tag}</span>
            </div>""",
            unsafe_allow_html=True,
        )
        if st.button(f"Buka {title}", key=f"open_{title}", use_container_width=True):
            st.switch_page(page)

# ROW 2
row2 = st.columns(3)
for col, (icon, title, desc, tag, page) in zip(row2, tools[3:]):
    with col:
        st.markdown(
            f"""<div class="tool-card">
                <div class="tool-card-icon">{icon}</div>
                <h3>{title}</h3>
                <p>{desc}</p>
                <span class="tool-card-tag">{tag}</span>
            </div>""",
            unsafe_allow_html=True,
        )
        if st.button(f"Buka {title}", key=f"open_{title}", use_container_width=True):
            st.switch_page(page)

# STEPS
st.markdown(
    """
    <div class="ilp-steps">
        <div class="ilp-step active"><div class="ilp-step-num">1</div><div class="ilp-step-label">Pilih Tool</div></div>
        <div class="ilp-step-line active"></div>
        <div class="ilp-step active"><div class="ilp-step-num">2</div><div class="ilp-step-label">Upload Data</div></div>
        <div class="ilp-step-line active"></div>
        <div class="ilp-step active"><div class="ilp-step-num">3</div><div class="ilp-step-label">Konfigurasi</div></div>
        <div class="ilp-step-line"></div>
        <div class="ilp-step"><div class="ilp-step-num">4</div><div class="ilp-step-label">Export</div></div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("---")
st.markdown(
    '<p style="text-align:center; color:#50506a; font-size:0.85rem;">'
    "Gunakan sidebar di kiri atau klik tombol di atas untuk navigasi antar tool</p>",
    unsafe_allow_html=True,
)

inject_footer()
