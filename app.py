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

# TOOLS ROW 1
row1 = st.columns(3)

with row1[0]:
    st.markdown(
        """<div class="tool-card">
            <div class="tool-card-icon">📸</div>
            <h3>QC Image Inserter</h3>
            <p>Upload foto QC dan susun otomatis ke template Excel master</p>
            <span class="tool-card-tag">TOOL UTAMA</span>
        </div>""",
        unsafe_allow_html=True,
    )
    st.page_link("pages/1_QC_Image_Inserter.py", label="Buka QC Image Inserter", icon=":material/arrow_forward:")

with row1[1]:
    st.markdown(
        """<div class="tool-card">
            <div class="tool-card-icon">🗜️</div>
            <h3>Image Compress</h3>
            <p>Kompres batch foto untuk kurangi ukuran file tanpa hilang kualitas</p>
            <span class="tool-card-tag">UTILITAS</span>
        </div>""",
        unsafe_allow_html=True,
    )
    st.page_link("pages/2_Image_Compress.py", label="Buka Image Compress", icon=":material/arrow_forward:")

with row1[2]:
    st.markdown(
        """<div class="tool-card">
            <div class="tool-card-icon">💧</div>
            <h3>Watermark</h3>
            <p>Tambah watermark teks di foto seperti lokasi, tanggal, dan nama toko</p>
            <span class="tool-card-tag">UTILITAS</span>
        </div>""",
        unsafe_allow_html=True,
    )
    st.page_link("pages/3_Watermark.py", label="Buka Watermark", icon=":material/arrow_forward:")

# TOOLS ROW 2
row2 = st.columns(3)

with row2[0]:
    st.markdown(
        """<div class="tool-card">
            <div class="tool-card-icon">✏️</div>
            <h3>Batch Rename</h3>
            <p>Rename ratusan foto sekaligus dengan format penamaan yang konsisten</p>
            <span class="tool-card-tag">UTILITAS</span>
        </div>""",
        unsafe_allow_html=True,
    )
    st.page_link("pages/4_Batch_Rename.py", label="Buka Batch Rename", icon=":material/arrow_forward:")

with row2[1]:
    st.markdown(
        """<div class="tool-card">
            <div class="tool-card-icon">📄</div>
            <h3>PDF Converter</h3>
            <p>Ubah file Excel hasil export menjadi PDF siap kirim ke klien</p>
            <span class="tool-card-tag">KONVERSI</span>
        </div>""",
        unsafe_allow_html=True,
    )
    st.page_link("pages/5_PDF_Converter.py", label="Buka PDF Converter", icon=":material/arrow_forward:")

with row2[2]:
    st.markdown(
        """<div class="tool-card">
            <div class="tool-card-icon">📊</div>
            <h3>Traffic Log</h3>
            <p>Pantau aktivitas pengguna dan statistik penggunaan sistem</p>
            <span class="tool-card-tag">ANALITIK</span>
        </div>""",
        unsafe_allow_html=True,
    )
    st.page_link("pages/6_Traffic.py", label="Buka Traffic Log", icon=":material/arrow_forward:")

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
