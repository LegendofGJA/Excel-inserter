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

# TOOLS GRID (2 rows x 3 cols)
st.markdown(
    """
    <div class="tools-grid">
        <a href="/QC_Image_Inserter" class="tool-card-link">
            <div class="tool-card">
                <div class="tool-card-icon">📸</div>
                <h3>QC Image Inserter</h3>
                <p>Upload foto QC dan susun otomatis ke template Excel master</p>
                <span class="tool-card-tag">TOOL UTAMA</span>
            </div>
        </a>
        <a href="/Image_Compress" class="tool-card-link">
            <div class="tool-card">
                <div class="tool-card-icon">🗜️</div>
                <h3>Image Compress</h3>
                <p>Kompres batch foto untuk kurangi ukuran file tanpa hilang kualitas</p>
                <span class="tool-card-tag">UTILITAS</span>
            </div>
        </a>
        <a href="/Watermark" class="tool-card-link">
            <div class="tool-card">
                <div class="tool-card-icon">💧</div>
                <h3>Watermark</h3>
                <p>Tambah watermark teks di foto seperti lokasi, tanggal, dan nama toko</p>
                <span class="tool-card-tag">UTILITAS</span>
            </div>
        </a>
        <a href="/Batch_Rename" class="tool-card-link">
            <div class="tool-card">
                <div class="tool-card-icon">✏️</div>
                <h3>Batch Rename</h3>
                <p>Rename ratusan foto sekaligus dengan format penamaan yang konsisten</p>
                <span class="tool-card-tag">UTILITAS</span>
            </div>
        </a>
        <a href="/PDF_Converter" class="tool-card-link">
            <div class="tool-card">
                <div class="tool-card-icon">📄</div>
                <h3>PDF Converter</h3>
                <p>Ubah file Excel hasil export menjadi PDF siap kirim ke klien</p>
                <span class="tool-card-tag">KONVERSI</span>
            </div>
        </a>
        <a href="/Traffic" class="tool-card-link">
            <div class="tool-card">
                <div class="tool-card-icon">📊</div>
                <h3>Traffic Log</h3>
                <p>Pantau aktivitas pengguna dan statistik penggunaan sistem</p>
                <span class="tool-card-tag">ANALITIK</span>
            </div>
        </a>
    </div>
    """,
    unsafe_allow_html=True,
)

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
    "Gunakan sidebar di kiri atau klik kartu tool di atas untuk mulai</p>",
    unsafe_allow_html=True,
)

inject_footer()
