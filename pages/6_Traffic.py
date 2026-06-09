import streamlit as st
import requests
import pandas as pd
from style import inject_css, inject_sidebar_brand, inject_footer

st.set_page_config(page_title="Traffic Log", page_icon="📊", layout="wide")
inject_css()
inject_sidebar_brand()

st.markdown(
    """
    <div class="page-head">
        <div class="page-head-icon">📊</div>
        <h2>Traffic Log</h2>
        <p>Pantau aktivitas pengguna dan statistik penggunaan sistem</p>
    </div>
    """,
    unsafe_allow_html=True,
)

JSONBLOB_ID = "019e8740-72c4-7731-8328-0e2c67465233"
API_URL = f"https://jsonblob.com/api/jsonBlob/{JSONBLOB_ID}"

if st.button("Refresh Data"):
    st.rerun()

data = []
try:
    response = requests.get(API_URL, timeout=10)
    data = response.json()
    if not isinstance(data, list):
        data = []
except Exception:
    st.error("Gagal mengambil data dari server.")

if data:
    df = pd.DataFrame(data)

    total_entries = len(df)
    unique_users = df["Nama Pengguna"].nunique() if "Nama Pengguna" in df.columns else 0
    unique_stores = df["Nama Toko"].nunique() if "Nama Toko" in df.columns else 0
    latest = df.iloc[-1].get("Timestamp", "-") if len(df) > 0 else "-"

    st.markdown(
        f"""
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-card-value">{total_entries}</div>
                <div class="stat-card-label">Total Aktivitas</div>
            </div>
            <div class="stat-card">
                <div class="stat-card-value">{unique_users}</div>
                <div class="stat-card-label">Pengguna Unik</div>
            </div>
            <div class="stat-card">
                <div class="stat-card-value">{unique_stores}</div>
                <div class="stat-card-label">Toko Terdata</div>
            </div>
            <div class="stat-card">
                <div class="stat-card-value" style="font-size:1.1rem;">{latest}</div>
                <div class="stat-card-label">Aktivitas Terakhir</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("---")
    st.markdown("##### Riwayat Aktivitas")

    # Display columns (handle old data without Template field)
    display_cols = [
        "Nama Pengguna", "Nama Toko", "Tanggal QC", "Timestamp", "Template", "Opsi Layout"
    ]
    available_cols = [c for c in display_cols if c in df.columns]

    # Also handle old "Ukuran Layout" field
    if "Ukuran Layout" in df.columns and "Opsi Layout" not in df.columns:
        df = df.rename(columns={"Ukuran Layout": "Opsi Layout"})
        if "Opsi Layout" not in available_cols and "Ukuran Layout" in available_cols:
            available_cols = [c if c != "Ukuran Layout" else "Opsi Layout" for c in available_cols]
        if "Opsi Layout" not in available_cols:
            available_cols.append("Opsi Layout")

    # Fill missing columns with dash
    for col in display_cols:
        if col not in df.columns:
            df[col] = "-"

    available_cols = [c for c in display_cols if c in df.columns]

    st.dataframe(df[available_cols].iloc[::-1], use_container_width=True, hide_index=True)
else:
    st.info("Belum ada data traffic yang tercatat.")

inject_footer()
