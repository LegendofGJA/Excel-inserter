import streamlit as st
import subprocess
import os
import shutil
from io import BytesIO
from style import inject_css, inject_sidebar_brand, inject_footer

st.set_page_config(page_title="PDF Converter", page_icon="📄", layout="wide")
inject_css()
inject_sidebar_brand()

st.markdown(
    """
    <div class="page-head">
        <div class="page-head-icon">📄</div>
        <h2>PDF Converter</h2>
        <p>Ubah file Excel (.xlsx) menjadi PDF siap kirim ke klien atau atasan</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.info("Tool ini menggunakan library openpyxl untuk mengonversi data Excel menjadi PDF sederhana berbasis HTML.")

uploaded_excel = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

if uploaded_excel:
    try:
        from openpyxl import load_workbook

        uploaded_excel.seek(0)
        wb = load_workbook(uploaded_excel, data_only=True)
        sheet_names = wb.sheetnames
        selected_sheet = st.selectbox("Pilih Sheet untuk konversi:", sheet_names)
        wb.close()

        if st.button("KONVERSI KE PDF", type="primary", use_container_width=True):
            with st.spinner("Mengonversi Excel ke format yang bisa dicetak..."):
                uploaded_excel.seek(0)
                wb = load_workbook(uploaded_excel, data_only=True)
                ws = wb[selected_sheet]

                # Bangun HTML dari data Excel
                html_parts = [
                    "<!DOCTYPE html><html><head><meta charset='utf-8'>",
                    "<style>",
                    "body { font-family: Arial, sans-serif; padding: 20px; }",
                    "h2 { color: #333; margin-bottom: 4px; }",
                    "p.sub { color: #888; font-size: 13px; margin-bottom: 20px; }",
                    "table { border-collapse: collapse; width: 100%; margin-top: 10px; }",
                    "th, td { border: 1px solid #ccc; padding: 8px 12px; text-align: left; font-size: 13px; }",
                    "th { background: #f5f5f5; font-weight: 600; }",
                    "tr:nth-child(even) { background: #fafafa; }",
                    "</style></head><body>",
                    f"<h2>{selected_sheet}</h2>",
                    f"<p class='sub'>Converted from: {uploaded_excel.name}</p>",
                    "<table>",
                ]

                for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
                    html_parts.append("<tr>")
                    tag = "th" if row_idx == 0 else "td"
                    for cell in row:
                        val = str(cell) if cell is not None else ""
                        html_parts.append(f"<{tag}>{val}</{tag}>")
                    html_parts.append("</tr>")

                html_parts.append("</table></body></html>")
                html_content = "\n".join(html_parts)

                wb.close()

                # Coba buat PDF dengan weasyprint, fallback ke HTML download
                try:
                    from weasyprint import HTML as WPHTML
                    pdf_bytes = WPHTML(string=html_content).write_pdf()
                    st.success("Konversi berhasil!")
                    st.download_button(
                        label="DOWNLOAD PDF",
                        data=pdf_bytes,
                        file_name=f"{selected_sheet}.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True,
                    )
                except ImportError:
                    st.warning("Library weasyprint tidak tersedia. Menghasilkan HTML yang bisa di-print ke PDF.")
                    st.download_button(
                        label="DOWNLOAD HTML (Print ke PDF dari browser)",
                        data=html_content.encode("utf-8"),
                        file_name=f"{selected_sheet}.html",
                        mime="text/html",
                        type="primary",
                        use_container_width=True,
                    )
                    st.markdown(
                        '<p style="color:#50506a; font-size:0.82rem;">'
                        "Tips: Buka file HTML di browser, lalu tekan Ctrl+P (atau Cmd+P) dan pilih "
                        "'Save as PDF' untuk menghasilkan PDF.</p>",
                        unsafe_allow_html=True,
                    )

    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")
else:
    st.caption("Upload file Excel di atas untuk mulai konversi.")

inject_footer()
