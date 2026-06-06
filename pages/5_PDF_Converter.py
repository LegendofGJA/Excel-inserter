import streamlit as st
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
            with st.spinner("Mengonversi Excel ke PDF..."):
                uploaded_excel.seek(0)
                wb = load_workbook(uploaded_excel, data_only=True)
                ws = wb[selected_sheet]

                # Ambil semua data dari sheet
                all_data = []
                for row in ws.iter_rows(values_only=True):
                    all_data.append(row)
                wb.close()

                if not all_data:
                    st.error("Sheet kosong, tidak ada data untuk dikonversi.")
                else:
                    try:
                        from fpdf import FPDF

                        class QCPDF(FPDF):
                            def header(self):
                                self.set_font("Helvetica", "B", 10)
                                self.set_text_color(100, 100, 100)
                                self.cell(0, 8, f"Converted from: {uploaded_excel.name}", align="R")
                                self.ln(4)
                                self.set_draw_color(229, 50, 45)
                                self.set_line_width(0.5)
                                self.line(10, self.get_y(), self.w - 10, self.get_y())
                                self.ln(4)

                            def footer(self):
                                self.set_y(-15)
                                self.set_font("Helvetica", "I", 8)
                                self.set_text_color(150, 150, 150)
                                self.cell(0, 10, f"Page {self.page_no()}/{{nb}}", align="C")

                        pdf = QCPDF(orientation="L", unit="mm", format="A4")
                        pdf.alias_nb_pages()
                        pdf.set_auto_page_break(auto=True, margin=20)
                        pdf.add_page()

                        # Title
                        pdf.set_font("Helvetica", "B", 14)
                        pdf.set_text_color(40, 40, 40)
                        pdf.cell(0, 10, selected_sheet, ln=True)
                        pdf.ln(4)

                        # Hitung lebar kolom
                        num_cols = max(len(row) for row in all_data) if all_data else 1
                        usable_width = pdf.w - 20
                        col_width = usable_width / num_cols
                        col_width = min(col_width, 55)
                        row_height = 8

                        # Header row (baris pertama)
                        pdf.set_font("Helvetica", "B", 8)
                        pdf.set_fill_color(240, 240, 240)
                        pdf.set_text_color(40, 40, 40)
                        pdf.set_draw_color(200, 200, 200)

                        header_row = all_data[0] if all_data else []
                        for cell_val in header_row:
                            text = str(cell_val) if cell_val is not None else ""
                            if len(text) > 40:
                                text = text[:37] + "..."
                            pdf.cell(col_width, row_height, text, border=1, fill=True)
                        pdf.ln()

                        # Data rows
                        pdf.set_font("Helvetica", "", 7)
                        pdf.set_fill_color(255, 255, 255)

                        for row_data in all_data[1:]:
                            # Cek apakah perlu halaman baru
                            if pdf.get_y() + row_height > pdf.h - 20:
                                pdf.add_page()
                                # Ulangi header
                                pdf.set_font("Helvetica", "B", 8)
                                pdf.set_fill_color(240, 240, 240)
                                for cell_val in header_row:
                                    text = str(cell_val) if cell_val is not None else ""
                                    if len(text) > 40:
                                        text = text[:37] + "..."
                                    pdf.cell(col_width, row_height, text, border=1, fill=True)
                                pdf.ln()
                                pdf.set_font("Helvetica", "", 7)
                                pdf.set_fill_color(255, 255, 255)

                            pdf.set_text_color(50, 50, 50)
                            for idx, cell_val in enumerate(row_data):
                                if idx >= num_cols:
                                    break
                                text = str(cell_val) if cell_val is not None else ""
                                if len(text) > 40:
                                    text = text[:37] + "..."
                                pdf.cell(col_width, row_height, text, border=1)
                            pdf.ln()

                        pdf_output = bytes(pdf.output())

                        st.success(f"Konversi berhasil! {len(all_data)} baris data.")

                        st.download_button(
                            label="DOWNLOAD PDF",
                            data=pdf_output,
                            file_name=f"{selected_sheet}.pdf",
                            mime="application/pdf",
                            type="primary",
                            use_container_width=True,
                        )

                    except ImportError:
                        st.error(
                            "Library fpdf2 belum terinstall. "
                            "Pastikan 'fpdf2' ada di requirements.txt"
                        )

    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")
else:
    st.caption("Upload file Excel di atas untuk mulai konversi.")

inject_footer()
