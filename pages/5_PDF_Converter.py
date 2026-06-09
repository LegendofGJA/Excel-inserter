import streamlit as st
import os
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

# Find available Unicode font
FONT_REGULAR = None
FONT_BOLD = None

_search_paths_reg = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "/usr/share/fonts/TTF/DejaVuSans.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans.ttf",
]
_search_paths_bold = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    "/usr/share/fonts/TTF/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
]

for p in _search_paths_reg:
    if os.path.exists(p):
        FONT_REGULAR = p
        break

for p in _search_paths_bold:
    if os.path.exists(p):
        FONT_BOLD = p
        break

uploaded_excel = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

if uploaded_excel:
    try:
        from openpyxl import load_workbook

        uploaded_excel.seek(0)
        wb = load_workbook(uploaded_excel, data_only=True)
        sheet_names = wb.sheetnames
        selected_sheet = st.selectbox("Pilih Sheet untuk konversi:", sheet_names)

        # Check if sheet has images
        ws_check = wb[selected_sheet]
        has_images = len(ws_check._images) > 0 if hasattr(ws_check, '_images') else False
        wb.close()

        if has_images:
            st.info(
                "Sheet ini mengandung gambar. "
                "Gambar akan di-extract dan dimasukkan ke PDF."
            )

        if st.button("KONVERSI KE PDF", type="primary", use_container_width=True):
            with st.spinner("Mengonversi Excel ke PDF..."):
                uploaded_excel.seek(0)
                wb = load_workbook(uploaded_excel, data_only=True)
                ws = wb[selected_sheet]

                all_data = []
                for row in ws.iter_rows(values_only=True):
                    all_data.append(row)

                # Extract images from sheet
                sheet_images = []
                if hasattr(ws, '_images'):
                    for img in ws._images:
                        try:
                            from openpyxl.drawing.image import Image as XlImg
                            img_data = img._data()
                            sheet_images.append({
                                "anchor": str(img.anchor) if hasattr(img, 'anchor') else "",
                                "data": img_data,
                            })
                        except Exception:
                            pass

                wb.close()

                if not all_data and not sheet_images:
                    st.error("Sheet kosong, tidak ada data.")
                else:
                    try:
                        from fpdf import FPDF
                        from fpdf.enums import XPos, YPos
                        from PIL import Image as PILImage

                        # Register fonts BEFORE creating FPDF instance
                        pdf = FPDF(orientation="L", unit="mm", format="A4")
                        pdf.alias_nb_pages()
                        pdf.set_auto_page_break(auto=True, margin=20)

                        # Register Unicode fonts
                        if FONT_REGULAR:
                            pdf.add_font("QCF", "", FONT_REGULAR)
                        if FONT_BOLD:
                            pdf.add_font("QCF", "B", FONT_BOLD)

                        def set_pdf_font(pdf, style, size):
                            if FONT_REGULAR:
                                if style == "B" and FONT_BOLD:
                                    pdf.set_font("QCF", "B", size)
                                else:
                                    pdf.set_font("QCF", "", size)
                            else:
                                pdf.set_font("Helvetica", style, size)

                        pdf.add_page()

                        # Title
                        set_pdf_font(pdf, "B", 14)
                        pdf.set_text_color(40, 40, 40)
                        pdf.cell(
                            0, 10, str(selected_sheet),
                            new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                        )
                        pdf.ln(2)

                        # Subtitle
                        set_pdf_font(pdf, "", 8)
                        pdf.set_text_color(120, 120, 120)
                        pdf.cell(
                            0, 6, f"Source: {uploaded_excel.name}",
                            new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                        )
                        pdf.ln(4)

                        # Separator line
                        pdf.set_draw_color(229, 50, 45)
                        pdf.set_line_width(0.5)
                        pdf.line(10, pdf.get_y(), pdf.w - 10, pdf.get_y())
                        pdf.ln(6)

                        # Table data
                        if all_data:
                            num_cols = max(len(row) for row in all_data) if all_data else 1
                            usable_width = pdf.w - 20
                            col_width = min(usable_width / num_cols, 55)
                            row_height = 8

                            # Header row
                            set_pdf_font(pdf, "B", 8)
                            pdf.set_fill_color(240, 240, 240)
                            pdf.set_text_color(40, 40, 40)
                            pdf.set_draw_color(200, 200, 200)

                            header_row = all_data[0] if all_data else []
                            for cell_val in header_row:
                                text = str(cell_val) if cell_val is not None else ""
                                # Sanitize: replace Unicode chars that font can't handle
                                text = text.replace("\u2013", "-").replace("\u2014", "-")
                                text = text.replace("\u2018", "'").replace("\u2019", "'")
                                text = text.replace("\u201c", '"').replace("\u201d", '"')
                                text = text.replace("\u2026", "...")
                                if len(text) > 45:
                                    text = text[:42] + "..."
                                pdf.cell(
                                    col_width, row_height, text,
                                    border=1, fill=True,
                                    new_x=XPos.RIGHT, new_y=YPos.TOP,
                                )
                            pdf.ln()

                            # Data rows
                            set_pdf_font(pdf, "", 7)
                            pdf.set_fill_color(255, 255, 255)

                            for row_data in all_data[1:]:
                                if pdf.get_y() + row_height > pdf.h - 20:
                                    pdf.add_page()
                                    set_pdf_font(pdf, "B", 8)
                                    pdf.set_fill_color(240, 240, 240)
                                    for cell_val in header_row:
                                        text = str(cell_val) if cell_val is not None else ""
                                        text = text.replace("\u2013", "-").replace("\u2014", "-")
                                        text = text.replace("\u2018", "'").replace("\u2019", "'")
                                        text = text.replace("\u201c", '"').replace("\u201d", '"')
                                        text = text.replace("\u2026", "...")
                                        if len(text) > 45:
                                            text = text[:42] + "..."
                                        pdf.cell(
                                            col_width, row_height, text,
                                            border=1, fill=True,
                                            new_x=XPos.RIGHT, new_y=YPos.TOP,
                                        )
                                    pdf.ln()
                                    set_pdf_font(pdf, "", 7)
                                    pdf.set_fill_color(255, 255, 255)

                                pdf.set_text_color(50, 50, 50)
                                for idx, cell_val in enumerate(row_data):
                                    if idx >= num_cols:
                                        break
                                    text = str(cell_val) if cell_val is not None else ""
                                    text = text.replace("\u2013", "-").replace("\u2014", "-")
                                    text = text.replace("\u2018", "'").replace("\u2019", "'")
                                    text = text.replace("\u201c", '"').replace("\u201d", '"')
                                    text = text.replace("\u2026", "...")
                                    if len(text) > 45:
                                        text = text[:42] + "..."
                                    pdf.cell(
                                        col_width, row_height, text,
                                        border=1,
                                        new_x=XPos.RIGHT, new_y=YPos.TOP,
                                    )
                                pdf.ln()

                        # Add images from sheet
                        if sheet_images:
                            pdf.add_page()
                            set_pdf_font(pdf, "B", 12)
                            pdf.set_text_color(40, 40, 40)
                            pdf.cell(
                                0, 10, "Embedded Images from Sheet",
                                new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                            )
                            pdf.ln(6)

                            for idx, img_info in enumerate(sheet_images):
                                try:
                                    img_buf = BytesIO(img_info["data"])
                                    pil_img = PILImage.open(img_buf)

                                    if pil_img.mode in ("RGBA", "P"):
                                        pil_img = pil_img.convert("RGB")

                                    # Save temp
                                    tmp = BytesIO()
                                    pil_img.save(tmp, format="JPEG", quality=90)
                                    tmp.seek(0)

                                    # Fit on page
                                    max_w = pdf.w - 20
                                    max_h = pdf.h - 40
                                    img_w, img_h = pil_img.size
                                    ratio = min(max_w / img_w, max_h / img_h)
                                    w = img_w * ratio
                                    h = img_h * ratio

                                    # Check page break
                                    if pdf.get_y() + h > pdf.h - 20:
                                        pdf.add_page()

                                    tmp_path = f"/tmp/pdf_img_{idx}.jpg"
                                    with open(tmp_path, "wb") as f:
                                        f.write(img_info["data"])

                                    pdf.image(tmp_path, x=10, y=pdf.get_y(), w=w, h=h)
                                    pdf.ln(h + 4)

                                    # Caption
                                    set_pdf_font(pdf, "", 7)
                                    pdf.set_text_color(150, 150, 150)
                                    pdf.cell(
                                        0, 5, f"Image {idx + 1}",
                                        new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                                    )
                                    pdf.ln(6)

                                    if os.path.exists(tmp_path):
                                        os.remove(tmp_path)

                                except Exception as img_err:
                                    set_pdf_font(pdf, "", 8)
                                    pdf.set_text_color(200, 60, 60)
                                    pdf.cell(
                                        0, 8, f"[Image {idx + 1} gagal dimuat: {str(img_err)[:50]}]",
                                        new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                                    )

                        # Footer on every page
                        total_pages = pdf.page
                        for pg in range(1, total_pages + 1):
                            pdf.page = pg
                            pdf.set_y(-15)
                            set_pdf_font(pdf, "", 7)
                            pdf.set_text_color(150, 150, 150)
                            pdf.cell(
                                0, 10, f"Page {pg}/{total_pages}",
                                align="C",
                                new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                            )

                        pdf_output = bytes(pdf.output())

                        img_count_msg = f" + {len(sheet_images)} gambar" if sheet_images else ""
                        st.success(f"Konversi berhasil! {len(all_data)} baris data{img_count_msg}.")

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
