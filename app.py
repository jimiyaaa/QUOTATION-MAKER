import streamlit as st
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from copy import deepcopy
from pathlib import Path
import tempfile, subprocess, shutil, os, re

APP_DIR = Path(__file__).resolve().parent
TEMPLATE = APP_DIR / "QT-format.docx"
TAX_RATE = 0.11

st.set_page_config(page_title="DSL Quotation Maker", page_icon="📄", layout="wide")
st.title("DSL Indonesia — Quotation Maker")
st.caption("Isi data quotation, tambahkan item bila diperlukan, lalu unduh dalam format DOCX dan PDF.")

def format_number(value):
    value = float(value or 0)
    if abs(value - round(value)) < 1e-9:
        raw = f"{int(round(value)):,}"
    else:
        raw = f"{value:,.2f}".rstrip("0").rstrip(".")
    return raw.replace(",", "_").replace(".", ",").replace("_", ".")

def set_highlighted_runs(paragraph, value):
    """Replace highlighted guide text while retaining the run's formatting."""
    for run in paragraph.runs:
        if run.font.highlight_color is not None:
            run.text = str(value)
            run.font.highlight_color = None

def clear_highlights(doc):
    for p in doc.paragraphs:
        for r in p.runs:
            r.font.highlight_color = None
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.highlight_color = None

def set_cell_text(cell, value, bold=None):
    # Keep the cell's table formatting; replace its paragraph contents.
    p = cell.paragraphs[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    run = p.add_run(str(value))
    if bold is not None:
        run.bold = bold
    for extra in cell.paragraphs[1:]:
        extra._element.getparent().remove(extra._element)

def add_item_rows(table, items):
    # Row 1 in the supplied template is the highlighted guide/example row.
    template_row = table.rows[1]
    template_xml = deepcopy(template_row._tr)
    # Remove the guide row; item rows will be inserted before the totals.
    table._tbl.remove(template_row._tr)
    totals_anchor = table.rows[1]._tr  # first totals row after guide removal
    for idx, item in enumerate(items, start=1):
        new_tr = deepcopy(template_xml)
        totals_anchor.addprevious(new_tr)
        # Reacquire row cells via a temporary wrapper.
        from docx.table import _Row
        new_row = _Row(new_tr, table)
        vals = [str(idx), item["description"], format_number(item["qty"]),
                format_number(item["unit_price"]), format_number(item["amount"])]
        for cell, val in zip(new_row.cells, vals):
            set_cell_text(cell, val)
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.highlight_color = None

def make_docx(data, items, subtotal, tax, grand_total, output_path):
    doc = Document(TEMPLATE)

    # Replace highlighted header/detail placeholders in the original template.
    # Paragraph indices refer to the supplied QT-format.docx.
    replacements = {
        16: data["company"],
        18: data["ref_num"],
        19: data["quoted_date"],
        20: data["valid_date"],
        31: data["delivery"],
        32: data["payment_term"],
        33: data["warranty"],
    }
    for index, value in replacements.items():
        set_highlighted_runs(doc.paragraphs[index], value)

    # Subject and Contact Person are separate highlighted runs in the same paragraph.
    p = doc.paragraphs[23]
    highlighted = [r for r in p.runs if r.font.highlight_color is not None]
    values = [data["subject"], data["contact_person"]]
    for run, value in zip(highlighted, values):
        run.text = str(value)
        run.font.highlight_color = None

    table = doc.tables[0]
    add_item_rows(table, items)

    # The highlighted "Note" text in the template is only a guide placeholder.
    # Remove that highlighted word while retaining the printed "Note :" label.
    for totals_row in table.rows[-4:-1]:
        for note_cell in totals_row.cells[:4]:
            for paragraph in note_cell.paragraphs:
                for run in paragraph.runs:
                    if run.font.highlight_color is not None:
                        run.text = ""
                        run.font.highlight_color = None

    # Totals are in the right-most cell of the three existing totals rows.
    table = doc.tables[0]
    totals_rows = table.rows[-4:-1]  # totals rows, followed by original spacer row
    # The row list is: header, item(s), subtotal, tax, grand total, spacer.
    set_cell_text(totals_rows[0].cells[4], format_number(subtotal), bold=True)
    set_cell_text(totals_rows[1].cells[4], format_number(tax))
    set_cell_text(totals_rows[2].cells[4], format_number(grand_total), bold=True)

    # Clear any remaining yellow guide highlighting anywhere in the document.
    clear_highlights(doc)
    doc.save(output_path)

def convert_to_pdf(docx_path, pdf_path):
    # Prefer LibreOffice/soffice; works without Microsoft Word.
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice and os.name == "nt":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        soffice = next((x for x in candidates if Path(x).exists()), None)
    if soffice:
        outdir = str(Path(pdf_path).parent)
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", outdir, str(docx_path)],
            capture_output=True, text=True, timeout=90
        )
        generated = Path(outdir) / (Path(docx_path).stem + ".pdf")
        if result.returncode == 0 and generated.exists():
            if generated.resolve() != Path(pdf_path).resolve():
                shutil.move(str(generated), str(pdf_path))
            return True, ""
        return False, (result.stderr or result.stdout or "LibreOffice gagal mengonversi dokumen.")
    # Fallback: docx2pdf requires Microsoft Word on Windows or macOS.
    try:
        from docx2pdf import convert
        convert(str(docx_path), str(pdf_path))
        return Path(pdf_path).exists(), ""
    except Exception as exc:
        return False, str(exc)

with st.form("quotation_form"):
    st.subheader("Informasi quotation")
    c1, c2 = st.columns(2)
    with c1:
        company = st.text_input("Nama Perusahaan *")
        subject = st.text_input("Subject")
        contact_person = st.text_input("Contact Person")
        ref_num = st.text_input("Ref Num")
        quoted_date = st.text_input("Quoted Date", placeholder="Contoh: 03 November 2026")
    with c2:
        valid_date = st.text_input("Valid Date", placeholder="Contoh: 30 November 2026")
        delivery = st.text_input("Delivery")
        payment_term = st.text_input("Payment term")
        warranty = st.text_input("Warranty")

    st.divider()
    st.subheader("Daftar item")
    st.caption("Minimal 1 baris. Isi Qty dan Unit Price; Amount akan dihitung otomatis setelah tombol Generate ditekan.")
    if "items" not in st.session_state:
        st.session_state.items = [{"description": "", "qty": 1.0, "unit_price": 0.0}]

    # Inputs are rendered outside the form on reruns, then the generate button below submits.
    # (The form is closed below after rendering all dynamic rows.)

    for i, item in enumerate(st.session_state.items):
        st.markdown(f"**Item {i+1}**")
        a, b, c = st.columns([4, 1, 2])
        item["description"] = a.text_input("Item desc", value=item["description"], key=f"desc_{i}")
        item["qty"] = b.number_input("Qty", min_value=0.0, value=float(item["qty"]), step=1.0, key=f"qty_{i}")
        item["unit_price"] = c.number_input("Unit price", min_value=0.0, value=float(item["unit_price"]), step=1000.0, key=f"price_{i}")
        amount = item["qty"] * item["unit_price"]
        st.caption(f"Amount: {format_number(amount)}")
    add_col, remove_col = st.columns(2)
    add_clicked = add_col.form_submit_button("＋ Tambah row", use_container_width=True)
    remove_clicked = remove_col.form_submit_button("－ Hapus row terakhir", use_container_width=True, disabled=len(st.session_state.items) <= 1)
    submitted = st.form_submit_button("Generate quotation", type="primary", use_container_width=True)

# Streamlit form submit buttons rerun the app; add/remove changes are handled here.
if add_clicked:
    st.session_state.items.append({"description": "", "qty": 1.0, "unit_price": 0.0})
    st.rerun()
if remove_clicked and len(st.session_state.items) > 1:
    st.session_state.items.pop()
    st.rerun()

if submitted:
    if not company.strip():
        st.error("Mohon isi Nama Perusahaan terlebih dahulu.")
    elif not TEMPLATE.exists():
        st.error("File QT-format.docx tidak ditemukan. Pastikan file template berada satu folder dengan app.py.")
    else:
        clean_items = []
        for item in st.session_state.items:
            qty = float(item["qty"])
            price = float(item["unit_price"])
            clean_items.append({
                "description": item["description"].strip(),
                "qty": qty,
                "unit_price": price,
                "amount": qty * price,
            })
        subtotal = sum(x["amount"] for x in clean_items)
        tax = subtotal * TAX_RATE
        grand_total = subtotal + tax
        data = {
            "company": company, "subject": subject, "contact_person": contact_person,
            "ref_num": ref_num, "quoted_date": quoted_date, "valid_date": valid_date,
            "delivery": delivery, "payment_term": payment_term, "warranty": warranty,
        }
        with tempfile.TemporaryDirectory() as tmp:
            docx_path = Path(tmp) / "Quotation.docx"
            pdf_path = Path(tmp) / "Quotation.pdf"
            make_docx(data, clean_items, subtotal, tax, grand_total, docx_path)
            docx_bytes = docx_path.read_bytes()
            st.success("Quotation DOCX berhasil dibuat.")
            st.download_button("⬇ Download DOCX", docx_bytes, file_name="Quotation.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               use_container_width=True)
            ok, message = convert_to_pdf(docx_path, pdf_path)
            if ok:
                st.success("Quotation PDF berhasil dibuat.")
                st.download_button("⬇ Download PDF", pdf_path.read_bytes(), file_name="Quotation.pdf",
                                   mime="application/pdf", use_container_width=True)
            else:
                st.warning("DOCX berhasil dibuat, tetapi PDF belum dapat dibuat di komputer ini. "
                           "Pasang LibreOffice (disarankan) atau Microsoft Word untuk konversi PDF.")
                with st.expander("Detail teknis"):
                    st.code(message or "PDF converter tidak ditemukan.")
                st.info("Anda tetap dapat membuka DOCX di Microsoft Word lalu pilih Save As → PDF.")
        st.markdown("### Ringkasan perhitungan")
        st.write(f"Sub total: **{format_number(subtotal)}**")
        st.write(f"PPN 11%: **{format_number(tax)}**")
        st.write(f"Grand total: **{format_number(grand_total)}**")

st.divider()
st.caption("Catatan: PPN dihitung tetap 11% sesuai instruksi. Periksa kembali detail quotation sebelum dikirim ke customer.")
