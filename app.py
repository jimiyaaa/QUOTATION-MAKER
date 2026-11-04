import streamlit as st
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from copy import deepcopy
from pathlib import Path
import tempfile, subprocess, shutil, os

APP_DIR = Path(__file__).resolve().parent
TEMPLATE = APP_DIR / "QT-new format.docx"
TAX_RATE = 0.11

st.set_page_config(page_title="DSL Quotation Maker", page_icon="📄", layout="wide")
st.title("DSL Indonesia — Quotation Maker")
st.caption("Buat quotation berdasarkan template terbaru DSL Indonesia.")

def money(value):
    value = float(value or 0)
    if abs(value - round(value)) < 1e-8:
        s = f"{int(round(value)):,}"
    else:
        s = f"{value:,.2f}".rstrip("0").rstrip(".")
    return s.replace(",", "_").replace(".", ",").replace("_", ".")

def replace_highlighted(paragraph, value):
    """Replace only highlighted guide text, retaining its run formatting."""
    found = False
    for run in paragraph.runs:
        if run.font.highlight_color is not None:
            if not found:
                run.text = str(value)
                run.font.highlight_color = None
                found = True
            else:
                run.text = ""
                run.font.highlight_color = None
    return found

def replace_cell_text(cell, value):
    """Replace cell content while keeping the original first run's formatting."""
    p = cell.paragraphs[0]
    runs = p.runs
    if runs:
        runs[0].text = str(value)
        for r in runs[1:]:
            r.text = ""
        runs[0].font.highlight_color = None
    else:
        p.add_run(str(value))
    for extra in cell.paragraphs[1:]:
        for r in extra.runs:
            r.text = ""

def clear_all_highlights(doc):
    for p in doc.paragraphs:
        for r in p.runs:
            r.font.highlight_color = None
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.highlight_color = None

def build_docx(data, items, subtotal, tax, grand_total, destination):
    doc = Document(TEMPLATE)

    # The template uses highlighted runs as input markers. Replace those markers
    # directly instead of relying on paragraph indexes (which caused shifted text).
    replacements = {
        "company": 16,
        "ref_num": 17,
        "quoted_date": 18,
        "valid_date": 19,
        "subject": 21,
        "contact_person": 22,
        "delivery": 30,
        "payment_term": 31,
        "warranty": 32,
    }
    for key, paragraph_index in replacements.items():
        if paragraph_index < len(doc.paragraphs):
            replace_highlighted(doc.paragraphs[paragraph_index], data[key])

    # Item table: row 1 is the highlighted sample row. Clone it for each item.
    table = doc.tables[0]
    sample_tr = deepcopy(table.rows[1]._tr)
    table._tbl.remove(table.rows[1]._tr)

    # After removing sample row, rows 1-3 are the three total rows.
    first_total_tr = table.rows[1]._tr
    for index, item in enumerate(items, start=1):
        new_tr = deepcopy(sample_tr)
        first_total_tr.addprevious(new_tr)
        from docx.table import _Row
        row = _Row(new_tr, table)
        values = [
            str(index),
            item["description"],
            money(item["qty"]),
            money(item["unit_price"]),
            money(item["amount"]),
        ]
        for cell, value in zip(row.cells, values):
            replace_cell_text(cell, value)

    # Amount placeholders are in the final cell of each total row.
    # The labels on the left side of the table remain untouched.
    table = doc.tables[0]
    # Rows: header, item rows..., subtotal, tax, grand total, spacer.
    total_start = 1 + len(items)
    totals = [
        (total_start, subtotal),
        (total_start + 1, tax),
        (total_start + 2, grand_total),
    ]
    for row_index, value in totals:
        replace_cell_text(table.rows[row_index].cells[4], money(value))

    # Remove the yellow "Note" guide only; retain the printed "Note :" label.
    for row_index in range(total_start, total_start + 3):
        cell = table.rows[row_index].cells[0]
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                if run.font.highlight_color is not None:
                    run.text = ""
                    run.font.highlight_color = None

    clear_all_highlights(doc)
    doc.save(destination)

def to_pdf(docx_path, pdf_path):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice and os.name == "nt":
        for candidate in (
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ):
            if Path(candidate).exists():
                soffice = candidate
                break
    if soffice:
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir",
             str(Path(pdf_path).parent), str(docx_path)],
            capture_output=True, text=True, timeout=90
        )
        generated = Path(pdf_path).parent / (Path(docx_path).stem + ".pdf")
        if result.returncode == 0 and generated.exists():
            if generated.resolve() != Path(pdf_path).resolve():
                shutil.move(str(generated), str(pdf_path))
            return True, ""
        return False, result.stderr or result.stdout or "Konversi PDF gagal."
    try:
        from docx2pdf import convert
        convert(str(docx_path), str(pdf_path))
        return Path(pdf_path).exists(), ""
    except Exception as exc:
        return False, str(exc)

# Form state
if "quote_items" not in st.session_state:
    st.session_state["quote_items"] = [{"description": "", "qty": 1.0, "unit_price": 0.0}]

st.subheader("Informasi quotation")
c1, c2 = st.columns(2)
with c1:
    company = st.text_input("Nama Perusahaan *", key="company")
    subject = st.text_input("Subject", key="subject")
    contact_person = st.text_input("Contact Person", key="contact_person")
    ref_num = st.text_input("Ref Num", key="ref_num")
    quoted_date = st.text_input("Quoted Date", placeholder="Contoh: 04/11/2026", key="quoted_date")
with c2:
    valid_date = st.text_input("Valid Date", placeholder="Contoh: 05/11/2026", key="valid_date")
    delivery = st.text_input("Delivery", key="delivery")
    payment_term = st.text_input("Payment term", key="payment_term")
    warranty = st.text_input("Warranty", key="warranty")

st.divider()
st.subheader("Daftar item")
st.caption("Minimal satu baris. Amount, Sub Total, PPN 11%, dan Grand Total dihitung otomatis.")
for i, item in enumerate(st.session_state["quote_items"]):
    st.markdown(f"**Item {i + 1}**")
    a, b, c, d = st.columns([4, 1, 2, 2])
    item["description"] = a.text_input("Item desc", value=item["description"], key=f"desc_{i}")
    item["qty"] = b.number_input("Qty", min_value=0.0, value=float(item["qty"]), step=1.0, key=f"qty_{i}")
    item["unit_price"] = c.number_input("Unit price", min_value=0.0, value=float(item["unit_price"]), step=1000.0, key=f"price_{i}")
    amount = float(item["qty"]) * float(item["unit_price"])
    d.metric("Amount", money(amount))

add_col, remove_col = st.columns(2)
if add_col.button("＋ Tambah row", use_container_width=True):
    st.session_state["quote_items"].append({"description": "", "qty": 1.0, "unit_price": 0.0})
    st.rerun()
if remove_col.button("－ Hapus row terakhir", disabled=len(st.session_state["quote_items"]) <= 1, use_container_width=True):
    st.session_state["quote_items"].pop()
    st.rerun()

items_calc = []
for item in st.session_state["quote_items"]:
    qty = float(item["qty"])
    price = float(item["unit_price"])
    items_calc.append({
        "description": item["description"],
        "qty": qty,
        "unit_price": price,
        "amount": qty * price,
    })
subtotal = sum(x["amount"] for x in items_calc)
tax = subtotal * TAX_RATE
grand_total = subtotal + tax

st.divider()
st.subheader("Ringkasan")
r1, r2, r3 = st.columns(3)
r1.metric("Sub Total", money(subtotal))
r2.metric("PPN 11%", money(tax))
r3.metric("Grand Total", money(grand_total))

if st.button("Generate quotation", type="primary", use_container_width=True):
    if not company.strip():
        st.error("Mohon isi Nama Perusahaan terlebih dahulu.")
    elif not TEMPLATE.exists():
        st.error("Template QT-new format.docx tidak ditemukan dalam folder aplikasi.")
    else:
        data = {
            "company": company, "subject": subject, "contact_person": contact_person,
            "ref_num": ref_num, "quoted_date": quoted_date, "valid_date": valid_date,
            "delivery": delivery, "payment_term": payment_term, "warranty": warranty,
        }
        with tempfile.TemporaryDirectory() as tmp:
            docx_path = Path(tmp) / "Quotation.docx"
            pdf_path = Path(tmp) / "Quotation.pdf"
            build_docx(data, items_calc, subtotal, tax, grand_total, docx_path)
            st.session_state["generated_docx"] = docx_path.read_bytes()
            ok, message = to_pdf(docx_path, pdf_path)
            st.session_state["generated_pdf"] = pdf_path.read_bytes() if ok else None
            st.session_state["pdf_error"] = message if not ok else ""
        st.success("Quotation berhasil dibuat. Silakan download file di bawah.")

if st.session_state.get("generated_docx"):
    st.download_button(
        "⬇ Download DOCX", st.session_state["generated_docx"],
        file_name="Quotation.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
if st.session_state.get("generated_pdf"):
    st.download_button(
        "⬇ Download PDF", st.session_state["generated_pdf"],
        file_name="Quotation.pdf", mime="application/pdf",
        use_container_width=True
    )
elif st.session_state.get("pdf_error") is not None and st.session_state.get("pdf_error"):
    st.warning("DOCX berhasil dibuat, tetapi PDF belum tersedia. Pasang LibreOffice untuk konversi PDF yang lebih stabil.")
    with st.expander("Detail teknis"):
        st.code(st.session_state["pdf_error"])

st.divider()
st.caption("PPN dihitung tetap 11%. Periksa kembali hasil quotation sebelum dikirim ke customer.")
