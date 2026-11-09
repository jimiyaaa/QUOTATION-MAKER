import streamlit as st
from docx import Document
from copy import deepcopy
from pathlib import Path
from datetime import datetime
import sqlite3, json, tempfile, subprocess, shutil, os, io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

APP_DIR = Path(__file__).resolve().parent
TEMPLATE = APP_DIR / "QT-new format.docx"
DB_PATH = APP_DIR / "quotation_history.db"
TAX_RATE = 0.11

st.set_page_config(page_title="QT Maker", page_icon="📄", layout="wide")
st.title("Quotation Maker")

# ----------------------- Database -----------------------

def db_connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with db_connect() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS quotations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                company TEXT NOT NULL,
                subject TEXT,
                contact_person TEXT,
                ref_num TEXT,
                quoted_date TEXT,
                valid_date TEXT,
                delivery TEXT,
                payment_term TEXT,
                warranty TEXT,
                items_json TEXT NOT NULL,
                subtotal REAL NOT NULL,
                tax REAL NOT NULL,
                grand_total REAL NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
        """)
        conn.commit()

def list_quotations(search_text=""):
    with db_connect() as conn:
        if search_text.strip():
            q = f"%{search_text.strip()}%"
            rows = conn.execute("""
                SELECT * FROM quotations
                WHERE company LIKE ? OR ref_num LIKE ? OR subject LIKE ?
                ORDER BY updated_at DESC, id DESC
            """, (q, q, q)).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM quotations ORDER BY updated_at DESC, id DESC"
            ).fetchall()
        return [dict(r) for r in rows]

def get_quotation(record_id):
    with db_connect() as conn:
        row = conn.execute("SELECT * FROM quotations WHERE id=?", (record_id,)).fetchone()
        return dict(row) if row else None

def save_quotation(data, items, subtotal, tax, grand_total, record_id=None):
    now = datetime.now().isoformat(timespec="seconds")
    values = (
        data["company"], data["subject"], data["contact_person"], data["ref_num"],
        data["quoted_date"], data["valid_date"], data["delivery"], data["payment_term"],
        data["warranty"], json.dumps(items, ensure_ascii=False), subtotal, tax,
        grand_total, now
    )
    with db_connect() as conn:
        if record_id:
            conn.execute("""
                UPDATE quotations SET
                    company=?, subject=?, contact_person=?, ref_num=?, quoted_date=?,
                    valid_date=?, delivery=?, payment_term=?, warranty=?, items_json=?,
                    subtotal=?, tax=?, grand_total=?, updated_at=?
                WHERE id=?
            """, values + (record_id,))
            conn.commit()
            return record_id
        cur = conn.execute("""
            INSERT INTO quotations (
                company, subject, contact_person, ref_num, quoted_date, valid_date,
                delivery, payment_term, warranty, items_json, subtotal, tax,
                grand_total, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, values[:13] + (now, now))
        conn.commit()
        return cur.lastrowid

def delete_quotation(record_id):
    with db_connect() as conn:
        conn.execute("DELETE FROM quotations WHERE id=?", (record_id,))
        conn.commit()

init_db()

# ----------------------- Formatting / DOCX -----------------------

def money(value):
    value = float(value or 0)
    if abs(value - round(value)) < 1e-8:
        s = f"{int(round(value)):,}"
    else:
        s = f"{value:,.2f}".rstrip("0").rstrip(".")
    return s.replace(",", "_").replace(".", ",").replace("_", ".")

def replace_highlighted(paragraph, value):
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
    replacements = {
        "company": 16, "ref_num": 17, "quoted_date": 18, "valid_date": 19,
        "subject": 21, "contact_person": 22, "delivery": 30,
        "payment_term": 31, "warranty": 32,
    }
    for key, paragraph_index in replacements.items():
        if paragraph_index < len(doc.paragraphs):
            replace_highlighted(doc.paragraphs[paragraph_index], data[key])

    table = doc.tables[0]
    sample_tr = deepcopy(table.rows[1]._tr)
    table._tbl.remove(table.rows[1]._tr)
    first_total_tr = table.rows[1]._tr
    for index, item in enumerate(items, start=1):
        new_tr = deepcopy(sample_tr)
        first_total_tr.addprevious(new_tr)
        from docx.table import _Row
        row = _Row(new_tr, table)
        values = [
            str(index), item["description"], money(item["qty"]),
            money(item["unit_price"]), money(item["amount"])
        ]
        for cell, value in zip(row.cells, values):
            replace_cell_text(cell, value)

    table = doc.tables[0]
    total_start = 1 + len(items)
    for row_index, value in [
        (total_start, subtotal),
        (total_start + 1, tax),
        (total_start + 2, grand_total),
    ]:
        replace_cell_text(table.rows[row_index].cells[4], money(value))

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
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        soffice = next((x for x in candidates if Path(x).exists()), None)
    if soffice:
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", str(Path(pdf_path).parent), str(docx_path)],
            capture_output=True, text=True, timeout=90
        )
        generated = Path(pdf_path).parent / (Path(docx_path).stem + ".pdf")
        if result.returncode == 0 and generated.exists():
            if generated.resolve() != Path(pdf_path).resolve():
                shutil.move(str(generated), str(pdf_path))
            return True, ""
        return False, result.stderr or result.stdout or "PDF conversion failed."
    try:
        from docx2pdf import convert
        convert(str(docx_path), str(pdf_path))
        return Path(pdf_path).exists(), ""
    except Exception as exc:
        return False, str(exc)

# ----------------------- XLSX summary -----------------------

def build_summary_xlsx(records):
    wb = Workbook()
    ws = wb.active
    ws.title = "Quotations"
    headers = [
        "Company", "Reference No.", "Subject", "Contact Person",
        "Quoted Date", "Valid Date", "Delivery", "Payment Term", "Warranty",
        "Item Count", "Sub Total", "Tax (11%)", "Grand Total",
        "Created At", "Last Updated"
    ]
    ws.append(headers)
    for rec in records:
        items = json.loads(rec["items_json"])
        ws.append([
            rec["company"], rec["ref_num"], rec["subject"],
            rec["contact_person"], rec["quoted_date"], rec["valid_date"],
            rec["delivery"], rec["payment_term"], rec["warranty"],
            len(items), rec["subtotal"], rec["tax"], rec["grand_total"],
            rec["created_at"], rec["updated_at"]
        ])

    detail = wb.create_sheet("Item Details")
    detail.append([
        "Company", "Reference No.", "Quoted Date",
        "Item No.", "Item Description", "Qty", "Unit Price", "Amount"
    ])
    for rec in records:
        for index, item in enumerate(json.loads(rec["items_json"]), start=1):
            detail.append([
                rec["company"], rec["ref_num"], rec["quoted_date"],
                index, item["description"], item["qty"],
                item["unit_price"], item["amount"]
            ])

    for sheet in wb.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="334155")
            cell.alignment = Alignment(vertical="center")
        for col in sheet.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=10)
            sheet.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 2, 12), 36)
    # Column indexes follow the visible workbook layout (internal record ID is omitted).
    for row in ws.iter_rows(min_row=2):
        for col_index in (11, 12, 13):  # Sub Total, Tax, Grand Total
            row[col_index - 1].number_format = '#,##0.00'
    for row in detail.iter_rows(min_row=2):
        for col_index in (7, 8):  # Unit Price, Amount
            row[col_index - 1].number_format = '#,##0.00'
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ----------------------- Session state helpers -----------------------

FIELD_DEFAULTS = {
    "company": "", "subject": "", "contact_person": "", "ref_num": "",
    "quoted_date": "", "valid_date": "", "delivery": "",
    "payment_term": "", "warranty": ""
}

def initialize_form():
    for key, value in FIELD_DEFAULTS.items():
        st.session_state.setdefault(key, value)
    st.session_state.setdefault("quote_items", [{"description": "", "qty": 1.0, "unit_price": 0.0}])
    st.session_state.setdefault("editing_id", None)

def capture_items_from_widgets():
    captured = []
    for i, old in enumerate(st.session_state["quote_items"]):
        captured.append({
            "description": st.session_state.get(f"item_desc_{i}", old.get("description", "")),
            "qty": float(st.session_state.get(f"item_qty_{i}", old.get("qty", 1.0))),
            "unit_price": float(st.session_state.get(f"item_price_{i}", old.get("unit_price", 0.0))),
        })
    return captured

def reset_form():
    for key, value in FIELD_DEFAULTS.items():
        st.session_state[key] = value
    for i in range(100):
        for prefix in ("item_desc_", "item_qty_", "item_price_"):
            st.session_state.pop(f"{prefix}{i}", None)
    st.session_state["quote_items"] = [{"description": "", "qty": 1.0, "unit_price": 0.0}]
    st.session_state["editing_id"] = None
    st.session_state.pop("generated_docx", None)
    st.session_state.pop("generated_pdf", None)
    st.session_state.pop("last_saved_id", None)
    st.session_state.pop("pdf_error", None)

def load_record(rec):
    reset_form()
    for key in FIELD_DEFAULTS:
        st.session_state[key] = rec.get(key) or ""
    loaded_items = json.loads(rec["items_json"])
    st.session_state["quote_items"] = loaded_items or [{"description": "", "qty": 1.0, "unit_price": 0.0}]
    st.session_state["editing_id"] = rec["id"]
    for i, item in enumerate(st.session_state["quote_items"]):
        st.session_state[f"item_desc_{i}"] = item.get("description", "")
        st.session_state[f"item_qty_{i}"] = float(item.get("qty", 1.0))
        st.session_state[f"item_price_{i}"] = float(item.get("unit_price", 0.0))

initialize_form()

# ----------------------- Navigation -----------------------

page = st.sidebar.radio("Menu", ["Create quotation", "Quotation history"], key="main_page")

if page == "Create quotation":
    editing_id = st.session_state.get("editing_id")
    if editing_id:
        st.info(f"Editing quotation record #{editing_id}. Saving will update this record.")
    if st.button("＋ New quotation", key="new_quotation"):
        reset_form()
        st.rerun()

    st.subheader("Quotation information")
    c1, c2 = st.columns(2)
    with c1:
        st.text_input("Company name *", key="company")
        st.text_input("Subject", key="subject")
        st.text_input("Contact person", key="contact_person")
        st.text_input("Reference number", key="ref_num")
        st.text_input("Quoted date", key="quoted_date")
    with c2:
        st.text_input("Valid date", key="valid_date")
        st.text_input("Delivery", key="delivery")
        st.text_input("Payment term", key="payment_term")
        st.text_input("Warranty", key="warranty")

    st.divider()
    st.subheader("Items")
    for i, item in enumerate(st.session_state["quote_items"]):
        st.markdown(f"**Item {i + 1}**")
        a, b, c, d = st.columns([4, 1, 2, 2])
        desc_key, qty_key, price_key = f"item_desc_{i}", f"item_qty_{i}", f"item_price_{i}"
        st.session_state.setdefault(desc_key, item.get("description", ""))
        st.session_state.setdefault(qty_key, float(item.get("qty", 1.0)))
        st.session_state.setdefault(price_key, float(item.get("unit_price", 0.0)))
        a.text_input("Item description", key=desc_key)
        b.number_input("Qty", min_value=0.0, key=qty_key, step=1.0)
        c.number_input("Unit price", min_value=0.0, key=price_key, step=1000.0)
        amount = float(st.session_state[qty_key]) * float(st.session_state[price_key])
        d.metric("Amount", money(amount))

    add_col, remove_col = st.columns(2)
    if add_col.button("＋ Add row", use_container_width=True):
        st.session_state["quote_items"] = capture_items_from_widgets()
        st.session_state["quote_items"].append({"description": "", "qty": 1.0, "unit_price": 0.0})
        st.rerun()
    if remove_col.button("－ Remove last row", disabled=len(st.session_state["quote_items"]) <= 1, use_container_width=True):
        st.session_state["quote_items"] = capture_items_from_widgets()
        st.session_state["quote_items"].pop()
        for prefix in ("item_desc_", "item_qty_", "item_price_"):
            st.session_state.pop(f"{prefix}{len(st.session_state['quote_items'])}", None)
        st.rerun()

    current_items = capture_items_from_widgets()
    calc_items = []
    for item in current_items:
        qty = float(item["qty"])
        price = float(item["unit_price"])
        calc_items.append({
            "description": item["description"], "qty": qty,
            "unit_price": price, "amount": qty * price
        })
    subtotal = sum(item["amount"] for item in calc_items)
    tax = subtotal * TAX_RATE
    grand_total = subtotal + tax

    st.divider()
    st.subheader("Summary")
    r1, r2, r3 = st.columns(3)
    r1.metric("Sub total", money(subtotal))
    r2.metric("Tax (11%)", money(tax))
    r3.metric("Grand total", money(grand_total))

    save_label = "Update quotation" if editing_id else "Save quotation"
    if st.button(save_label, type="primary", use_container_width=True):
        if not st.session_state["company"].strip():
            st.error("Please enter a company name.")
        else:
            data = {key: st.session_state.get(key, "") for key in FIELD_DEFAULTS}
            saved_id = save_quotation(
                data, calc_items, subtotal, tax, grand_total, editing_id
            )
            st.session_state["editing_id"] = saved_id
            st.session_state["last_saved_id"] = saved_id
            st.session_state["generated_docx"] = None
            st.session_state["generated_pdf"] = None
            st.success(f"Quotation saved. Record ID: {saved_id}")

    if st.session_state.get("last_saved_id"):
        saved_id = st.session_state["last_saved_id"]
        record = get_quotation(saved_id)
        if record:
            record_items = json.loads(record["items_json"])
            record_data = {key: record.get(key) or "" for key in FIELD_DEFAULTS}
            with tempfile.TemporaryDirectory() as tmp:
                docx_path = Path(tmp) / "Quotation.docx"
                pdf_path = Path(tmp) / "Quotation.pdf"
                build_docx(record_data, record_items, record["subtotal"], record["tax"], record["grand_total"], docx_path)
                st.session_state["generated_docx"] = docx_path.read_bytes()
                ok, msg = to_pdf(docx_path, pdf_path)
                st.session_state["generated_pdf"] = pdf_path.read_bytes() if ok else None
                st.session_state["pdf_error"] = msg if not ok else ""
            st.download_button(
                "Download DOCX", st.session_state["generated_docx"],
                file_name=f"Quotation_{saved_id}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            if st.session_state["generated_pdf"]:
                st.download_button(
                    "Download PDF", st.session_state["generated_pdf"],
                    file_name=f"Quotation_{saved_id}.pdf", mime="application/pdf",
                    use_container_width=True
                )
            elif st.session_state.get("pdf_error"):
                st.warning("DOCX is ready, but PDF conversion is unavailable. Install LibreOffice for reliable PDF export.")
                with st.expander("Technical details"):
                    st.code(st.session_state["pdf_error"])

elif page == "Quotation history":
    st.subheader("Quotation history")
    records_all = list_quotations()
    search_text = st.text_input("Search by company, reference number, or subject")
    records = list_quotations(search_text)
    st.caption(f"{len(records)} quotation record(s) found.")
    if records:
        rows = []
        for rec in records:
            rows.append({
                "Company": rec["company"],
                "Reference": rec["ref_num"],
                "Subject": rec["subject"],
                "Quoted date": rec["quoted_date"],
                "Grand total": money(rec["grand_total"]),
                "Last updated": rec["updated_at"].replace("T", " "),
            })
        st.dataframe(rows, use_container_width=True, hide_index=True)

        record_lookup = {rec["id"]: rec for rec in records}

        def quotation_label(record_id):
            rec = record_lookup[record_id]
            parts = [rec["company"] or "Unnamed company"]
            if rec["ref_num"]:
                parts.append(rec["ref_num"])
            if rec["quoted_date"]:
                parts.append(rec["quoted_date"])
            if rec["subject"]:
                parts.append(rec["subject"])
            return " | ".join(parts)

        selected_id = st.selectbox(
            "Select a quotation",
            options=list(record_lookup.keys()),
            format_func=quotation_label
        )
        selected = get_quotation(selected_id)
        left, right = st.columns(2)
        def open_selected_quotation(record):
            load_record(record)
            st.session_state["last_saved_id"] = None
            st.session_state["main_page"] = "Create quotation"

        left.button(
            "Open and edit",
            type="primary",
            use_container_width=True,
            on_click=open_selected_quotation,
            args=(selected,)
        )
        if right.button("Delete selected quotation", use_container_width=True):
            st.session_state["confirm_delete_id"] = selected_id
        if st.session_state.get("confirm_delete_id") == selected_id:
            st.warning(f"Delete quotation record #{selected_id}? This cannot be undone.")
            yes, no = st.columns(2)
            if yes.button("Confirm delete", type="primary"):
                delete_quotation(selected_id)
                st.session_state["confirm_delete_id"] = None
                st.success("Quotation deleted.")
                st.rerun()
            if no.button("Cancel delete"):
                st.session_state["confirm_delete_id"] = None
                st.rerun()
        with st.expander("Selected quotation details"):
            st.write(f"**Company:** {selected['company']}")
            st.write(f"**Reference:** {selected['ref_num'] or '-'}")
            st.write(f"**Subject:** {selected['subject'] or '-'}")
            st.write(f"**Grand total:** {money(selected['grand_total'])}")
            item_rows = json.loads(selected["items_json"])
            st.dataframe([
                {"No.": i, "Description": it["description"], "Qty": it["qty"],
                 "Unit price": money(it["unit_price"]), "Amount": money(it["amount"])}
                for i, it in enumerate(item_rows, 1)
            ], use_container_width=True, hide_index=True)
    else:
        st.info("No quotation records found.")

    st.divider()
    all_records = list_quotations()
    if all_records:
        xlsx_bytes = build_summary_xlsx(all_records)
        st.download_button(
            "Download all quotation summary (.xlsx)", xlsx_bytes,
            file_name=f"Quotation_Summary_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.button("Download all quotation summary (.xlsx)", disabled=True, use_container_width=True)
