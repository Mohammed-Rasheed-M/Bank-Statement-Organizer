

import streamlit as st
from io import BytesIO, StringIO
import pandas as pd
import sqlite3
import asyncio
from datetime import datetime
from zoneinfo import ZoneInfo
import re
from typing import List, Tuple, Dict, Any, Optional
import hashlib

# AgGrid
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode

# ReportLab
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet

# ----------------------------
# CONFIG / CONSTANTS
# ----------------------------
DB_PATH = "statements_app.db"
DEFAULT_ITEMS = ["Hardware", "Plywoods", "False Ceiling", "Labours", "Transportation", "Factory finishing", "Miscellaneous", "Glass work", "civil work", "Electrical ,plumbing , civil", "Office"]
DATE_CANDIDATES = ["Date", "Transaction Date", "Value Dt", "Value Date"]
DESC_CANDIDATES = ["Narration", "Description", "Narration/Remarks", "Details"]
WITHDRAW_CANDIDATES = ["Withdrawal", "Withdrawal Amt.", "Withdrawal Amt", "Withdrawal(Dr)", "Withdrawal(Dr)/ Deposit(Cr)"]
DEPOSIT_CANDIDATES = ["Deposit", "Deposit Amt.", "Deposit Amt", "Deposit(Cr)", "Withdrawal(Dr)/ Deposit(Cr)"]
PAGE_TITLE = "Bank Statement Collector"
BANK_NAME_FROM_FILENAME = True

st.set_page_config(page_title=PAGE_TITLE, layout="wide")

# ----------------------------
# UTIL: SQLite persistence (kept; no explicit save button in UI)
# ----------------------------
def init_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS items (
        id INTEGER PRIMARY KEY,
        name TEXT UNIQUE NOT NULL
    )""")
    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        username TEXT PRIMARY KEY,
        last_seen TIMESTAMP
    )""")
    cur.execute("""
    CREATE TABLE IF NOT EXISTS rows (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT,
        s_no INTEGER,
        date TEXT,
        description TEXT,
        amount REAL,
        type TEXT,
        bank TEXT,
        site_name TEXT,
        item TEXT,
        worknotes TEXT,
        source_file TEXT,
        UNIQUE(username, s_no, source_file)
    )""")
    conn.commit()
    for it in DEFAULT_ITEMS:
        try:
            cur.execute("INSERT INTO items (name) VALUES (?)", (it,))
        except sqlite3.IntegrityError:
            pass
    conn.commit()
    return conn

db_conn = init_db()

def get_items_from_db() -> List[str]:
    cur = db_conn.cursor()
    cur.execute("SELECT name FROM items ORDER BY name")
    return [r[0] for r in cur.fetchall()]

def add_item_to_db(name: str):
    cur = db_conn.cursor()
    try:
        cur.execute("INSERT INTO items (name) VALUES (?)", (name,))
        db_conn.commit()
    except sqlite3.IntegrityError:
        pass


def remove_item_from_db(names):
    if not names:
        return
    cur = db_conn.cursor()
    if isinstance(names, (list, tuple, set)):
        cur.executemany("DELETE FROM items WHERE name = ?", [(n,) for n in names])
    else:
        cur.execute("DELETE FROM items WHERE name = ?", (names,))
    db_conn.commit()

# ----------------------------
# FILE PARSING HELPERS
# ----------------------------
def _line_has_table_header(line: str) -> bool:
    low = line.lower()
    return ("date" in low) and (("narration" in low or "description" in low) or ("withdraw" in low or "deposit" in low))

def extract_bank_name_from_text(lines: List[str]) -> Optional[str]:
    joined = " | ".join(lines)
    m = re.search(r'([A-Z][A-Z &.\'-]{2,}?BANK(?:\s+LIMITED|\s+LTD\.?)?)', joined, flags=re.I)
    if m:
        return re.sub(r'\s+', ' ', m.group(1)).strip()
    m2 = re.search(r'(HDFC|ICICI|SBI|STATE BANK OF INDIA|KOTAK|AXIS|INDUSIND|YES BANK|IDFC FIRST|CANARA|UNION BANK|RBL)\b', joined, flags=re.I)
    if m2:
        return (m2.group(1).upper() + " BANK").strip()
    for ln in lines:
        s = str(ln).strip()
        if s and len(s) <= 60 and "statement" not in s.lower() and "page" not in s.lower():
            return s
    return None

def read_csv_table_only(file_bytes: bytes) -> Tuple[pd.DataFrame, List[str], Optional[str]]:
    warnings = []
    text = file_bytes.decode("latin-1", errors="ignore")
    lines = text.splitlines()
    bank = extract_bank_name_from_text(lines[:120])
    if bank:
        warnings.append(f"Detected bank: {bank}")
    header_idx = None
    for i, line in enumerate(lines[:200]):
        if _line_has_table_header(line):
            header_idx = i
            break
    if header_idx is None:
        df = pd.read_csv(BytesIO(file_bytes), engine="python", on_bad_lines="skip")
    else:
        sliced = "\n".join(lines[header_idx:])
        df = pd.read_csv(StringIO(sliced), engine="python", on_bad_lines="skip")
    return df, warnings, bank

def read_excel_first_sheet_table_only(file_bytes: bytes) -> Tuple[pd.DataFrame, Optional[str], List[str], Optional[str]]:
    warnings = []
    xls = pd.ExcelFile(BytesIO(file_bytes))
    sheet_name = xls.sheet_names[0]
    warnings.append(f"Processing only first sheet: '{sheet_name}'")
    tmp = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, header=None, dtype=str)
    bank = extract_bank_name_from_text(tmp.head(40).astype(str).fillna("").apply(lambda r: " ".join(r), axis=1).tolist())
    if bank:
        warnings.append(f"Detected bank: {bank}")
    header_row = None
    max_scan = min(200, len(tmp))
    for i in range(max_scan):
        row_vals = [str(v) for v in tmp.iloc[i].fillna("").tolist()]
        if _line_has_table_header(",".join(row_vals)):
            header_row = i
            break
    if header_row is None:
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, dtype=str)
        return df, sheet_name, warnings, bank
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, header=header_row, dtype=str)
    return df, sheet_name, warnings, bank

def normalize_date(val: Any) -> str:
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if not s:
        return ""
    fmts = ("%d-%m-%Y","%d/%m/%Y","%d-%m-%y","%Y-%m-%d","%d %b %Y","%d %b %y","%d/%m/%y","%d-%b-%Y","%d-%b-%y")
    for f in fmts:
        try:
            return datetime.strptime(s, f).strftime("%Y-%m-%d")
        except:
            pass
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        return "" if pd.isna(dt) else dt.strftime("%Y-%m-%d")
    except:
        return ""

def parse_amount_and_type(val: Any) -> Tuple[float, str]:
    if pd.isna(val):
        return 0.0, ""
    s = str(val).strip()
    if not s:
        return 0.0, ""
    typ = ""
    if re.search(r"\bdr\b|\(dr\)|\bdebit\b", s, flags=re.I):
        typ = "Dr"
    if re.search(r"\bcr\b|\(cr\)|\bcredit\b", s, flags=re.I):
        typ = "Cr" if typ == "" else typ
    m = re.search(r"[-]?\d[\d,]*\.?\d*", s.replace("(", "").replace(")", ""))
    if m:
        try:
            return float(m.group(0).replace(",", "")), typ
        except:
            return 0.0, typ
    digits = re.sub(r"[^\d.-]", "", s)
    try:
        return (float(digits) if digits not in ("", "-", ".") else 0.0), typ
    except:
        return 0.0, typ

def standardize_df(df: pd.DataFrame, filename: str, inferred_bank: Optional[str]) -> Tuple[pd.DataFrame, Optional[str]]:
    if df is None or df.empty:
        return pd.DataFrame(), "Empty sheet or file."
    df.columns = [str(c).strip() for c in df.columns]

    def pick(cands: List[str]) -> Optional[str]:
        for c in cands:
            if c in df.columns:
                return c
            for actual in df.columns:
                if actual.lower() == c.lower():
                    return actual
        return None

    date_col = pick(DATE_CANDIDATES)
    desc_col = pick(DESC_CANDIDATES)
    withdraw_col = pick(WITHDRAW_CANDIDATES)
    deposit_col = pick(DEPOSIT_CANDIDATES)

    amount_col = None
    if not withdraw_col and not deposit_col:
        for alt in ["Withdrawal(Dr)/ Deposit(Cr)", "Amount", "Amount(INR)"]:
            if alt in df.columns: amount_col = alt; break

    if not date_col or not desc_col or (not withdraw_col and not deposit_col and not amount_col):
        return pd.DataFrame(), f"Missing required columns in {filename}. Found: {', '.join(df.columns)}"

    bank_for_rows = inferred_bank or (filename.split()[0] if BANK_NAME_FROM_FILENAME else "")

    out_rows = []
    for _, r in df.iterrows():
        d = normalize_date(r.get(date_col, ""))
        if not d:
            continue
        desc = str(r.get(desc_col, ""))

        amount = 0.0
        typ = ""

        if withdraw_col:
            a, _ = parse_amount_and_type(r.get(withdraw_col, ""))
            if a:
                amount = a; typ = "Dr"
        if (amount == 0.0) and deposit_col:
            a, _ = parse_amount_and_type(r.get(deposit_col, ""))
            if a:
                amount = a; typ = "Cr"

        if amount == 0.0 and amount_col:
            a, t = parse_amount_and_type(r.get(amount_col, ""))
            amount = a
            if t:
                typ = t
            else:
                if a < 0:
                    typ = "Dr"; amount = abs(a)
                elif a > 0:
                    typ = "Cr"

        out_rows.append({
            "Date": d,
            "Description": desc,
            "Amount": float(amount) if amount else 0.0,
            "Type": typ,
            "Bank": bank_for_rows,
            "SourceFile": filename,
            "Site Name": "",
            "Item": "",
            "WorkNotes": ""
        })
    return pd.DataFrame(out_rows), None

async def parse_single_file(uploaded) -> Tuple[pd.DataFrame, List[str]]:
    warnings = []
    name = uploaded.name
    raw = uploaded.read()
    inferred_bank = ""

    try:
        if name.lower().endswith(".csv"):
            df, warns, bank = read_csv_table_only(raw)
            warnings.extend(warns)
            inferred_bank = bank or ""
            std, warn = standardize_df(df, name, inferred_bank)
            if warn: warnings.append(warn)
            return std, warnings

        elif name.lower().endswith((".xlsx", ".xls")):
            df, sheet_name, warns, bank = read_excel_first_sheet_table_only(raw)
            warnings.extend(warns)
            inferred_bank = bank or ""
            std, swarn = standardize_df(df, name, inferred_bank)
            if swarn: warnings.append(swarn)
            return std, warnings

        else:
            return pd.DataFrame(), [f"Unsupported file type: {name}"]
    except Exception as e:
        return pd.DataFrame(), [f"Failed to parse {name}: {e}"]

async def parse_multiple_files(files) -> Tuple[pd.DataFrame, List[str]]:
    tasks = [parse_single_file(f) for f in files]
    results = await asyncio.gather(*tasks)
    dfs = []
    warnings = []
    for df, ws in results:
        if not df.empty:
            dfs.append(df)
        if ws:
            warnings.extend(ws)
    if not dfs:
        return pd.DataFrame(), warnings
    combined = pd.concat(dfs, ignore_index=True)
    combined["Date_sort"] = pd.to_datetime(combined["Date"], errors="coerce")
    combined = combined.sort_values(["Date_sort","Bank","SourceFile"]).drop(columns=["Date_sort"]).reset_index(drop=True)
    combined.insert(0, "S.No", range(1, len(combined)+1))
    return combined, warnings

# ----------------------------
# PDF generation (grid table, IST time, multipage, dynamic widths)
# ----------------------------
def escape_html(s: str) -> str:
    return (s or "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def generate_pdf_bytes(rows: List[Dict[str,Any]], username: str, visible_cols: Optional[List[str]] = None) -> bytes:
    buf = BytesIO()
    page = landscape(A4)
    doc = SimpleDocTemplate(
        buf,
        pagesize=page,
        leftMargin=12*mm, rightMargin=12*mm, topMargin=12*mm, bottomMargin=12*mm
    )
    styles = getSampleStyleSheet()
    story = []

    now_ist_str = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%Y-%m-%d %H:%M IST")
    story.append(Paragraph(f"<b>Statement Export</b> &nbsp;&nbsp; User: <b>{escape_html(username)}</b> &nbsp;&nbsp; Generated: <b>{now_ist_str}</b>", styles["Normal"]))
    story.append(Spacer(1, 6))

    # Column order and visibility
    default_order = ["S.No","Date","Description","Amount","Type","Site Name","WorkNotes","Item","Bank"]
    if visible_cols:
        headers = [c for c in default_order if c in visible_cols] + [c for c in visible_cols if c not in default_order]
    else:
        headers = default_order

    data = [headers]
    for r in rows:
        row_vals = []
        for h in headers:
            val = r.get(h, "")
            if h in ["Description", "Site Name", "WorkNotes", "Bank", "Item"]:
                row_vals.append(Paragraph(escape_html(str(val)), styles["Normal"]))
            elif h == "Amount":
                try:
                    row_vals.append(f"{float(val):.2f}")
                except:
                    row_vals.append(str(val))
            else:
                row_vals.append(str(val))
        data.append(row_vals)

    printable_width = page[0] - (doc.leftMargin + doc.rightMargin)
    default_ratios = {
        "S.No": 0.06, "Date": 0.10, "Description": 0.26, "Amount": 0.08, "Type": 0.06,
        "Site Name": 0.15, "WorkNotes": 0.15, "Item": 0.10, "Bank": 0.09
    }
    ratios = [default_ratios.get(h, 0.10) for h in headers]
    s = sum(ratios) or 1.0
    ratios = [r/s for r in ratios]
    col_widths = [printable_width * r for r in ratios]

    PAGE_ROWS = 800
    start = 0
    while start < len(data):
        chunk = data[start:start+PAGE_ROWS] if start == 0 else [data[0]] + data[start:start+PAGE_ROWS]
        tbl = Table(chunk, colWidths=col_widths, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("GRID", (0,0), (-1,-1), 0.5, colors.black),
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#f0f0f0")),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("FONTSIZE", (0,0), (-1,-1), 8),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("ALIGN", (headers.index("S.No") if "S.No" in headers else 0,1),
                      (headers.index("S.No") if "S.No" in headers else 0,-1), "RIGHT"),
            ("ALIGN", (headers.index("Amount") if "Amount" in headers else 0,1),
                      (headers.index("Amount") if "Amount" in headers else 0,-1), "RIGHT"),
            ("ALIGN", (headers.index("Type") if "Type" in headers else 0,1),
                      (headers.index("Type") if "Type" in headers else 0,-1), "CENTER"),
        ]))
        story.append(tbl)
        start += PAGE_ROWS
        if start < len(data):
            story.append(PageBreak())

    
    # --- Totals row (Cr/Dr) ---
    try:
        total_cr = sum(float(r.get("Amount", 0) or 0) for r in rows if str(r.get("Type","")).strip().lower() == "cr")
        total_dr = sum(float(r.get("Amount", 0) or 0) for r in rows if str(r.get("Type","")).strip().lower() == "dr")
    except Exception:
        total_cr, total_dr = 0.0, 0.0

    # Build a single totals row aligned with headers
    totals_vals = []
    for h in headers:
        if h == "Type":
            totals_vals.append("Totals")
        elif h == "Amount":
            totals_vals.append(f"Cr: {total_cr:.2f} / Dr: {total_dr:.2f}")
        else:
            totals_vals.append("")
    totals_tbl = Table([totals_vals], colWidths=col_widths)
    totals_tbl.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#e8f4ff")),
        ("FONTNAME", (0,0), (-1,-1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("ALIGN", (headers.index("Amount") if "Amount" in headers else 0,0),
                  (headers.index("Amount") if "Amount" in headers else 0,0), "RIGHT"),
        ("ALIGN", (headers.index("Type") if "Type" in headers else 0,0),
                  (headers.index("Type") if "Type" in headers else 0,0), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ]))
    story.append(Spacer(1, 4))
    story.append(totals_tbl)
    doc.build(story)
    pdf = buf.getvalue()
    buf.close()
    return pdf

# ----------------------------
# AgGrid helpers
# ----------------------------
def order_columns_for_grid(df: pd.DataFrame) -> pd.DataFrame:
    desired = ["S.No","Date","Description","Amount","Type","Site Name","WorkNotes","Item","Bank"]
    present = [c for c in desired if c in df.columns]
    rest = [c for c in df.columns if c not in present and c != "SourceFile"]
    return df[present + rest]

def build_grid(df: pd.DataFrame, items_choices: List[str]) -> Tuple[GridOptionsBuilder, dict]:
    # Reorder columns and hide SourceFile
    df = order_columns_for_grid(df)

    g = GridOptionsBuilder.from_dataframe(df, enableValue=True, enableRowGroup=True, enablePivot=False)

    # Column configs in the required order
    g.configure_column("S.No", editable=False, width=90, pinned="left")
    g.configure_column("Date", editable=False, width=120, pinned="left")
    g.configure_column("Description", editable=True, width=420, wrapText=True, autoHeight=True)
    g.configure_column("Amount", editable=True, type=["numericColumn"], width=120, valueParser="Number(newValue)")
    g.configure_column("Type", editable=True, cellEditor="agSelectCellEditor", cellEditorParams={"values": ["Cr","Dr",""]}, width=100)
    g.configure_column("Site Name", editable=True, width=260, wrapText=True, autoHeight=True)
    g.configure_column("WorkNotes", editable=True, width=280, wrapText=True, autoHeight=True)
    # "rest" fields
    if "Item" in df.columns:
        g.configure_column("Item", editable=True, cellEditor="agSelectCellEditor", cellEditorParams={"values": items_choices}, width=140)
    if "Bank" in df.columns:
        g.configure_column("Bank", editable=False, width=180)
    if "SourceFile" in df.columns:
        g.configure_column("SourceFile", editable=False, width=180, hide=True)  # system field hidden

    # Conditional styling: ONLY Type column (Cr/Dr)
    type_style = JsCode("""
    function(params) {
      if (params.value === 'Cr') {
        return {'color': '#0b772d', 'fontWeight': '600'};
      } else if (params.value === 'Dr') {
        return {'color': '#b30000', 'fontWeight': '600'};
      }
      return null;
    }
    """)
    g.configure_column("Type", cellStyle=type_style)

    # General grid options
    g.configure_grid_options(
        domLayout="normal",
        rowSelection="single",
        animateRows=True,
        pagination=True,
        paginationAutoPageSize=False,
        paginationPageSize=25,
        suppressRowClickSelection=True,
        undoRedoCellEditing=True,
        enableCellTextSelection=True,
    )
    g.configure_grid_options(defaultColDef={
        "resizable": True,
        "sortable": True,
        "filter": True,
        "wrapText": True,
        "autoHeight": True,
        "editable": True,
    })

    grid_options = g.build()
    return g, grid_options, df

def files_signature(files) -> str:
    h = hashlib.sha256()
    for f in files:
        h.update(f.name.encode("utf-8"))
    return h.hexdigest()

# ----------------------------
# STREAMLIT UI
# ----------------------------
def main():
    st.title(PAGE_TITLE)
    st.markdown("Upload **CSV** or **Excel** statements. Click **Commit Changes** to sync edits.")

    # Sidebar
    st.sidebar.header("User / Settings")
    username = st.sidebar.text_input("Username (for context only)", value=st.session_state.get("username", ""))
    if username:
        st.session_state["username"] = username

    st.sidebar.markdown("---")
    st.sidebar.subheader("Manage Item choices")
    existing_items = get_items_from_db()
    with st.sidebar.form("items_form_add", clear_on_submit=True):
        new_item = st.text_input("Add new Item choice")
        add_it = st.form_submit_button("Add Item")
    if add_it and new_item:
        add_item_to_db(new_item)
        if hasattr(st, "rerun"): st.rerun()
        else: st.experimental_rerun()

    st.sidebar.markdown("**Existing choices**")
    to_remove = st.sidebar.multiselect("Select items to remove", options=existing_items, default=[])
    col_del_a, col_del_b = st.sidebar.columns([1,1])
    with col_del_a:
        remove_btn = st.button("Remove selected")
    with col_del_b:
        select_all = st.button("Select all")
    if select_all:
        st.session_state['__remove_all_toggle'] = True
        st.rerun() if hasattr(st, 'rerun') else st.experimental_rerun()
    if st.session_state.get('__remove_all_toggle'):
        to_remove = existing_items
        st.session_state.pop('__remove_all_toggle', None)
    if remove_btn:
        remove_item_from_db(to_remove)
        st.success(f"Removed: {', '.join(to_remove) if to_remove else 'None'}")
        if hasattr(st, "rerun"): st.rerun()
        else: st.experimental_rerun()

    uploaded_files = st.file_uploader(
        "Drop CSV/XLSX files (multiple allowed)",
        type=["csv","xlsx","xls"],
        accept_multiple_files=True
    )

    warns_holder = st.expander("Messages", expanded=False)

    if not uploaded_files:
        st.info("No files uploaded yet. Drop CSV or Excel files to start.")
        return

    sig = files_signature(uploaded_files)
    if "last_sig" not in st.session_state or st.session_state["last_sig"] != sig:
        with st.spinner("Parsing files..."):
            combined, warns = asyncio.run(parse_multiple_files(uploaded_files))
        st.session_state["last_sig"] = sig
        st.session_state["grid_df"] = combined.copy()
        st.session_state["committed_df"] = combined.copy()
        st.session_state["warns"] = warns
    else:
        warns = st.session_state.get("warns", [])

    with warns_holder:
        for w in st.session_state.get("warns", []):
            st.warning(w)

    if st.session_state["grid_df"].empty:
        st.error("No valid rows parsed. Please check your files/columns.")
        return

    st.success(f"Parsed {len(st.session_state['grid_df'])} rows from {len(uploaded_files)} file(s).")

    items_choices = get_items_from_db() or DEFAULT_ITEMS

    # Build and render AgGrid
    gob, grid_options, ordered_df = build_grid(st.session_state["grid_df"], items_choices)

    grid_response = AgGrid(
        ordered_df,
        gridOptions=grid_options,
        height=620,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        update_mode=GridUpdateMode.NO_UPDATE,  # don't push edits to Python automatically
        allow_unsafe_jscode=True,
        fit_columns_on_grid_load=True,
        enable_enterprise_modules=False,
        theme="balham",
    )

    # Action buttons (4 only)
    c1, c2, c3, c4 = st.columns([1,1,1,1])
    with c1:
        if st.button("Commit Changes"):
            current_df = pd.DataFrame(grid_response["data"])
            # numeric
            if "Amount" in current_df.columns:
                current_df["Amount"] = pd.to_numeric(current_df["Amount"], errors="coerce").fillna(0.0)
            # keep order
            current_df = order_columns_for_grid(current_df)
            st.session_state["grid_df"] = current_df.copy()
            st.session_state["committed_df"] = current_df.copy()
            st.success("Edits committed.")
    with c2:
        pdf_bytes = generate_pdf_bytes(
            st.session_state["committed_df"].to_dict(orient="records"),
            username or "anon",
            visible_cols=list(order_columns_for_grid(st.session_state["committed_df"]).columns)
        )
        st.download_button("Download PDF", data=pdf_bytes, file_name=f"statement_{(username or 'anon')}_{datetime.now(ZoneInfo('Asia/Kolkata')).strftime('%Y%m%d_%H%M%S')}_IST.pdf", mime="application/pdf")
    with c3:
        csv_bytes = order_columns_for_grid(st.session_state["committed_df"]).to_csv(index=False).encode("utf-8")
        st.download_button("Download CSV", data=csv_bytes, file_name="statement_final.csv", mime="text/csv")
    with c4:
        xls_buf = BytesIO()
        order_columns_for_grid(st.session_state["committed_df"]).to_excel(xls_buf, index=False, engine="openpyxl")
        xls_buf.seek(0)
        st.download_button("Download Excel", data=xls_buf, file_name="statement_final.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
