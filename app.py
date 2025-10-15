# app.py — Crushers parts by Asset (no Excel formulas in sheet required)

import io, os, re
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st

# ==================== Config ====================
DATA_PATH   = Path("Crushers.xlsx")   # workbook in repo root
SHEET_NAME  = None                    # or set a specific sheet name
MANUALS_DIR = Path("manuals")         # put <Asset>.pdf here

# Optional GitHub raw link (set to enable a clickable link)
GH_OWNER  = ""       # e.g., "youruser"
GH_REPO   = ""       # e.g., "crushers-repo"
GH_BRANCH = "main"

st.set_page_config(page_title="Crushers Parts", layout="wide")
st.title("Crushers Parts")

# ==================== Helpers ====================
def pick_col(df, candidates):
    """Return first matching column (case/spacing tolerant)."""
    norm = {c: re.sub(r"\W+", "", str(c).lower()) for c in df.columns}
    cands = [re.sub(r"\W+", "", x.lower()) for x in candidates]
    for c, n in norm.items():
        if n in cands:
            return c
    for cand in cands:
        for c in df.columns:
            if cand in str(c).lower():
                return c
    return None

def to_int_or_nan(v):
    try:
        return int(float(v))
    except Exception:
        return pd.NA

def sanitize_filename(name: str) -> str:
    return re.sub(r"[^\w\-]+", "_", str(name)).strip("_")

def normalize_fname(s: str) -> str:
    # lower + unify separators for tolerant matching
    return re.sub(r"[^\w\-]+", "_", str(s)).strip("_").lower().replace("-", "_")

def find_manual_path(asset_name: str) -> Path | None:
    """Find <Asset>.pdf in manuals/ (case/spacing tolerant)."""
    if not MANUALS_DIR.exists():
        return None
    target = normalize_fname(asset_name)
    for p in MANUALS_DIR.glob("*.pdf"):
        if normalize_fname(p.stem) == target:
            return p
    return None

def github_raw_url(asset_name: str) -> str:
    if not (GH_OWNER and GH_REPO):
        return ""
    rel = (MANUALS_DIR / (sanitize_filename(asset_name) + ".pdf")).as_posix()
    return f"https://raw.githubusercontent.com/{GH_OWNER}/{GH_REPO}/{GH_BRANCH}/{rel}"

def to_excel_bytes(dfout: pd.DataFrame, sheet_name="Parts") -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as xw:
        dfout.to_excel(xw, index=False, sheet_name=sheet_name)
        # set nice column widths
        ws = xw.sheets.get(sheet_name)
        if ws:
            for idx, width in enumerate([10, 14, 8, 18, 36, 30], start=1):
                try:
                    ws.set_column(idx-1, idx-1, width)
                except Exception:
                    pass
    bio.seek(0)
    return bio.getvalue()

@st.cache_data
def load_data(_data_path: str, _sheet_name: str | None):
    if _sheet_name:
        df = pd.read_excel(_data_path, sheet_name=_sheet_name)
        used = _sheet_name
    else:
        xls = pd.ExcelFile(_data_path)
        df, used = None, None
        for sh in xls.sheet_names:
            t = pd.read_excel(xls, sheet_name=sh)
            if any("instk" in str(c).lower() or "in stock" in str(c).lower() for c in t.columns):
                df, used = t, sh
                break
        if df is None:
            df = pd.read_excel(xls, sheet_name=0)
            used = xls.sheet_names[0]
    return df, used

# ==================== Load sheet ====================
try:
    df, used_sheet = load_data(str(DATA_PATH), SHEET_NAME)
    st.caption(f"Loaded sheet: **{used_sheet}** from `{DATA_PATH}`")
except Exception as e:
    st.error(f"Could not open `{DATA_PATH}`: {e}")
    st.stop()

# ==================== Column mapping ====================
col_asset = pick_col(df, ["Asset", "Asset #", "Asset ID", "Equipment", "Equipment #", "Equipment ID", "Machine"])
col_page  = pick_col(df, ["Page", "Pg"])
col_item  = pick_col(df, ["Item Number", "Item #", "Item No", "Item", "Ref", "Ref #"])
col_qty   = pick_col(df, ["QTY", "Qty", "Quantity"])
col_pn    = pick_col(df, ["Part Number", "Part Numbers", "PN"])
col_name  = pick_col(df, ["Part Name", "Name", "Description", "Part Description"])
col_instk = pick_col(df, ["InStk", "In Stock", "LocAreaQty", "Loc Area Qty", "Location: Area,; Qty", "In_Stock"])
col_model = pick_col(df, ["Model", "Machine Model"])
col_serial= pick_col(df, ["Serial", "Serial Number", "S/N", "SN"])

needed = {
    "Asset": col_asset, "Page": col_page, "Item Number": col_item,
    "QTY": col_qty, "Part Number": col_pn, "Part Name": col_name, "InStk": col_instk
}
missing = [n for n, c in needed.items() if c is None]
if missing:
    st.error(f"Missing required column(s): {', '.join(missing)}\n\nColumns found: {list(df.columns)}")
    st.stop()

# ==================== Top controls (compact) ====================
assets = sorted(df[col_asset].dropna().astype(str).unique().tolist())

# Row: Asset | Page | Page filter | Search parts
c_asset, c_page, c_pagefilter, c_search = st.columns([1.6, 1.0, 0.9, 2.0])

with c_asset:
    sel_asset = st.selectbox("Asset", assets, index=0)

pages_for_asset = df.loc[df[col_asset].astype(str)==str(sel_asset), col_page].dropna().astype(str).unique().tolist()
pages_for_asset = sorted(pages_for_asset, key=lambda x: (to_int_or_nan(x), x))

with c_pagefilter:
    page_filter = st.text_input("Page filter", value="", placeholder="e.g. 42")

page_opts = [p for p in pages_for_asset if page_filter.lower() in p.lower()] if page_filter else pages_for_asset

with c_page:
    sel_page = st.selectbox("Page", ["All"] + page_opts, index=0)

with c_search:
    search_parts = st.text_input("Search parts", value="", placeholder="Part number or name")

# Manual links/downloads
manual_local  = find_manual_path(sel_asset)
manual_expect = (MANUALS_DIR / (sanitize_filename(sel_asset) + ".pdf")).as_posix()
cols_top = st.columns([1, 1, 6])
with cols_top[0]:
    if manual_local and manual_local.exists():
        try:
            pdf_bytes = manual_local.read_bytes()
            st.download_button("Manual PDF", data=pdf_bytes,
                               file_name=manual_local.name, mime="application/pdf")
        except Exception as e:
            st.warning(f"Manual found but couldn't read: {e}")
    else:
        st.caption(f"No local manual. Expected: `{manual_expect}`")
with cols_top[1]:
    gh_url = github_raw_url(sel_asset)
    if gh_url:
        try:
            st.link_button("Manual (GitHub raw)", gh_url)
        except Exception:
            st.markdown(f"[Manual (GitHub raw)]({gh_url})")

# ==================== View ====================
f = (df[col_asset].astype(str) == str(sel_asset))
if sel_page != "All":
    f &= (df[col_page].astype(str) == str(sel_page))

view = df.loc[f, [col_page, col_item, col_qty, col_pn, col_name, col_instk]].copy()
view.rename(columns={
    col_page: "Page",
    col_item: "Item Number",
    col_qty: "QTY",
    col_pn: "Part Number",
    col_name: "Part Name",
    col_instk: "InStk"
}, inplace=True)

if search_parts:
    sp = search_parts.strip().lower()
    view = view[
        view["Part Number"].astype(str).str.lower().str.contains(sp, na=False) |
        view["Part Name"].astype(str).str.lower().str.contains(sp, na=False)
    ]

# Sort (numeric-friendly)
view["_page_sort"] = view["Page"].map(to_int_or_nan)
view["_item_sort"] = view["Item Number"].map(to_int_or_nan)
view = view.sort_values(by=["_page_sort", "Page", "_item_sort", "Item Number"]).drop(columns=["_page_sort", "_item_sort"])

# Editable selection column
view_disp = view.copy()
view_disp.insert(0, "Select", False)

st.subheader(f"Parts for: {sel_asset}")
edited = st.data_editor(
    view_disp,
    hide_index=True,
    use_container_width=True,
    column_config={
        "Select": st.column_config.CheckboxColumn("Select"),
        "Page": st.column_config.TextColumn("Page"),
        "Item Number": st.column_config.TextColumn("Item Number"),
        "QTY": st.column_config.NumberColumn("QTY"),
        "Part Number": st.column_config.TextColumn("Part Number"),
        "Part Name": st.column_config.TextColumn("Part Name"),
        "InStk": st.column_config.TextColumn("InStk"),
    },
    disabled=["Page", "Item Number", "QTY", "Part Number", "Part Name", "InStk"]
)

# ==================== Quote request (persistent) ====================
sel_rows = edited[edited["Select"]].drop(columns=["Select"])

hdr = df[df[col_asset].astype(str) == str(sel_asset)]
model_val  = str(hdr.iloc[0][col_model])  if (col_model and not hdr.empty and pd.notna(hdr.iloc[0][col_model])) else ""
serial_val = str(hdr.iloc[0][col_serial]) if (col_serial and not hdr.empty and pd.notna(hdr.iloc[0][col_serial])) else ""

if "quote_ready" not in st.session_state:
    st.session_state.quote_ready = False
    st.session_state.quote_rows = None
    st.session_state.quote_asset = ""
    st.session_state.quote_model = ""
    st.session_state.quote_serial = ""
    st.session_state.quote_ts = ""

c1, c2, _ = st.columns([1, 1, 6])
with c1:
    st.write(f"**Selected rows:** {len(sel_rows)}")
with c2:
    gen = st.button("Generate Quote Request", type="primary")

def quote_xlsx_bytes(dfq: pd.DataFrame) -> bytes:
    cols = ["Page", "Item Number", "Part Number", "Part Name", "QTY"]
    return to_excel_bytes(dfq[cols], sheet_name="Quote")

def quote_docx_bytes(asset: str, model: str, serial: str, dfq: pd.DataFrame) -> bytes | None:
    try:
        from docx import Document
    except Exception:
        return None
    doc = Document()
    doc.add_heading("Quote Request", level=1)
    p = doc.add_paragraph()
    p.add_run("Asset: ").bold = True; p.add_run(str(asset))
    if model:
        p = doc.add_paragraph(); p.add_run("Model: ").bold = True; p.add_run(str(model))
    if serial:
        p = doc.add_paragraph(); p.add_run("Serial: ").bold = True; p.add_run(str(serial))
    doc.add_paragraph("")
    table = doc.add_table(rows=1, cols=5)
    hdr = table.rows[0].cells
    hdr[0].text = "Page"; hdr[1].text = "Item Number"; hdr[2].text = "Part Number"; hdr[3].text = "Part Name"; hdr[4].text = "QTY"
    for _, r in dfq.iterrows():
        row = table.add_row().cells
        row[0].text = str(r["Page"])
        row[1].text = str(r["Item Number"])
        row[2].text = str(r["Part Number"])
        row[3].text = str(r["Part Name"])
        row[4].text = str(r["QTY"])
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

# When clicked, persist selection + metadata for download on subsequent reruns
if gen:
    if sel_rows.empty:
        st.warning("No rows selected. Check the boxes in the left column.")
    else:
        st.session_state.quote_ready = True
        st.session_state.quote_rows = sel_rows[["Page", "Item Number", "Part Number", "Part Name", "QTY"]].copy()
        st.session_state.quote_asset = sel_asset
        st.session_state.quote_model = model_val
        st.session_state.quote_serial = serial_val
        st.session_state.quote_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.success("Quote prepared below. Use the download buttons.")

# Always offer: download the current filtered table (xlsx)
dl_cols = st.columns([1, 3, 6])
with dl_cols[0]:
    current_view_xlsx = to_excel_bytes(view, sheet_name="Parts")
    st.download_button("Download Current View (Excel)",
                       data=current_view_xlsx,
                       file_name=f"Parts_{sanitize_filename(sel_asset)}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# If a quote is ready, show persistent download buttons
if st.session_state.quote_ready and st.session_state.quote_rows is not None:
    base = f"QuoteRequest_{sanitize_filename(st.session_state.quote_asset)}_{st.session_state.quote_ts}"
    qxlsx = quote_xlsx_bytes(st.session_state.quote_rows)
    colq1, colq2, _ = st.columns([1, 1, 6])
    with colq1:
        st.download_button("Download Quote (Excel)",
                           data= qxlsx,
                           file_name=f"{base}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    # Optional Word
    qdocx = quote_docx_bytes(st.session_state.quote_asset,
                             st.session_state.quote_model,
                             st.session_state.quote_serial,
                             st.session_state.quote_rows)
    with colq2:
        if qdocx:
            st.download_button("Download Quote (Word)",
                               data=qdocx,
                               file_name=f"{base}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        else:
            st.caption("Install `python-docx` to enable Word download.")

# Debug helper (optional)
with st.expander("Debug: manuals present in repo"):
    try:
        files = sorted([p.name for p in MANUALS_DIR.glob("*.pdf")])
        st.write(files if files else "(none)")
    except Exception as e:
        st.write(f"Error listing manuals: {e}")

