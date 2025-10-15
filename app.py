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
        ws = xw.sheets.get(sheet_name)
        if ws:
            for idx, width in enumerate([10, 14, 8, 18, 36, 30], start=1):
                try:
                    ws.set_column(idx-1, idx-1, width)
                except Exception:
                    pass
    bio.seek(0)
    return bio.getvalue()

def quote_xlsx_bytes(asset: str, model: str, serial: str, dfq: pd.DataFrame) -> bytes:
    """Two-sheet Excel: Quote rows + Header (Asset/Model/Serial)."""
    cols = ["Page", "Item Number", "Part Number", "Part Name", "QTY", "InStk", "Asset"]
    dfq2 = dfq.copy()
    for c in cols:
        if c not in dfq2.columns:
            dfq2[c] = ""
    order = ["Page", "Item Number", "Part Number", "Part Name", "QTY", "InStk", "Asset"]
    dfq2 = dfq2[order]

    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as xw:
        # Header sheet
        hdr = pd.DataFrame({
            "Field": ["Asset", "Model", "Serial", "Generated"],
            "Value": [asset, model, serial, datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
        })
        hdr.to_excel(xw, index=False, sheet_name="Header")
        ws_hdr = xw.sheets["Header"]
        try:
            ws_hdr.set_column(0, 0, 12)
            ws_hdr.set_column(1, 1, 40)
        except Exception:
            pass

        # Quote sheet
        dfq2.to_excel(xw, index=False, sheet_name="Quote")
        ws_q = xw.sheets["Quote"]
        try:
            for idx, width in enumerate([10, 14, 18, 36, 8, 30, 18], start=1):
                ws_q.set_column(idx-1, idx-1, width)
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

view = df.loc[f, [col_page, col_item, col_qty, col_pn, col_name, col_instk, col_asset]].copy()
view.rename(columns={
    col_page: "Page",
    col_item: "Item Number",
    col_qty: "QTY",
    col_pn: "Part Number",
    col_name: "Part Name",
    col_instk: "InStk",
    col_asset: "Asset"
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
        "Asset": st.column_config.TextColumn("Asset"),
    },
    disabled=["Page", "Item Number", "QTY", "Part Number", "Part Name", "InStk", "Asset"]
)

# ==================== Quote cart (multi-page / multi-asset) ====================
sel_rows_now = edited[edited["Select"]].drop(columns=["Select"]).copy()

if "quote_cart" not in st.session_state:
    st.session_state.quote_cart = pd.DataFrame(columns=["Page","Item Number","Part Number","Part Name","QTY","InStk","Asset"])

def add_to_cart(df_add: pd.DataFrame):
    if df_add.empty:
        return
    cart = st.session_state.quote_cart
    combined = pd.concat([cart, df_add], ignore_index=True)
    # de-dupe by Asset + Page + Item Number + Part Number
    dedupe_keys = ["Asset","Page","Item Number","Part Number"]
    combined = combined.sort_values(dedupe_keys).drop_duplicates(subset=dedupe_keys, keep="last")
    st.session_state.quote_cart = combined

c_cart1, c_cart2, c_cart3, _ = st.columns([1.2, 1.6, 1.2, 5])
with c_cart1:
    if st.button("Add selected to cart", type="primary"):
        add_to_cart(sel_rows_now)
with c_cart2:
    if st.button("Clear cart"):
        st.session_state.quote_cart = st.session_state.quote_cart.iloc[0:0].copy()
with c_cart3:
    st.caption(f"Cart items: **{len(st.session_state.quote_cart)}**")

with st.expander("View quote cart"):
    st.dataframe(st.session_state.quote_cart, hide_index=True, use_container_width=True)

# ==================== Generate (direct download) ====================
# Build the rows to export: prefer cart if not empty; otherwise current selection
if not st.session_state.quote_cart.empty:
    rows_for_quote = st.session_state.quote_cart.copy()
else:
    rows_for_quote = sel_rows_now.copy()

# Pull model/serial (first row under current asset, if present)
hdr = df[df[col_asset].astype(str) == str(sel_asset)]
model_val  = str(hdr.iloc[0][col_model])  if (col_model and not hdr.empty and pd.notna(hdr.iloc[0][col_model])) else ""
serial_val = str(hdr.iloc[0][col_serial]) if (col_serial and not hdr.empty and pd.notna(hdr.iloc[0][col_serial])) else ""

ts = datetime.now().strftime("%Y%m%d_%H%M%S")
base_name = f"QuoteRequest_{sanitize_filename(sel_asset)}_{ts}"
xlsx_bytes_now = quote_xlsx_bytes(sel_asset, model_val, serial_val, rows_for_quote) if not rows_for_quote.empty else b""

# One-click download (Excel). If nothing selected and cart empty, disable.
st.download_button(
    "Generate Quote (.xlsx)",
    data=xlsx_bytes_now if xlsx_bytes_now else b"No items selected",
    file_name=f"{base_name}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    disabled=rows_for_quote.empty
)

# Also offer "Download Current View (Excel)" (whatever you're looking at)
current_view_xlsx = to_excel_bytes(view, sheet_name="Parts")
st.download_button("Download Current View (Excel)",
                   data=current_view_xlsx,
                   file_name=f"Parts_{sanitize_filename(sel_asset)}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


