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

# Try to import python-docx for Word export
try:
    from docx import Document
    from docx.shared import Pt
    DOCX_AVAILABLE = True
except Exception:
    DOCX_AVAILABLE = False

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
            # widths tuned for: Page | Item No | QTY | Part No | Part Name | InStk
            for idx, width in enumerate([10, 14, 8, 18, 36, 30], start=1):
                try:
                    ws.set_column(idx-1, idx-1, width)
                except Exception:
                    pass
    bio.seek(0)
    return bio.getvalue()

def word_bytes_from_rows(dfq: pd.DataFrame, df_all: pd.DataFrame,
                         col_asset: str, col_model: str | None, col_serial: str | None) -> bytes:
    """
    Create a .docx with:
      - Heading + Generated timestamp
      - Vendor (blank line)
      - Asset header block (Asset • Model • Serial) for each asset in dfq (top of doc)
      - Parts table: Part Number • Part Name • QTY  (+ 10 blank rows)
    """
    if not DOCX_AVAILABLE or dfq.empty:
        return b""

    # Build asset metadata from source sheet
    assets = pd.Series(sorted(dfq["Asset"].astype(str).unique().tolist()))
    meta_rows = []
    for a in assets:
        row = {"Asset": a, "Model": "", "Serial": ""}
        try:
            rows = df_all[df_all[col_asset].astype(str) == a]
            if not rows.empty:
                if col_model and pd.notna(rows.iloc[0][col_model]):  row["Model"]  = str(rows.iloc[0][col_model])
                if col_serial and pd.notna(rows.iloc[0][col_serial]): row["Serial"] = str(rows.iloc[0][col_serial])
        except Exception:
            pass
        meta_rows.append(row)
    meta_df = pd.DataFrame(meta_rows, columns=["Asset","Model","Serial"])

    # Parts table (only the 3 requested columns)
    cols_needed = ["Part Number","Part Name","QTY"]
    df_parts = pd.DataFrame(columns=cols_needed)
    # dfq already has these renamed columns from the view
    if not dfq.empty:
        for c in cols_needed:
            if c not in dfq.columns:
                dfq[c] = ""
        df_parts = dfq[cols_needed].copy()

    # Add 10 blank rows for manual add-ons
    df_parts = pd.concat(
        [df_parts, pd.DataFrame([{"Part Number": "", "Part Name": "", "QTY": ""} for _ in range(10)])],
        ignore_index=True
    )

    # Build doc
    doc = Document()
    try:
        doc.styles["Normal"].font.name = "Calibri"
        doc.styles["Normal"].font.size = Pt(10)
    except Exception:
        pass

    doc.add_heading("Quote Request", level=1)
    doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    # Vendor left blank for users to fill in
    doc.add_paragraph("Vendor: ______________________________")

    # Asset block at top
    if not meta_df.empty:
        doc.add_paragraph("")  # spacing
        for _, r in meta_df.iterrows():
            a = str(r["Asset"]) if pd.notna(r["Asset"]) else ""
            m = str(r["Model"]) if pd.notna(r["Model"]) else ""
            s = str(r["Serial"]) if pd.notna(r["Serial"]) else ""
            doc.add_paragraph(f"Asset: {a}    Model: {m}    Serial: {s}")

    doc.add_paragraph("")  # spacing
    doc.add_heading("Requested Parts", level=2)

    # Parts table (3 columns)
    t2 = doc.add_table(rows=1, cols=len(cols_needed))
    hdr2 = t2.rows[0].cells
    for i, col in enumerate(cols_needed):
        hdr2[i].text = col
    for _, r in df_parts.iterrows():
        cells = t2.add_row().cells
        cells[0].text = str(r["Part Number"]) if pd.notna(r["Part Number"]) else ""
        cells[1].text = str(r["Part Name"])   if pd.notna(r["Part Name"]) else ""
        cells[2].text = str(r["QTY"])         if pd.notna(r["QTY"]) else ""

    bio = io.BytesIO()
    doc.save(bio)
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

# Keep Asset internally, but don't display it
view = df.loc[f, [col_asset, col_page, col_item, col_qty, col_pn, col_name, col_instk]].copy()
view.rename(columns={
    col_asset: "Asset",
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

# ---------- Display table WITHOUT Asset ----------
view_disp = view.drop(columns=["Asset"]).copy()
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

# ==================== Quote cart (multi-page / multi-asset) ====================
# edited doesn't include Asset; map selected rows back to 'view' to reattach Asset
sel_rows_now_display = edited[edited["Select"]].drop(columns=["Select"]).copy()

# Merge back to recover Asset (use a robust key set)
merge_keys = ["Page","Item Number","Part Number","Part Name","QTY","InStk"]
sel_rows_now = pd.merge(sel_rows_now_display, view, on=merge_keys, how="left")

if "quote_cart" not in st.session_state:
    st.session_state.quote_cart = pd.DataFrame(columns=["Asset","Page","Item Number","Part Number","Part Name","QTY","InStk"])

def add_to_cart(df_add: pd.DataFrame):
    if df_add.empty:
        return
    cart = st.session_state.quote_cart
    combined = pd.concat([cart, df_add], ignore_index=True)
    # De-dupe by Asset + Page + Item Number + Part Number
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

# ---- Editable Cart (change QTY, remove items) ----
st.markdown("### Cart")
cart_df = st.session_state.quote_cart.copy()

if cart_df.empty:
    st.caption("No items yet. Select rows above and click **Add selected to cart**.")
else:
    # Build a user-facing cart view; keep Asset internally but show for context
    cart_view = cart_df.copy()
    cart_view.insert(0, "Remove", False)

    with st.form("cart_form", clear_on_submit=False):
        edited_cart = st.data_editor(
            cart_view[["Remove","Asset","Page","Item Number","Part Number","Part Name","QTY","InStk"]],
            hide_index=True,
            use_container_width=True,
            column_config={
                "Remove": st.column_config.CheckboxColumn("Remove", help="Check to remove selected rows"),
                "QTY": st.column_config.NumberColumn("QTY", step=1, min_value=0),
                "Asset": st.column_config.TextColumn("Asset"),
                "Page": st.column_config.TextColumn("Page"),
                "Item Number": st.column_config.TextColumn("Item Number"),
                "Part Number": st.column_config.TextColumn("Part Number"),
                "Part Name": st.column_config.TextColumn("Part Name"),
                "InStk": st.column_config.TextColumn("InStk"),
            },
            disabled=["Asset","Page","Item Number","Part Number","Part Name","InStk"],
            key="cart_editor",
        )
        c_s, c_r, c_c = st.columns([1.2, 1.2, 1.2])
        save_qty = c_s.form_submit_button("💾 Save Qty", use_container_width=True)
        remove_it = c_r.form_submit_button("🗑️ Remove selected", use_container_width=True)
        clear_it  = c_c.form_submit_button("🧼 Clear cart", use_container_width=True)

    if save_qty and "QTY" in edited_cart.columns:
        # Persist QTY edits back to the session cart (align by index)
        st.session_state.quote_cart.loc[edited_cart.index, "QTY"] = edited_cart["QTY"].values
        st.success("Saved quantities.")

    if remove_it:
        try:
            idx = edited_cart.index[edited_cart["Remove"] == True]
        except Exception:
            idx = []
        if len(idx):
            st.session_state.quote_cart = st.session_state.quote_cart.drop(index=idx).reset_index(drop=True)
            st.rerun()

    if clear_it:
        st.session_state.quote_cart = st.session_state.quote_cart.iloc[0:0].copy()
        st.rerun()

with st.expander("View raw cart data"):
    st.dataframe(st.session_state.quote_cart, hide_index=True, use_container_width=True)

# Build rows for export: prefer cart; else current selection
rows_for_quote = st.session_state.quote_cart.copy() if not st.session_state.quote_cart.empty else sel_rows_now.copy()

# ==================== Downloads ====================
# 1) Generate Quote (Word) — .docx (Parts table = PN | Name | QTY; +10 blank rows; top shows Vendor blank + Asset/Model/Serial)
if DOCX_AVAILABLE:
    word_bytes = word_bytes_from_rows(rows_for_quote, df, col_asset, col_model, col_serial) if not rows_for_quote.empty else b""
    base_name = f"QuoteRequest_{sanitize_filename(sel_asset)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    st.download_button(
        "Generate Quote (Word)",
        data=word_bytes if rows_for_quote is not None and not rows_for_quote.empty else b"",
        file_name=f"{base_name}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        disabled=rows_for_quote is None or rows_for_quote.empty
    )
else:
    st.warning("Word export requires `python-docx`. Add it to requirements.txt.")

# 2) Download Current View (Excel) — exports what you see (NO Asset column)
current_view_xlsx = to_excel_bytes(view_disp.drop(columns=["Select"]), sheet_name="Parts")
st.download_button("Download Current View (Excel)",
                   data=current_view_xlsx,
                   file_name=f"Parts_{sanitize_filename(sel_asset)}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


