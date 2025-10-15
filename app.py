# app.py — Crushers parts by Asset (no Excel formulas)
# - Loads Crushers.xlsx from the repo
# - Asset dropdown (short label), optional Page filter (All or one)
# - Shows: Page | Item Number | QTY | Part Number | Part Name | InStk
# - Row checkboxes -> Generate Quote Request (CSV + optional DOCX)
# - Optional manual link: manuals/<Asset>.pdf in the repo

import io, re
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st

# ---------- Config: set paths relative to repo ----------
DATA_PATH    = Path("Crushers.xlsx")  # your workbook in the repo
SHEET_NAME   = None                   # or set to a sheet name if you prefer
MANUALS_DIR  = Path("manuals")        # put <Asset>.pdf here

# Optional: if you want a GitHub "Open manual" link
GH_OWNER   = ""        # e.g., "youruser"
GH_REPO    = ""        # e.g., "crushers-repo"
GH_BRANCH  = "main"    # e.g., "main"

st.set_page_config(page_title="Crushers Parts", layout="wide")
st.title("Crushers Parts")

# ---------- Helpers ----------
def pick_col(df, candidates):
    """Return the first matching column from candidate names (case/spacing tolerant)."""
    norm = {c: re.sub(r"\W+", "", str(c).lower()) for c in df.columns}
    cands = [re.sub(r"\W+", "", x.lower()) for x in candidates]
    # exact normalized match
    for c, n in norm.items():
        if n in cands:
            return c
    # relaxed contains match
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

@st.cache_data
def load_data():
    if SHEET_NAME:
        df = pd.read_excel(DATA_PATH, sheet_name=SHEET_NAME)
        used = SHEET_NAME
    else:
        xls = pd.ExcelFile(DATA_PATH)
        pick = 0
        for i, sh in enumerate(xls.sheet_names):
            df_try = pd.read_excel(xls, sheet_name=sh)
            if any("instk" in str(c).lower() or "in stock" in str(c).lower() for c in df_try.columns):
                df, used = df_try, sh
                break
        else:
            df = pd.read_excel(xls, sheet_name=0)
            used = xls.sheet_names[0]
    return df, used

# ---------- Load ----------
try:
    df, used_sheet = load_data()
    st.caption(f"Loaded sheet: **{used_sheet}** from `{DATA_PATH}`")
except Exception as e:
    st.error(f"Could not open `{DATA_PATH}`: {e}")
    st.stop()

# ---------- Column mapping ----------
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

# ---------- Top controls ----------
assets = sorted(df[col_asset].dropna().astype(str).unique().tolist())
sel_asset = st.selectbox("Asset", assets, index=0)

# Optional manual link/download
manual_name = sanitize_filename(sel_asset) + ".pdf"
manual_path = (MANUALS_DIR / manual_name)
man_cols = st.columns([1, 1, 6])
with man_cols[0]:
    if manual_path.exists():
        try:
            pdf_bytes = manual_path.read_bytes()
            st.download_button(
                "Download Manual (PDF)",
                data=pdf_bytes, file_name=manual_name, mime="application/pdf"
            )
        except Exception as e:
            st.warning(f"Manual found but couldn't read: {e}")
    else:
        st.caption(f"Place manual at `{MANUALS_DIR/{manual_name}}` for download")

with man_cols[1]:
    if GH_OWNER and GH_REPO:
        gh_url = f"https://raw.githubusercontent.com/{GH_OWNER}/{GH_REPO}/{GH_BRANCH}/{(MANUALS_DIR/manual_name).as_posix()}"
        st.link_button("Open Manual (GitHub raw)", gh_url, disabled=not (GH_OWNER and GH_REPO))

# Page filter: All pages or a specific page
pages_all = df.loc[df[col_asset].astype(str)==str(sel_asset), col_page].dropna()
pages_labels = sorted(pages_all.astype(str).unique().tolist())
opt = st.radio("Pages", ["All pages", "Specific page"], horizontal=True)
sel_page = None
if opt == "Specific page":
    sel_page = st.selectbox("Page", pages_labels, index=0)

# ---------- Filter & view ----------
f = df[col_asset].astype(str) == str(sel_asset)
if sel_page is not None:
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

# Sort by Page then Item Number (numeric where possible)
view["_page_sort"] = view["Page"].map(to_int_or_nan)
view["_item_sort"] = view["Item Number"].map(to_int_or_nan)
view = view.sort_values(by=["_page_sort", "Page", "_item_sort", "Item Number"]).drop(columns=["_page_sort", "_item_sort"])

# Add selection checkbox column
view_display = view.copy()
view_display.insert(0, "Select", False)

st.subheader(f"Parts for: {sel_asset}")
edited = st.data_editor(
    view_display,
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

# ---------- Quote request ----------
sel_rows = edited[edited["Select"]].drop(columns=["Select"])
c1, c2, c3 = st.columns([1, 1, 6])

# Pull model/serial from the asset's first row (if available)
hdr = df[df[col_asset].astype(str) == str(sel_asset)]
model_val  = str(hdr.iloc[0][col_model]) if (col_model and not hdr.empty and pd.notna(hdr.iloc[0][col_model])) else ""
serial_val = str(hdr.iloc[0][col_serial]) if (col_serial and not hdr.empty and pd.notna(hdr.iloc[0][col_serial])) else ""

with c1:
    st.write(f"**Selected rows:** {len(sel_rows)}")
with c2:
    gen = st.button("Generate Quote Request")

def quote_csv_bytes(dfq: pd.DataFrame) -> bytes:
    cols = ["Page", "Item Number", "Part Number", "Part Name", "QTY"]
    return dfq[cols].to_csv(index=False).encode("utf-8")

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
    doc.add_paragraph("")  # spacer

    table = doc.add_table(rows=1, cols=5)
    hdr = table.rows[0].cells
    hdr[0].text = "Page"; hdr[1].text = "Item Number"; hdr[2].text = "Part Number"; hdr[3].text = "Part Name"; hdr[4].text = "QTY"
    for _, r in dfq.iterrows():
        cells = table.add_row().cells
        cells[0].text = str(r["Page"])
        cells[1].text = str(r["Item Number"])
        cells[2].text = str(r["Part Number"])
        cells[3].text = str(r["Part Name"])
        cells[4].text = str(r["QTY"])
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

if gen:
    if sel_rows.empty:
        st.warning("No rows selected. Check the boxes in the left column.")
    else:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname_base = f"QuoteRequest_{sanitize_filename(sel_asset)}_{ts}"

        # CSV download
        csv_bytes = quote_csv_bytes(sel_rows)
        st.download_button("Download Quote (CSV)", data=csv_bytes,
                           file_name=f"{fname_base}.csv", mime="text/csv")

        # DOCX download (if python-docx is available)
        docx_bytes = quote_docx_bytes(sel_asset, model_val, serial_val, sel_rows)
        if docx_bytes:
            st.download_button("Download Quote (Word)", data=docx_bytes,
                               file_name=f"{fname_base}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        else:
            st.info("Install `python-docx` to enable Word download (see requirements.txt).")

# Footer hint
st.caption("Tip: Put PDFs in `manuals/` named exactly as the Asset (e.g., `Crushers A.pdf`) to enable the manual download/link.")
