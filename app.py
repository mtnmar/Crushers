# app.py — Parts by Asset (no Excel formulas)
# - Reads an .xlsx file (upload or local path)
# - Select Asset -> shows Page | Item Number | QTY | Part Name | InStk
# - Uses only sheet values; no calculations/formulas required
# - Download filtered table as Excel/CSV

import re, io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Parts by Asset", layout="wide")

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

def to_int_or_nan(s):
    try:
        return int(float(s))
    except Exception:
        return pd.NA

@st.cache_data
def load_excel(file_bytes: bytes | None, local_path: str | None):
    if file_bytes:
        xls = pd.ExcelFile(io.BytesIO(file_bytes))
    else:
        xls = pd.ExcelFile(local_path)
    # pick the first sheet that contains an "InStk"-like column
    for sh in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sh)
        if any("instk" in str(c).lower() or "in stock" in str(c).lower() for c in df.columns):
            return df, sh
    # fallback to first sheet
    return pd.read_excel(xls, sheet_name=0), xls.sheet_names[0]

# ---------- UI: File input ----------
st.title("Parts by Asset")

left, right = st.columns([2,1])
with left:
    up = st.file_uploader("Upload your Excel file (.xlsx) or leave blank to use a local file path below", type=["xlsx"])
with right:
    local_path = st.text_input("Local file path (optional)", value="Crushers.xlsx")

if not up and not local_path:
    st.info("Upload a file or enter a local file path (e.g., Crushers.xlsx).")
    st.stop()

try:
    df, used_sheet = load_excel(up.read() if up else None, local_path if not up else None)
except Exception as e:
    st.error(f"Could not open the workbook: {e}")
    st.stop()

st.caption(f"Loaded sheet: **{used_sheet}**")

# ---------- Column mapping (best effort) ----------
col_asset = pick_col(df, ["Asset", "Asset #", "Asset ID", "Equipment", "Equipment #", "Equipment ID", "Machine"])
col_page  = pick_col(df, ["Page", "Pg"])
col_item  = pick_col(df, ["Item Number", "Item #", "Item No", "Item", "Ref", "Ref #"])
col_qty   = pick_col(df, ["QTY", "Qty", "Quantity"])
col_name  = pick_col(df, ["Part Name", "Name", "Description", "Part Description"])
col_instk = pick_col(df, ["InStk", "In Stock", "LocAreaQty", "Loc Area Qty", "Location: Area,; Qty", "In_Stock"])

missing = [n for n, c in {
    "Asset": col_asset, "Page": col_page, "Item Number": col_item,
    "QTY": col_qty, "Part Name": col_name, "InStk": col_instk
}.items() if c is None]

if missing:
    st.error(f"Missing required column(s) in the sheet: {', '.join(missing)}\n\n"
             f"Found columns:\n{list(df.columns)}")
    st.stop()

# ---------- Asset selection ----------
assets = sorted(df[col_asset].dropna().astype(str).unique().tolist())
sel = st.selectbox("Choose Asset", assets, index=0)

# Filter and shape the view
view = df[df[col_asset].astype(str) == str(sel)].copy()

# Keep only the requested columns and rename for display
view = view[[col_page, col_item, col_qty, col_name, col_instk]].rename(columns={
    col_page: "Page",
    col_item: "Item Number",
    col_qty: "QTY",
    col_name: "Part Name",
    col_instk: "InStk"
})

# Sort: Page (numeric if possible), then Item Number (numeric if possible)
view["_page_sort"] = view["Page"].map(to_int_or_nan)
view["_item_sort"] = view["Item Number"].map(to_int_or_nan)
view = view.sort_values(by=["_page_sort", "Page", "_item_sort", "Item Number"]).drop(columns=["_page_sort", "_item_sort"])

st.subheader(f"Parts for Asset: {sel}")
st.dataframe(view, hide_index=True, use_container_width=True)

# ---------- Downloads ----------
def to_excel_bytes(dfout: pd.DataFrame) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as xw:
        dfout.to_excel(xw, index=False, sheet_name="Parts")
        for idx, width in enumerate([12, 16, 8, 36, 30], start=1):  # nice widths
            try:
                xw.sheets["Parts"].set_column(idx-1, idx-1, width)
            except Exception:
                pass
    bio.seek(0)
    return bio.getvalue()

c1, c2 = st.columns(2)
with c1:
    st.download_button("Download CSV", data=view.to_csv(index=False).encode("utf-8"),
                       file_name=f"{str(sel).strip()}_parts.csv", mime="text/csv")
with c2:
    st.download_button("Download Excel", data=to_excel_bytes(view),
                       file_name=f"{str(sel).strip()}_parts.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
