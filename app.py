# import streamlit as st
# import pandas as pd
# from io import BytesIO

# st.set_page_config(page_title="S.BAG Cleaner", page_icon="💎", layout="centered")

# st.title("💎 S.BAG Report Cleaner")
# st.markdown("Upload your **S.BAG (2)** format file to convert it into the clean **SBAG11** output format.")

# # ── Column mapping from S.BAG(2) → SBAG11 ──────────────────────────────────
# # S.BAG(2) real header row is row index 1 (0-indexed), data starts from row 2
# # SBAG11 columns: Item, Shape, Quality, Color, Grade, (weight col 5), MemoOut, Udf7

# SOURCE_COL_MAP = {
#     "Item":     "Item",
#     "Shape":    "Shape",
#     "Quality":  "Quality",
#     "Color":    "Color",
#     "Grade":    "Grade",
#     "MemoOut":  "MemoOut",
#     "Udf7":     "Udf7",
# }
# # Weight column in S.BAG(2) is col index 8 (Qty / Pcs header in row 0), header name is NaN
# # In SBAG11 col index 5 has no header (NaN), it holds weight value
# WEIGHT_COL_IDX = 8   # 0-based index in source

# TARGET_COLUMNS = ["Item", "Shape", "Quality", "Color", "Grade", None, "MemoOut", "Udf7"]


# def clean_sbag(raw_df: pd.DataFrame) -> pd.DataFrame:
#     """Convert raw S.BAG(2) DataFrame (no header parsed) to SBAG11 format."""
#     # Row 1 = real headers, Row 2+ = data
#     headers = raw_df.iloc[1].tolist()
#     data = raw_df.iloc[2:].reset_index(drop=True)
#     data.columns = headers

#     # Build output dataframe
#     out = pd.DataFrame()
#     out["Item"]    = data["Item"].str.strip() if "Item" in data.columns else ""
#     out["Shape"]   = data["Shape"].str.strip() if "Shape" in data.columns else ""
#     out["Quality"] = data["Quality"].str.strip() if "Quality" in data.columns else ""
#     out["Color"]   = data["Color"].str.strip() if "Color" in data.columns else ""
#     out["Grade"]   = data["Grade"].str.strip() if "Grade" in data.columns else ""

#     # Weight column (col index 8 in source, NaN header → Qty/Pcs row 0)
#     # Accessed by positional index since header is NaN
#     weight_series = raw_df.iloc[2:, WEIGHT_COL_IDX].reset_index(drop=True)
#     out[None] = pd.to_numeric(weight_series, errors="coerce")

#     out["MemoOut"] = pd.to_numeric(data.get("MemoOut", 0), errors="coerce").fillna(0)
#     out["Udf7"]    = data.get("Udf7", "").astype(str).str.strip()

#     # Rename the None column to empty string for Excel output
#     out.columns = ["Item", "Shape", "Quality", "Color", "Grade", "", "MemoOut", "Udf7"]

#     # Drop fully empty rows
#     out = out.dropna(subset=["Item"]).reset_index(drop=True)
#     out = out[out["Item"].astype(str).str.strip() != ""]

#     return out


# def to_excel(df: pd.DataFrame) -> bytes:
#     """Write cleaned DataFrame to Excel bytes matching SBAG11 style."""
#     buf = BytesIO()
#     with pd.ExcelWriter(buf, engine="openpyxl") as writer:
#         df.to_excel(writer, index=False, sheet_name="Sheet1")

#         from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
#         ws = writer.sheets["Sheet1"]

#         header_fill = PatternFill("solid", start_color="4472C4", end_color="4472C4")
#         header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
#         data_font   = Font(name="Arial", size=10)
#         thin        = Side(style="thin", color="BFBFBF")
#         border      = Border(left=thin, right=thin, top=thin, bottom=thin)
#         center      = Alignment(horizontal="center", vertical="center")

#         col_widths = [14, 18, 10, 8, 8, 8, 10, 8]

#         for col_idx, (cell, width) in enumerate(zip(ws[1], col_widths), start=1):
#             cell.font      = header_font
#             cell.fill      = header_fill
#             cell.alignment = center
#             cell.border    = border
#             ws.column_dimensions[cell.column_letter].width = width

#         for row in ws.iter_rows(min_row=2):
#             for cell in row:
#                 cell.font      = data_font
#                 cell.alignment = Alignment(horizontal="center", vertical="center")
#                 cell.border    = border

#     return buf.getvalue()


# # ── UI ───────────────────────────────────────────────────────────────────────

# uploaded = st.file_uploader(
#     "📂 Upload S.BAG (2) file (.xlsx)",
#     type=["xlsx"],
#     accept_multiple_files=False,
# )

# if uploaded:
#     try:
#         raw = pd.read_excel(uploaded, sheet_name="Sheet1", header=None)

#         with st.expander("🔍 Raw Input Preview", expanded=False):
#             st.dataframe(raw, use_container_width=True)

#         cleaned = clean_sbag(raw)

#         st.success(f"✅ Cleaned successfully — **{len(cleaned)} rows** ready for download.")
#         st.subheader("📋 Cleaned Output (SBAG11 Format)")
#         st.dataframe(cleaned, use_container_width=True)

#         excel_bytes = to_excel(cleaned)

#         st.download_button(
#             label="⬇️ Download SBAG11.xlsx",
#             data=excel_bytes,
#             file_name="SBAG11.xlsx",
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         )

#     except Exception as e:
#         st.error(f"❌ Error processing file: {e}")
#         st.exception(e)

# else:
#     st.info("👆 Upload an **S.BAG (2)** format Excel file to get started.")

#     with st.expander("ℹ️ What does this tool do?"):
#         st.markdown("""
# | S.BAG (2) Column | → | SBAG11 Column |
# |---|---|---|
# | Item | → | Item |
# | Shape | → | Shape |
# | Quality | → | Quality |
# | Color | → | Color |
# | Grade | → | Grade |
# | Qty/Wt (col 8) | → | *(weight)* |
# | MemoOut | → | MemoOut |
# | Udf7 | → | Udf7 |

# The tool strips extra header rows, removes blank rows, and exports a clean `.xlsx` file.
#         """)


import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile

st.set_page_config(page_title="Fancy Layout Reportn Automation", page_icon="💎", layout="wide")

WEIGHT_COL_IDX = 8  # 0-based positional index in source file


def clean_sbag(raw_df: pd.DataFrame) -> pd.DataFrame:
    headers = raw_df.iloc[1].tolist()
    data = raw_df.iloc[2:].reset_index(drop=True)
    data.columns = headers

    out = pd.DataFrame()
    out["Item"]    = data["Item"].astype(str).str.strip() if "Item" in data.columns else ""
    out["Shape"]   = data["Shape"].astype(str).str.strip() if "Shape" in data.columns else ""
    out["Quality"] = data["Quality"].astype(str).str.strip() if "Quality" in data.columns else ""
    out["Color"]   = data["Color"].astype(str).str.strip() if "Color" in data.columns else ""
    out["Grade"]   = data["Grade"].astype(str).str.strip() if "Grade" in data.columns else ""
    out[""]        = pd.to_numeric(raw_df.iloc[2:, WEIGHT_COL_IDX].reset_index(drop=True), errors="coerce")
    out["MemoOut"] = pd.to_numeric(data.get("MemoOut", 0), errors="coerce").fillna(0)
    out["Udf7"]    = data.get("Udf7", pd.Series([""] * len(data))).astype(str).str.strip()

    out = out[out["Item"].str.strip().ne("") & out["Item"].str.lower().ne("nan")]
    return out.reset_index(drop=True)


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        ws = writer.sheets["Sheet1"]

        header_fill = PatternFill("solid", start_color="4472C4", end_color="4472C4")
        header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        data_font   = Font(name="Arial", size=10)
        thin        = Side(style="thin", color="BFBFBF")
        border      = Border(left=thin, right=thin, top=thin, bottom=thin)
        col_widths  = [14, 18, 10, 8, 8, 8, 10, 8]

        for cell, width in zip(ws[1], col_widths):
            cell.font      = header_font
            cell.fill      = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border    = border
            ws.column_dimensions[cell.column_letter].width = width

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font      = data_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border    = border
    return buf.getvalue()


def make_zip(results: list[dict]) -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in results:
            zf.writestr(r["out_name"], r["excel_bytes"])
    return buf.getvalue()


# ── UI ────────────────────────────────────────────────────────────────────────

st.title("💎 S.BAG Report Cleaner")
st.markdown("Upload **one or more** Massy Data File `.xlsx` files — all will be cleaned and returned in Correct Cleaned format.")

uploaded_files = st.file_uploader(
    "📂 Upload Massy Data Files (.xlsx) — you can select multiple",
    type=["xlsx"],
    accept_multiple_files=True,
)

if uploaded_files:
    results   = []
    all_dfs   = []
    errors    = []

    st.markdown(f"### Processing **{len(uploaded_files)}** file(s)…")
    progress = st.progress(0)

    for i, f in enumerate(uploaded_files):
        try:
            raw     = pd.read_excel(f, sheet_name="Sheet1", header=None)
            cleaned = clean_sbag(raw)
            cleaned.insert(0, "_source_file", f.name)  # tag for combined view
            all_dfs.append(cleaned)

            out_name = f.name.replace(".xlsx", "").replace(" ", "_") + "_SBAG11.xlsx"
            excel_bytes = to_excel_bytes(cleaned.drop(columns=["_source_file"]))
            results.append({"orig_name": f.name, "out_name": out_name,
                            "cleaned": cleaned, "excel_bytes": excel_bytes, "rows": len(cleaned)})
        except Exception as e:
            errors.append({"name": f.name, "error": str(e)})

        progress.progress((i + 1) / len(uploaded_files))

    progress.empty()

    # ── Summary bar ──────────────────────────────────────────────────────────
    col1, col2, col3 = st.columns(3)
    col1.metric("Files Uploaded", len(uploaded_files))
    col2.metric("Cleaned Successfully", len(results))
    col3.metric("Errors", len(errors))

    # ── Errors ───────────────────────────────────────────────────────────────
    if errors:
        st.error("⚠️ Some files could not be processed:")
        for e in errors:
            st.markdown(f"- **{e['name']}**: `{e['error']}`")

    # ── Per-file results ──────────────────────────────────────────────────────
    if results:
        st.markdown("---")
        st.subheader("📋 Cleaned Files")

        for r in results:
            with st.expander(f"✅  {r['orig_name']}  →  **{r['rows']} rows**", expanded=len(results) == 1):
                st.dataframe(r["cleaned"].drop(columns=["_source_file"]), use_container_width=True)
                st.download_button(
                    label=f"⬇️ Download {r['out_name']}",
                    data=r["excel_bytes"],
                    file_name=r["out_name"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{r['out_name']}",
                )

        # ── Combined download (all rows merged) ───────────────────────────────
        st.markdown("---")
        st.subheader("📦 Download Options")

        dl_col1, dl_col2 = st.columns(2)

        # Option 1 – ZIP of individual files
        zip_bytes = make_zip(results)
        dl_col1.download_button(
            label=f"🗜️ Download All as ZIP ({len(results)} files)",
            data=zip_bytes,
            file_name="SBAG11_all.zip",
            mime="application/zip",
            use_container_width=True,
        )

        # Option 2 – Single merged Excel (each file = one sheet)
        merged_buf = BytesIO()
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        with pd.ExcelWriter(merged_buf, engine="openpyxl") as writer:
            for r in results:
                sheet_name = r["orig_name"][:31].replace("/", "-").replace("\\", "-").replace("*", "").replace("?", "").replace("[", "").replace("]", "").replace(":", "")
                df_sheet = r["cleaned"].drop(columns=["_source_file"])
                df_sheet.to_excel(writer, index=False, sheet_name=sheet_name)
                ws = writer.sheets[sheet_name]
                header_fill = PatternFill("solid", start_color="4472C4", end_color="4472C4")
                col_widths = [14, 18, 10, 8, 8, 8, 10, 8]
                for cell, width in zip(ws[1], col_widths):
                    cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    ws.column_dimensions[cell.column_letter].width = width
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        cell.font = Font(name="Arial", size=10)
                        cell.alignment = Alignment(horizontal="center", vertical="center")

        dl_col2.download_button(
            label=f"📊 Download Merged Excel (all sheets)",
            data=merged_buf.getvalue(),
            file_name="Cleaned_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

else:
    st.info("👆 Upload one or more **Massy Data Files** Excel files to get started.")
    with st.expander("ℹ️ Column mapping reference"):
        st.markdown("""
| S.BAG (2) Source | → | SBAG11 Output |
|---|---|---|
| Item | → | Item |
| Shape | → | Shape |
| Quality | → | Quality |
| Color | → | Color |
| Grade | → | Grade |
| Col 8 (Qty/Wt) | → | *(weight)* |
| MemoOut | → | MemoOut |
| Udf7 | → | Udf7 |
        """)

# import streamlit as st
# import pandas as pd
# from io import BytesIO
# import zipfile, shutil, os

# st.set_page_config(page_title="Diamond Report Tool", page_icon="💎", layout="wide")

# # ── Constants ─────────────────────────────────────────────────────────────────
# WEIGHT_COL_IDX = 8   # positional index in S.BAG(2) for weight column

# # Shape → Sheet name map in Fancy Layout Report
# SHAPE_TO_SHEET = {
#     "ASSCHER":          "ASSCHER",
#     "EMERALD":          "EMERALD",
#     "HEART":            "HEART",
#     "LONG RADIANT":     "LONG RADIANT",
#     "MARQUISE":         "MARQUISE",
#     "OVAL":             "OVAL",
#     "PEAR":             "PEAR",
#     "PRINCESS":         "PRINCESS",
#     "SQUARE CUSHION":   "SQUARE CUSHION",
#     "STRAIGHT BAG":     "STRAIGHT BAG",
#     "TAPERED BAGUETTE": "TAPERED BAGUETTE",
#     "SQUARE RADIANT":   "SQUARE RADIANT",
#     "OLD EUROPEAN":     "ROUND OLD EUROPEAN ",
#     "LONG CUSHION":     "LONG CUSHION",
#     "OLD MINER CUSHION":"OLD MINE CUSHION",
#     "TRILLION":         "TRILLION",
# }

# # Quality → Sheet name map
# QUALITY_TO_SHEET = {
#     "CVD":  "TO ORDER CVD",
#     "HPHT": "TO ORDER HPHT",
# }

# # ── S.BAG(2) Cleaner ──────────────────────────────────────────────────────────
# def clean_sbag(raw_df: pd.DataFrame) -> pd.DataFrame:
#     headers = raw_df.iloc[1].tolist()
#     data    = raw_df.iloc[2:].reset_index(drop=True)
#     data.columns = headers
#     out = pd.DataFrame()
#     out["Item"]    = data["Item"].astype(str).str.strip()
#     out["Shape"]   = data["Shape"].astype(str).str.strip()
#     out["Quality"] = data["Quality"].astype(str).str.strip()
#     out["Color"]   = data["Color"].astype(str).str.strip()
#     out["Grade"]   = data["Grade"].astype(str).str.strip()
#     out[""]        = pd.to_numeric(raw_df.iloc[2:, WEIGHT_COL_IDX].reset_index(drop=True), errors="coerce")
#     out["MemoOut"] = pd.to_numeric(data.get("MemoOut", 0), errors="coerce").fillna(0)
#     out["Udf7"]    = data.get("Udf7", pd.Series([""] * len(data))).astype(str).str.strip()
#     out = out[out["Item"].str.strip().ne("") & out["Item"].str.lower().ne("nan")]
#     return out.reset_index(drop=True)


# def to_excel_bytes_simple(df: pd.DataFrame) -> bytes:
#     from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
#     buf = BytesIO()
#     with pd.ExcelWriter(buf, engine="openpyxl") as writer:
#         df.to_excel(writer, index=False, sheet_name="Sheet1")
#         ws = writer.sheets["Sheet1"]
#         hfill = PatternFill("solid", start_color="4472C4", end_color="4472C4")
#         hfont = Font(name="Arial", bold=True, color="FFFFFF", size=10)
#         thin  = Side(style="thin", color="BFBFBF")
#         brd   = Border(left=thin, right=thin, top=thin, bottom=thin)
#         widths = [14, 18, 10, 8, 8, 8, 10, 8]
#         for cell, w in zip(ws[1], widths):
#             cell.font = hfont; cell.fill = hfill
#             cell.alignment = Alignment(horizontal="center", vertical="center")
#             cell.border = brd
#             ws.column_dimensions[cell.column_letter].width = w
#         for row in ws.iter_rows(min_row=2):
#             for cell in row:
#                 cell.font = Font(name="Arial", size=10)
#                 cell.alignment = Alignment(horizontal="center", vertical="center")
#                 cell.border = brd
#     return buf.getvalue()


# def make_zip(results):
#     buf = BytesIO()
#     with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
#         for r in results:
#             zf.writestr(r["out_name"], r["excel_bytes"])
#     return buf.getvalue()


# # ── Fancy Layout Report: paste data into correct sheets ──────────────────────
# def paste_into_fancy_layout(layout_file, sbag_dfs: list[pd.DataFrame]) -> bytes:
#     """
#     For each cleaned row in sbag_dfs:
#       - If Shape matches a shape sheet → paste into that sheet (col A onwards, after last data row)
#       - If Quality == CVD  → paste into TO ORDER CVD sheet
#       - If Quality == HPHT → paste into TO ORDER HPHT sheet
#     Returns bytes of the modified workbook.
#     """
#     import openpyxl
#     from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

#     wb = openpyxl.load_workbook(layout_file)

#     # combine all cleaned data
#     all_data = pd.concat(sbag_dfs, ignore_index=True) if sbag_dfs else pd.DataFrame()
#     if all_data.empty:
#         buf = BytesIO()
#         wb.save(buf)
#         return buf.getvalue()

#     paste_style = {
#         "font":      Font(name="Arial", size=9, bold=False),
#         "alignment": Alignment(horizontal="center", vertical="center"),
#         "border":    Border(
#             left=Side(style="thin", color="BFBFBF"),
#             right=Side(style="thin", color="BFBFBF"),
#             top=Side(style="thin", color="BFBFBF"),
#             bottom=Side(style="thin", color="BFBFBF"),
#         ),
#         "fill": PatternFill("solid", start_color="E8F4FD", end_color="E8F4FD"),
#     }

#     # track what got pasted
#     paste_log = {}

#     def _next_row(ws):
#         """Find next empty row in col A, after row 3 (header area)."""
#         for r in range(ws.max_row, 3, -1):
#             if ws.cell(row=r, column=1).value not in (None, ""):
#                 return r + 1
#         return 4

#     def _append_rows(ws, rows_df):
#         start = _next_row(ws)
#         cols = ["Item", "Shape", "Quality", "Color", "Grade", "", "MemoOut", "Udf7"]
#         for offset, (_, row) in enumerate(rows_df.iterrows()):
#             for c_idx, col_name in enumerate(cols, start=1):
#                 cell = ws.cell(row=start + offset, column=c_idx)
#                 cell.value     = row.get(col_name, "")
#                 cell.font      = paste_style["font"]
#                 cell.alignment = paste_style["alignment"]
#                 cell.border    = paste_style["border"]
#                 cell.fill      = paste_style["fill"]
#         return len(rows_df)

#     for _, row in all_data.iterrows():
#         shape   = str(row.get("Shape", "")).strip().upper()
#         quality = str(row.get("Quality", "")).strip().upper()

#         # 1. Try shape sheet
#         sheet_name = SHAPE_TO_SHEET.get(shape)
#         if sheet_name and sheet_name in wb.sheetnames:
#             ws = wb[sheet_name]
#             cols = ["Item", "Shape", "Quality", "Color", "Grade", "", "MemoOut", "Udf7"]
#             next_r = _next_row(ws)
#             for c_idx, col_name in enumerate(cols, start=1):
#                 cell = ws.cell(row=next_r, column=c_idx)
#                 cell.value     = row.get(col_name, "")
#                 cell.font      = paste_style["font"]
#                 cell.alignment = paste_style["alignment"]
#                 cell.border    = paste_style["border"]
#                 cell.fill      = paste_style["fill"]
#             paste_log.setdefault(sheet_name, 0)
#             paste_log[sheet_name] += 1

#         # 2. Also route CVD / HPHT to TO ORDER sheets
#         to_order_sheet = QUALITY_TO_SHEET.get(quality)
#         if to_order_sheet and to_order_sheet in wb.sheetnames:
#             ws = wb[to_order_sheet]
#             cols = ["Item", "Shape", "Quality", "Color", "Grade", "", "MemoOut", "Udf7"]
#             next_r = _next_row(ws)
#             for c_idx, col_name in enumerate(cols, start=1):
#                 cell = ws.cell(row=next_r, column=c_idx)
#                 cell.value     = row.get(col_name, "")
#                 cell.font      = paste_style["font"]
#                 cell.alignment = paste_style["alignment"]
#                 cell.border    = paste_style["border"]
#                 cell.fill      = paste_style["fill"]
#             paste_log.setdefault(to_order_sheet, 0)
#             paste_log[to_order_sheet] += 1

#     buf = BytesIO()
#     wb.save(buf)
#     return buf.getvalue(), paste_log


# # ── Melee Stock: highlight Total rows ─────────────────────────────────────────
# def highlight_melee_totals(melee_file) -> bytes:
#     import openpyxl
#     from openpyxl.styles import Font, PatternFill, Alignment

#     wb = openpyxl.load_workbook(melee_file)
#     highlight_fill   = PatternFill("solid", start_color="FFD700", end_color="FFD700")  # gold
#     highlight_font   = Font(name="Arial", size=10, bold=True, color="333333")
#     normal_fill      = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")

#     total_rows_found = 0
#     for sh in wb.sheetnames:
#         ws = wb[sh]
#         for row in ws.iter_rows():
#             first_val = str(row[0].value or "").strip().lower()
#             if first_val.startswith("total"):
#                 total_rows_found += 1
#                 for cell in row:
#                     cell.fill = highlight_fill
#                     if cell.value is not None:
#                         cell.font = highlight_font
            
#     buf = BytesIO()
#     wb.save(buf)
#     return buf.getvalue(), total_rows_found


# # ── UI ────────────────────────────────────────────────────────────────────────

# st.title("💎 Diamond Report Tool")

# tab1, tab2, tab3 = st.tabs([
#     "📋 S.BAG Cleaner (SBAG11)",
#     "📊 Fancy Layout Report Filler",
#     "🔦 Melee Stock Highlighter"
# ])


# # ════════════════════════════════════════════════════════════════════════════
# # TAB 1 — S.BAG Cleaner
# # ════════════════════════════════════════════════════════════════════════════
# with tab1:
#     st.subheader("Upload S.BAG (2) Files → Clean to SBAG11 Format")
#     st.markdown("Upload **one or more** S.BAG (2) `.xlsx` files.")

#     uploaded_sbag = st.file_uploader(
#         "📂 Upload S.BAG (2) files",
#         type=["xlsx"],
#         accept_multiple_files=True,
#         key="sbag_uploader",
#     )

#     if uploaded_sbag:
#         results, errors = [], []
#         prog = st.progress(0)
#         for i, f in enumerate(uploaded_sbag):
#             try:
#                 raw     = pd.read_excel(f, sheet_name="Sheet1", header=None)
#                 cleaned = clean_sbag(raw)
#                 cleaned.insert(0, "_src", f.name)
#                 out_name = f.name.replace(".xlsx","").replace(" ","_") + "_SBAG11.xlsx"
#                 results.append({
#                     "orig": f.name, "out_name": out_name,
#                     "cleaned": cleaned, "excel_bytes": to_excel_bytes_simple(cleaned.drop(columns=["_src"])),
#                     "rows": len(cleaned),
#                 })
#             except Exception as e:
#                 errors.append({"name": f.name, "error": str(e)})
#             prog.progress((i+1)/len(uploaded_sbag))
#         prog.empty()

#         c1, c2, c3 = st.columns(3)
#         c1.metric("Uploaded", len(uploaded_sbag))
#         c2.metric("Cleaned", len(results))
#         c3.metric("Errors", len(errors))

#         if errors:
#             st.error("Some files failed:")
#             for e in errors: st.markdown(f"- **{e['name']}**: `{e['error']}`")

#         if results:
#             st.markdown("---")
#             for r in results:
#                 with st.expander(f"✅ {r['orig']}  →  **{r['rows']} rows**", expanded=len(results)==1):
#                     st.dataframe(r["cleaned"].drop(columns=["_src"]), use_container_width=True)
#                     st.download_button(
#                         f"⬇️ Download {r['out_name']}", data=r["excel_bytes"],
#                         file_name=r["out_name"],
#                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#                         key=f"dl1_{r['out_name']}",
#                     )

#             st.markdown("---")
#             dc1, dc2 = st.columns(2)
#             dc1.download_button(
#                 f"🗜️ Download All as ZIP ({len(results)} files)",
#                 data=make_zip(results), file_name="SBAG11_all.zip", mime="application/zip",
#                 use_container_width=True,
#             )

#             # Merged workbook
#             mbuf = BytesIO()
#             from openpyxl.styles import Font as OFont, PatternFill as OFill, Alignment as OAlign
#             import openpyxl as oxl
#             mwb = oxl.Workbook()
#             mwb.remove(mwb.active)
#             for r in results:
#                 sname = r["orig"][:31].replace("/","-").replace("\\","-").replace("*","").replace("?","").replace("[","").replace("]","").replace(":","")
#                 ws = mwb.create_sheet(sname)
#                 df = r["cleaned"].drop(columns=["_src"])
#                 ws.append(list(df.columns))
#                 for row in df.itertuples(index=False):
#                     ws.append(list(row))
#                 hfill = OFill("solid", start_color="4472C4", end_color="4472C4")
#                 for cell in ws[1]:
#                     cell.font = OFont(bold=True, color="FFFFFF")
#                     cell.fill = hfill
#                     cell.alignment = OAlign(horizontal="center")
#             mwb.save(mbuf)
#             dc2.download_button(
#                 "📊 Download Merged Excel",
#                 data=mbuf.getvalue(), file_name="SBAG11_merged.xlsx",
#                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#                 use_container_width=True,
#             )
#     else:
#         st.info("👆 Upload one or more S.BAG (2) Excel files.")
#         with st.expander("ℹ️ Column mapping"):
#             st.markdown("""
# | S.BAG (2) | → | SBAG11 |
# |---|---|---|
# | Item | → | Item |
# | Shape | → | Shape |
# | Quality | → | Quality |
# | Color | → | Color |
# | Grade | → | Grade |
# | Col 8 (Wt) | → | *(weight)* |
# | MemoOut | → | MemoOut |
# | Udf7 | → | Udf7 |
# """)


# # ════════════════════════════════════════════════════════════════════════════
# # TAB 2 — Fancy Layout Report Filler
# # ════════════════════════════════════════════════════════════════════════════
# with tab2:
#     st.subheader("Paste Cleaned Data into Fancy Layout Report Sheets")
#     st.markdown("""
# Upload the **Fancy Layout Report** template (`.xlsx`) and **one or more S.BAG (2)** files.
# - Rows will be pasted into the **matching shape sheet** (ASSCHER, EMERALD, etc.)
# - CVD rows → also pasted into **TO ORDER CVD** sheet
# - HPHT rows → also pasted into **TO ORDER HPHT** sheet
# - Pasted rows are highlighted in **light blue**
# """)

#     col_a, col_b = st.columns(2)
#     with col_a:
#         layout_file = st.file_uploader(
#             "📄 Upload Fancy Layout Report (.xlsx)",
#             type=["xlsx"], key="layout_uploader",
#         )
#     with col_b:
#         sbag_files_2 = st.file_uploader(
#             "📂 Upload S.BAG (2) file(s) (.xlsx)",
#             type=["xlsx"], accept_multiple_files=True, key="sbag2_uploader",
#         )

#     if layout_file and sbag_files_2:
#         if st.button("🚀 Process & Fill Layout Report", type="primary", use_container_width=True):
#             with st.spinner("Cleaning S.BAG files and pasting into layout…"):
#                 # clean all sbag files
#                 cleaned_dfs = []
#                 sbag_errors = []
#                 for f in sbag_files_2:
#                     try:
#                         raw     = pd.read_excel(f, sheet_name="Sheet1", header=None)
#                         cleaned = clean_sbag(raw)
#                         cleaned_dfs.append(cleaned)
#                     except Exception as e:
#                         sbag_errors.append(f"{f.name}: {e}")

#                 if sbag_errors:
#                     st.warning("Some S.BAG files had errors:\n" + "\n".join(sbag_errors))

#                 if cleaned_dfs:
#                     total_rows = sum(len(d) for d in cleaned_dfs)
#                     try:
#                         layout_bytes, paste_log = paste_into_fancy_layout(layout_file, cleaned_dfs)
#                         st.success(f"✅ Pasted **{total_rows} rows** across **{len(paste_log)} sheets**")

#                         # Show summary
#                         if paste_log:
#                             summary_df = pd.DataFrame(
#                                 [{"Sheet": sh, "Rows Added": cnt} for sh, cnt in sorted(paste_log.items())]
#                             )
#                             st.dataframe(summary_df, use_container_width=True, hide_index=True)

#                         st.download_button(
#                             "⬇️ Download Filled Layout Report",
#                             data=layout_bytes,
#                             file_name="Fancy_Layout_Report_FILLED.xlsx",
#                             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#                             use_container_width=True,
#                         )
#                     except Exception as e:
#                         st.error(f"Error filling layout: {e}")
#                         st.exception(e)
#                 else:
#                     st.warning("No cleaned data to paste — check your S.BAG files.")
#     elif not layout_file:
#         st.info("👆 Upload the Fancy Layout Report template to get started.")
#     elif not sbag_files_2:
#         st.info("👆 Upload at least one S.BAG (2) file.")

#     st.markdown("---")
#     with st.expander("ℹ️ Shape → Sheet routing table"):
#         routes = [{"Shape Value": k, "→ Sheet": v} for k, v in SHAPE_TO_SHEET.items()]
#         routes += [{"Shape Value": "CVD (Quality)", "→ Sheet": "TO ORDER CVD"},
#                    {"Shape Value": "HPHT (Quality)", "→ Sheet": "TO ORDER HPHT"}]
#         st.dataframe(pd.DataFrame(routes), use_container_width=True, hide_index=True)


# # ════════════════════════════════════════════════════════════════════════════
# # TAB 3 — Melee Stock Highlighter
# # ════════════════════════════════════════════════════════════════════════════
# with tab3:
#     st.subheader("Melee Stock — Highlight Total Rows")
#     st.markdown("""
# Upload the **Melee Stock** file (`.xls` or `.xlsx`).  
# All rows whose first cell starts with **"Total"** will be highlighted in **gold** 🟡.
# """)

#     melee_file = st.file_uploader(
#         "📂 Upload Melee Stock file (.xls / .xlsx)",
#         type=["xls", "xlsx"],
#         key="melee_uploader",
#     )

#     if melee_file:
#         if st.button("🔦 Highlight Total Rows", type="primary", use_container_width=True):
#             with st.spinner("Processing Melee Stock file…"):
#                 try:
#                     # convert .xls → .xlsx in memory if needed
#                     fname = melee_file.name.lower()
#                     if fname.endswith(".xls"):
#                         # use xlrd + openpyxl conversion path
#                         import xlrd, openpyxl

#                         # read with xlrd
#                         xls_wb = xlrd.open_workbook(file_contents=melee_file.read())
#                         oxl_wb = openpyxl.Workbook()
#                         oxl_wb.remove(oxl_wb.active)

#                         for sheet_name in xls_wb.sheet_names():
#                             xls_ws = xls_wb.sheet_by_name(sheet_name)
#                             oxl_ws = oxl_wb.create_sheet(sheet_name)
#                             for rx in range(xls_ws.nrows):
#                                 for cx in range(xls_ws.ncols):
#                                     cell = xls_ws.cell(rx, cx)
#                                     oxl_ws.cell(row=rx+1, column=cx+1, value=cell.value)

#                         tmp_buf = BytesIO()
#                         oxl_wb.save(tmp_buf)
#                         tmp_buf.seek(0)
#                         work_file = tmp_buf
#                     else:
#                         work_file = melee_file

#                     result_bytes, total_count = highlight_melee_totals(work_file)

#                     st.success(f"✅ Highlighted **{total_count} Total rows** in gold.")

#                     # Preview
#                     preview = pd.read_excel(BytesIO(result_bytes), header=None)
#                     total_mask = preview.iloc[:, 0].astype(str).str.lower().str.startswith("total")
#                     st.markdown(f"**Preview** — {total_mask.sum()} Total rows (shown in table):")

#                     def color_total(row):
#                         if str(row.iloc[0]).lower().startswith("total"):
#                             return ["background-color: #FFD700; font-weight: bold"] * len(row)
#                         return [""] * len(row)

#                     st.dataframe(
#                         preview.head(80).style.apply(color_total, axis=1),
#                         use_container_width=True,
#                         height=400,
#                     )

#                     st.download_button(
#                         "⬇️ Download Highlighted Melee Stock",
#                         data=result_bytes,
#                         file_name="Melee_Stock_Highlighted.xlsx",
#                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#                         use_container_width=True,
#                     )
#                 except Exception as e:
#                     st.error(f"Error: {e}")
#                     st.exception(e)
#     else:
#         st.info("👆 Upload the Melee Stock .xls or .xlsx file.")
#         with st.expander("ℹ️ What gets highlighted?"):
#             st.markdown("""
# Rows where the **Item column** (col A) starts with `Total` — for example:
# - `Total-08-13`
# - `Total-CVD`
# - `Total-ASSCHER`
# - `Total-HPHT`
# - etc.

# These are highlighted in **gold** (`#FFD700`) with **bold text**.
# """)