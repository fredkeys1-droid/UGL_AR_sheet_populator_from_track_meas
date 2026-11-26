# -*- coding: utf-8 -*-
"""
UGL AR-Sheet / TRW Before & After Populator (Streamlit App)
===========================================================
Ingests a client UGL AR-Sheet (TRWxxxx) template and a TrackMeas app export
workbook, populates header metadata, matches existing UGL rows within +/- 1m
(0.001 km) of the TrackMeas export meterage without inserting extra rows,
populates the TCM Error formula in Column O, formats Ballast ticks in
Wingdings 2 ('P') whenever After measurements exist, appends a 'Comments'
column, and exports the styled workbook with a user-defined filename prefix.
"""

import io
import datetime
from copy import copy

import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Font
import pandas as pd
import streamlit as st

# ============================================================
# Page Configuration & Header
# ============================================================
st.set_page_config(
    page_title="UGL TRW Before & After Populator",
    page_icon="🚆",
    layout="wide",
)

st.title("🚆 UGL AR-Sheet (TRW) Before & After Populator")
st.markdown(
    """
    Upload your client **UGL AR-Sheet (`TRWxxxx`)** design template and your **TrackMeas App Export** 
    to automatically populate:
    * **Header Metadata:** `Date of Works`, `Supervisor`, and `Certifying Officer`
    * **Before & After Measurements:** Matched strictly to existing UGL rows within **±1 m (`0.001 km`)** of the measured location (no extra rows inserted)
    * **TCM Error Formula (`Col O`):** `=IF(M{r}="","",IF(K{r}="H",L{r}-N{r},N{r}-L{r}))` applied across data rows
    * **Ballast Ticks (`Down`, `Crib`, `Up`):** Formatted in **Wingdings 2** (`P` = ✔) on every row containing an **After** measurement
    * **Comments Column:** Appended cleanly to the right end of the table (`Col AC`)
    """
)
st.divider()


# ============================================================
# Helper Functions
# ============================================================
def copy_cell_style(src_cell, dst_cell):
    """Copy font, fill, border, alignment, and number_format from src_cell to dst_cell."""
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.border = copy(src_cell.border)
        dst_cell.alignment = copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format


def clean_int(val):
    """Convert numeric value to rounded Python int, or return None if blank/invalid."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    s = str(val).strip()
    if s in ("", "None", "nan", "NaN", "?"):
        return None
    try:
        return int(round(float(s)))
    except (ValueError, TypeError):
        return None


def clean_str(val):
    """Return stripped string or None if empty."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    s = str(val).strip()
    if s in ("", "None", "nan", "NaN"):
        return None
    return s


def format_date_val(val):
    """Format date value neatly for the TRW header cell."""
    if val is None:
        return ""
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%d/%m/%Y")
    s = str(val).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            dt = datetime.datetime.strptime(s, fmt)
            return dt.strftime("%d/%m/%Y")
        except ValueError:
            continue
    return s


def parse_app_export(app_bytes: bytes):
    """
    Parse the TrackMeas app export workbook ('Deliverables' or first matching sheet).
    Returns:
      - meta: dict with 'date', 'team_leader'
      - records: list of dicts with exact 'km' float (Kilometerage / 1000) and mapped fields
      - logs: list of status strings
    """
    logs = []
    wb_app = openpyxl.load_workbook(io.BytesIO(app_bytes), data_only=True)
    sheet_name = "Deliverables" if "Deliverables" in wb_app.sheetnames else wb_app.sheetnames[0]
    ws = wb_app[sheet_name]
    logs.append(f"Reading TrackMeas export sheet: '{sheet_name}'")

    # 1. Extract Header Metadata by label scanning (with fallback to known coordinates M4, T8)
    date_val = None
    team_leader = None

    for r in range(1, min(25, ws.max_row + 1)):
        for c in range(1, min(30, ws.max_column + 1)):
            cell_txt = str(ws.cell(r, c).value or "").strip().upper()
            if "TEAM LEADER" in cell_txt and team_leader is None:
                for offset in range(1, 6):
                    candidate = clean_str(ws.cell(r, c + offset).value)
                    if candidate:
                        team_leader = candidate
                        break
            elif cell_txt.startswith("DATE") and date_val is None:
                for offset in range(1, 6):
                    candidate = ws.cell(r, c + offset).value
                    if candidate is not None and str(candidate).strip() != "":
                        date_val = format_date_val(candidate)
                        break

    if team_leader is None:
        team_leader = clean_str(ws["M4"].value) or ""
    team_leader = team_leader.strip()
    if date_val is None:
        date_val = format_date_val(ws["T8"].value)

    logs.append(f"Extracted Header -> Date: '{date_val}' | Team Leader: '{team_leader}'")

    # 2. Locate Table Header Row (look for 'KILOMETRAGE')
    header_row = None
    for r in range(1, min(35, ws.max_row + 1)):
        row_vals = [str(ws.cell(r, c).value or "").strip().upper() for c in range(1, ws.max_column + 1)]
        if "KILOMETRAGE" in row_vals:
            header_row = r
            break

    if header_row is None:
        header_row = 22
        logs.append("Warning: 'KILOMETRAGE' header not found dynamically; defaulting to Row 22.")

    col_map = {
        "km": 2,          # B
        "peg": 3,         # C (U / D)
        "align_des": 4,   # D
        "align_bef": 5,   # E
        "align_aft": 6,   # F
        "super_des": 8,   # H
        "super_bef": 9,   # I
        "super_aft": 10,  # J
        "tc_des": 12,     # L
        "tc_bef": 13,     # M
        "tc_aft": 14,     # N
        "comments": 24,   # X
    }

    current_group = ""
    for c in range(1, ws.max_column + 1):
        grp = str(ws.cell(header_row, c).value or "").strip().upper()
        sub = str(ws.cell(header_row + 1, c).value or "").strip().upper()
        if grp:
            current_group = grp

        if grp == "KILOMETRAGE":
            col_map["km"] = c
        elif grp == "COMMENTS":
            col_map["comments"] = c
        elif current_group == "ALIGNMENT":
            if sub == "DESIGN":
                col_map["align_des"] = c
            elif sub == "BEFORE":
                col_map["align_bef"] = c
            elif sub == "AFTER":
                col_map["align_aft"] = c
        elif current_group == "SUPER":
            if sub == "DESIGN":
                col_map["super_des"] = c
            elif sub == "BEFORE":
                col_map["super_bef"] = c
            elif sub == "AFTER":
                col_map["super_aft"] = c
        elif "TRACK CENTRES" in current_group or "CENTRES" in current_group:
            if sub == "DESIGN":
                col_map["tc_des"] = c
            elif sub == "BEFORE":
                col_map["tc_bef"] = c
            elif sub == "AFTER":
                col_map["tc_aft"] = c

    data_start = header_row + 3
    records = []

    for r in range(data_start, ws.max_row + 1):
        raw_km = ws.cell(r, col_map["km"]).value
        if raw_km is None or str(raw_km).strip() == "":
            continue
        try:
            # Keep full precision in km (1m = 0.001 km) so +/- 1m matching isn't skewed by premature rounding
            km_float = float(raw_km) / 1000.0
        except (ValueError, TypeError):
            continue

        rec = {
            "km": km_float,
            "peg": clean_str(ws.cell(r, col_map["peg"]).value),
            "align_des": clean_int(ws.cell(r, col_map["align_des"]).value),
            "align_bef": clean_int(ws.cell(r, col_map["align_bef"]).value),
            "align_aft": clean_int(ws.cell(r, col_map["align_aft"]).value),
            "tc_des": clean_int(ws.cell(r, col_map["tc_des"]).value),
            "tc_bef": clean_int(ws.cell(r, col_map["tc_bef"]).value),
            "tc_aft": clean_int(ws.cell(r, col_map["tc_aft"]).value),
            "super_des": clean_int(ws.cell(r, col_map["super_des"]).value),
            "super_bef": clean_int(ws.cell(r, col_map["super_bef"]).value),
            "super_aft": clean_int(ws.cell(r, col_map["super_aft"]).value),
            "comments": clean_str(ws.cell(r, col_map["comments"]).value),
        }

        has_any_data = any(
            rec[k] is not None
            for k in [
                "align_bef", "align_aft",
                "tc_bef", "tc_aft",
                "super_bef", "super_aft",
                "comments",
            ]
        )
        if has_any_data:
            records.append(rec)

    logs.append(f"Extracted {len(records)} active measurement rows from TrackMeas export.")
    return {"date": date_val, "team_leader": team_leader}, records, logs


def populate_trw_workbook(trw_bytes: bytes, meta: dict, records: list, tolerance_m: float = 1.0):
    """
    Populate the UGL TRW workbook in-place without inserting extra rows.
    Matches existing TRW rows to TrackMeas export measurements within +/- tolerance_m (default 1.0m),
    picking the closest measurement (and assigning each measurement 1-to-1 to its nearest TRW row).
    """
    logs = []
    wb = openpyxl.load_workbook(io.BytesIO(trw_bytes))
    ws = wb.active
    logs.append(f"Loaded TRW template sheet: '{ws.title}'")

    # 1. Populate Header Block (F7 = Date of Works, F9 = Supervisor, F10 = Certifying Officer)
    if meta.get("date"):
        ws["F7"].value = meta["date"]
    if meta.get("team_leader"):
        ws["F9"].value = meta["team_leader"]
        ws["F10"].value = meta["team_leader"]

    # 2. Add 'Comments' Column at Column AC (Col 29) matching Row 12-14 header styles
    comments_col = 29  # AC
    copy_cell_style(ws.cell(12, 28), ws.cell(12, comments_col))
    copy_cell_style(ws.cell(13, 28), ws.cell(13, comments_col))
    copy_cell_style(ws.cell(14, 28), ws.cell(14, comments_col))

    ws.cell(12, comments_col).value = "Comments"
    ws.cell(13, comments_col).value = ""
    ws.cell(14, comments_col).value = ""
    ws.cell(12, comments_col).alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions["AC"].width = 32

    # Style existing rows (Row 15..max_row) in Col AC to match Col AB
    for r in range(15, ws.max_row + 1):
        copy_cell_style(ws.cell(r, 28), ws.cell(r, comments_col))
        ws.cell(r, comments_col).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # 3. Collect existing TRW data rows (starting at Row 15)
    data_start_row = 15
    trw_rows = []  # list of (row_idx, km_float)
    for r in range(data_start_row, ws.max_row + 1):
        val = ws.cell(r, 2).value
        if val is not None and str(val).strip() != "":
            try:
                km_k = float(val)
                trw_rows.append((r, km_k))
            except (ValueError, TypeError):
                continue

    # 4. Match TrackMeas records to existing TRW rows within +/- tolerance_m (1-to-1 closest match)
    tolerance_km = tolerance_m / 1000.0
    candidate_pairs = []
    for trw_idx, (r, trw_km) in enumerate(trw_rows):
        for rec_idx, rec in enumerate(records):
            diff_km = abs(trw_km - rec["km"])
            if diff_km <= tolerance_km + 1e-9:
                candidate_pairs.append((diff_km, trw_idx, rec_idx))

    # Sort by smallest distance first so the closest meterage pair always wins
    candidate_pairs.sort(key=lambda x: x[0])

    matched_by_row = {}  # row_idx -> (rec, diff_m)
    used_trw = set()
    used_rec = set()

    for diff_km, trw_idx, rec_idx in candidate_pairs:
        if trw_idx not in used_trw and rec_idx not in used_rec:
            used_trw.add(trw_idx)
            used_rec.add(rec_idx)
            row_idx = trw_rows[trw_idx][0]
            matched_by_row[row_idx] = (records[rec_idx], round(diff_km * 1000.0, 2))

    unmatched_count = len(records) - len(used_rec)
    logs.append(
        f"Matched {len(matched_by_row)} existing TRW rows within ±{tolerance_m:g}m "
        f"(ignored {unmatched_count} TrackMeas rows outside ±{tolerance_m:g}m; 0 rows inserted)."
    )

    # 5. Populate Measurement Values, TCM Error Formula, Wingdings 2 Ballast Ticks, and Comments
    wingdings_font = Font(name="Wingdings 2", size=11, bold=False)
    center_align = Alignment(horizontal="center", vertical="center")

    populated_count = 0
    ticked_count = 0
    preview_rows = []

    for r, km_k in trw_rows:
        ws.cell(r, 2).number_format = "0.000"

        # Apply TCM Error formula in Column O (Col 15) for every data row:
        # =IF(M{r}="","",IF(K{r}="H",L{r}-N{r},N{r}-L{r}))
        err_cell = ws.cell(r, 15)
        err_cell.value = f'=IF(M{r}="","",IF(K{r}="H",L{r}-N{r},N{r}-L{r}))'
        err_cell.number_format = "0"

        if r in matched_by_row:
            rec, diff_m = matched_by_row[r]
            populated_count += 1

            # TCM Before (Col M = 13) & After (Col N = 14)
            if rec["align_bef"] is not None:
                ws.cell(r, 13).value = rec["align_bef"]
            if rec["align_aft"] is not None:
                ws.cell(r, 14).value = rec["align_aft"]

            # Track Centres Before (Col Q = 17) & After (Col R = 18)
            if rec["tc_bef"] is not None:
                ws.cell(r, 17).value = rec["tc_bef"]
            if rec["tc_aft"] is not None:
                ws.cell(r, 18).value = rec["tc_aft"]

            # Superelevation Before (Col T = 20) & After (Col U = 21)
            if rec["super_bef"] is not None:
                ws.cell(r, 20).value = rec["super_bef"]
            if rec["super_aft"] is not None:
                ws.cell(r, 21).value = rec["super_aft"]

            # Ballast Ticks: Down (Col X = 24), Crib (Col Y = 25), Up (Col AA = 27)
            # Only populate 'P' (Wingdings 2 tick) when an AFTER measurement exists on that row
            has_after = any(
                rec[k] is not None for k in ("align_aft", "tc_aft", "super_aft")
            )
            if has_after:
                ticked_count += 1
                for b_col in (24, 25, 27):
                    b_cell = ws.cell(r, b_col)
                    b_cell.value = "P"
                    b_cell.font = wingdings_font
                    b_cell.alignment = center_align

            # Comments (Col AC = 29)
            if rec["comments"]:
                ws.cell(r, comments_col).value = rec["comments"]

            # Preview calculated TCM Error value for display in UI
            tcm_des = clean_int(ws.cell(r, 12).value)
            tcm_side_hl = (clean_str(ws.cell(r, 11).value) or "").upper()
            tcm_err_preview = None
            if rec["align_bef"] is not None and tcm_des is not None and rec["align_aft"] is not None:
                tcm_err_preview = (tcm_des - rec["align_aft"]) if tcm_side_hl == "H" else (rec["align_aft"] - tcm_des)

            preview_rows.append({
                "UGL Row KM": f"{km_k:.3f}",
                "Meas KM": f"{rec['km']:.3f}",
                "Offset (m)": diff_m,
                "TCM Design": tcm_des,
                "TCM Before": rec["align_bef"],
                "TCM After": rec["align_aft"],
                "TCM Error": tcm_err_preview,
                "Centres Before": rec["tc_bef"],
                "Centres After": rec["tc_aft"],
                "Super Before": rec["super_bef"],
                "Super After": rec["super_aft"],
                "Ballast Ticks": "✔ (P)" if has_after else "",
                "Comments": rec["comments"] or "",
            })

    logs.append(
        f"Populated {populated_count} existing UGL rows ({ticked_count} rows ticked with Wingdings 2 'P') "
        f"and applied TCM Error formulas in Col O."
    )

    out_buffer = io.BytesIO()
    wb.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer.getvalue(), pd.DataFrame(preview_rows), logs


# ============================================================
# Streamlit User Interface
# ============================================================
col1, col2 = st.columns(2)

with col1:
    trw_file = st.file_uploader(
        "1. Upload Client UGL AR-Sheet Template (TRWxxxx.xlsx)",
        type=["xlsx"],
        key="trw_uploader",
    )

with col2:
    app_file = st.file_uploader(
        "2. Upload TrackMeas App Export (.xlsx)",
        type=["xlsx"],
        key="app_uploader",
    )

prefix_text = st.text_input(
    "3. Filename Prefix (prepended to the uploaded TRW filename)",
    placeholder="e.g., Completed -  or  FINAL_",
    help="Whatever text you type here is placed directly at the start of the uploaded TRW filename.",
)

if trw_file and app_file:
    if st.button("🚀 Populate UGL AR-Sheet", type="primary"):
        with st.spinner("Matching kilometrages within ±1 m and populating TRW sheet..."):
            try:
                meta, records, app_logs = parse_app_export(app_file.getvalue())
                out_bytes, preview_df, trw_logs = populate_trw_workbook(
                    trw_file.getvalue(), meta, records, tolerance_m=1.0
                )

                output_filename = f"{prefix_text}{trw_file.name}"

                st.session_state["trw_result"] = {
                    "bytes": out_bytes,
                    "filename": output_filename,
                    "meta": meta,
                    "preview_df": preview_df,
                    "logs": app_logs + trw_logs,
                }
            except Exception as ex:
                st.error(f"Error processing files: {ex}")

if "trw_result" in st.session_state:
    res = st.session_state["trw_result"]
    st.success(f"Ready for download: **{res['filename']}**")

    m1, m2, m3 = st.columns(3)
    m1.metric("Date of Works (F7)", res["meta"].get("date") or "—")
    m2.metric("Supervisor (F9)", res["meta"].get("team_leader") or "—")
    m3.metric("Certifying Officer (F10)", res["meta"].get("team_leader") or "—")

    st.download_button(
        label=f"⬇️ Download {res['filename']}",
        data=res["bytes"],
        file_name=res["filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    if not res["preview_df"].empty:
        with st.expander("Preview Populated Rows", expanded=True):
            st.dataframe(res["preview_df"], use_container_width=True)

    with st.expander("Processing Log"):
        st.code("\n".join(res["logs"]), language="text")
