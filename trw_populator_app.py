# -*- coding: utf-8 -*-
"""
UGL AR-Sheet / TRW Before & After Populator (Streamlit App)
===========================================================
Ingests a client UGL AR-Sheet (TRWxxxx) template and a TrackMeas app export
workbook, populates header metadata, matches/inserts measurement rows by KM
(rounded to 3 decimal places), formats Ballast ticks in Wingdings 2 ('P')
whenever After measurements exist, appends a 'Comments' column, and exports
the styled workbook with a user-defined filename prefix.
"""

import io
import datetime
from copy import copy
from pathlib import Path

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
    * **Before & After Measurements:** Matched by Location KM (`Kilometerage / 1000`, rounded to 3 decimals) with any extra measured locations inserted in chronological KM order
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
    # Handle ISO timestamp strings like '2026-11-22 00:00:00'
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
      - records: dict keyed by rounded 3-decimal KM float -> dict of mapped fields
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
                # Look right across next 5 cells for value
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
        header_row = 22  # Default fallback
        logs.append("Warning: 'KILOMETRAGE' header not found dynamically; defaulting to Row 22.")

    # Map column indices dynamically from group header (header_row) + subheader (header_row + 1)
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

    # Data starts after subheader rows (header_row + 3 in template where Row 22=group, 23=sub, 24=units, 25=data)
    data_start = header_row + 3
    records = {}

    for r in range(data_start, ws.max_row + 1):
        raw_km = ws.cell(r, col_map["km"]).value
        if raw_km is None or str(raw_km).strip() == "":
            continue
        try:
            km_float = round(float(raw_km) / 1000.0, 3)
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

        # Check if any measurement or comment exists on this row
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
            records[km_float] = rec

    logs.append(f"Extracted {len(records)} active measurement rows from TrackMeas export.")
    return {"date": date_val, "team_leader": team_leader}, records, logs


def populate_trw_workbook(trw_bytes: bytes, meta: dict, records: dict):
    """
    Populate the UGL TRW workbook in-place, preserving formatting and merged cells,
    inserting extra KM rows in sorted order, adding Wingdings 2 'P' ticks where After
    data exists, and appending the 'Comments' column in Column AC.
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

    # 3. Map existing TRW data rows (starting at Row 15)
    data_start_row = 15
    existing_km_rows = {}  # km_float -> row_idx
    for r in range(data_start_row, ws.max_row + 1):
        val = ws.cell(r, 2).value
        if val is not None and str(val).strip() != "":
            try:
                km_k = round(float(val), 3)
                existing_km_rows[km_k] = r
            except (ValueError, TypeError):
                continue

    # 4. Identify extra KMs from TrackMeas export that don't exist in TRW template
    extra_kms = sorted([km for km in records.keys() if km not in existing_km_rows])
    logs.append(f"Matched {len(records) - len(extra_kms)} existing KM rows; inserting {len(extra_kms)} new KM rows.")

    if extra_kms:
        # Snapshot existing template rows so we can rebuild cleanly without breaking merged cells (C:E, V:W, Y:Z)
        template_rows_data = []
        # Find a clean standard style row (e.g. Row 16 if available, else Row 15) for newly inserted rows
        ref_style_row = 16 if ws.max_row >= 16 else 15
        ref_styles = [copy(ws.cell(ref_style_row, c)._style) for c in range(1, comments_col + 1)]
        ref_height = ws.row_dimensions[ref_style_row].height

        for r in range(data_start_row, ws.max_row + 1):
            val = ws.cell(r, 2).value
            try:
                km_k = round(float(val), 3)
            except (ValueError, TypeError):
                continue
            row_vals = [ws.cell(r, c).value for c in range(1, comments_col + 1)]
            row_styles = [copy(ws.cell(r, c)._style) for c in range(1, comments_col + 1)]
            row_ht = ws.row_dimensions[r].height
            template_rows_data.append({
                "km": km_k,
                "vals": row_vals,
                "styles": row_styles,
                "height": row_ht,
                "is_new": False,
            })

        for km_new in extra_kms:
            rec = records[km_new]
            new_vals = [None] * comments_col
            new_vals[1] = km_new  # Col B: Location KM
            if rec.get("peg"):
                new_vals[9] = rec["peg"]  # Col J: TCM Side U/D
            if rec.get("align_des") is not None:
                new_vals[11] = rec["align_des"]  # Col L: TCM Design
            if rec.get("tc_des") is not None:
                new_vals[15] = rec["tc_des"]  # Col P: Track Centres Design
            if rec.get("super_des") is not None:
                new_vals[18] = rec["super_des"]  # Col S: Superelevation Design

            template_rows_data.append({
                "km": km_new,
                "vals": new_vals,
                "styles": [copy(s) for s in ref_styles],
                "height": ref_height,
                "is_new": True,
            })

        # Sort all rows chronologically by KM
        template_rows_data.sort(key=lambda x: x["km"])

        # Unmerge any existing row-level merges from Row 15 downward before rewriting
        merges_to_remove = [
            rng for rng in list(ws.merged_cells.ranges) if rng.min_row >= data_start_row
        ]
        for rng in merges_to_remove:
            ws.unmerge_cells(str(rng))

        # Write sorted rows back and re-apply standard row merges (C:E, V:W, Y:Z)
        for idx, item in enumerate(template_rows_data):
            r = data_start_row + idx
            if item["height"] is not None:
                ws.row_dimensions[r].height = item["height"]
            for c_idx in range(1, comments_col + 1):
                cell = ws.cell(r, c_idx)
                cell.value = item["vals"][c_idx - 1]
                cell._style = copy(item["styles"][c_idx - 1])

            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)    # C:E
            ws.merge_cells(start_row=r, start_column=22, end_row=r, end_column=23)  # V:W
            ws.merge_cells(start_row=r, start_column=25, end_row=r, end_column=26)  # Y:Z

    # 5. Populate Measurement Values, Wingdings 2 Ballast Ticks, and Comments
    wingdings_font = Font(name="Wingdings 2", size=11, bold=False)
    center_align = Alignment(horizontal="center", vertical="center")

    populated_count = 0
    ticked_count = 0
    preview_rows = []

    for r in range(data_start_row, ws.max_row + 1):
        val = ws.cell(r, 2).value
        if val is None or str(val).strip() == "":
            continue
        try:
            km_k = round(float(val), 3)
        except (ValueError, TypeError):
            continue

        ws.cell(r, 2).number_format = "0.000"

        if km_k in records:
            rec = records[km_k]
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

            preview_rows.append({
                "Location KM": f"{km_k:.3f}",
                "TCM Before": rec["align_bef"],
                "TCM After": rec["align_aft"],
                "Centres Before": rec["tc_bef"],
                "Centres After": rec["tc_aft"],
                "Super Before": rec["super_bef"],
                "Super After": rec["super_aft"],
                "Ballast Ticks": "✔ (P)" if has_after else "",
                "Comments": rec["comments"] or "",
            })

    logs.append(f"Populated {populated_count} measurement rows ({ticked_count} rows ticked with Wingdings 2 'P').")

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
        with st.spinner("Matching kilometrages and populating TRW sheet..."):
            try:
                meta, records, app_logs = parse_app_export(app_file.getvalue())
                out_bytes, preview_df, trw_logs = populate_trw_workbook(
                    trw_file.getvalue(), meta, records
                )

                # Build output filename: prefix + original TRW filename
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

    # Show extracted header summary
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
