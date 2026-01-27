import streamlit as st
import pandas as pd
import io
from datetime import datetime, date
import difflib
import json
import copy
from typing import Any, Dict, List, Optional, Tuple
from io_helpers import (
    _safe_load_workbook,
    build_normalized_entity_index,
    patch_xlsx_workbook_xml_from_template,
    suggest_target_entity,
)
from openpyxl.utils import get_column_letter


CONTRIB_SHARED_SOURCE_KEY = "contrib_shared_source_file"
CONTRIB_SHARED_TARGET_BYTES_KEY = "contrib_shared_target_file_bytes"
CONTRIB_SHARED_TARGET_NAME_KEY = "contrib_shared_target_file_name"


def _norm(s: Any) -> str:
    return str(s or "").strip()


def _format_date_long(d: date) -> str:
    # Matches the template’s display style like “January 1, 2026”
    return f"{d.strftime('%B')} {d.day}, {d.year}"


def _to_date(v: Any) -> Optional[date]:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        ts = pd.to_datetime(str(v), errors="coerce")
        if pd.isna(ts):
            return None
        return ts.date()
    except Exception:
        return None


def _find_total_ie_row(ws) -> Optional[int]:
    # Looks for a cell like “Total IE: 69” in column A
    for r in range(1, ws.max_row + 1):
        v = _norm(ws.cell(row=r, column=1).value)
        if v.lower().startswith("total ie"):
            return r
    return None


def _copy_cell(src, dst) -> None:
    dst.value = src.value
    if src.has_style:
        dst._style = copy.copy(src._style)
    dst.number_format = src.number_format
    dst.font = copy.copy(src.font)
    dst.border = copy.copy(src.border)
    dst.fill = copy.copy(src.fill)
    dst.alignment = copy.copy(src.alignment)
    dst.protection = copy.copy(src.protection)
    dst.comment = src.comment


def _append_contrib_period_blocks(
    ws,
    *,
    base_marker: int,
    last_marker: int,
    entity_rows: List[int],
    total_ie_row: int,
    periods: List[Tuple[date, str]],
    width: int = 3,
    copy_row_end: Optional[int] = None,
) -> None:
    """Append 3-col contribution blocks, copying the base block and setting date/name/investment-date."""
    if copy_row_end is None:
        copy_row_end = total_ie_row

    base_close_cell = ws.cell(row=2, column=base_marker + 1)
    base_close_val = base_close_cell.value
    base_close_is_date = isinstance(base_close_val, (datetime, date))
    base_close_number_format = base_close_cell.number_format

    base_inv_number_format = "m/d/yyyy"
    if entity_rows:
        base_inv_number_format = ws.cell(row=entity_rows[0], column=base_marker + 2).number_format or base_inv_number_format

    for idx, (d, nm) in enumerate(periods):
        new_marker = last_marker + width * (idx + 1)

        # Copy columns (3-wide block) for all relevant rows
        for col_off in range(width):
            src_col = base_marker + col_off
            dst_col = new_marker + col_off

            # Copy column width if present
            src_letter = get_column_letter(src_col)
            dst_letter = get_column_letter(dst_col)
            if ws.column_dimensions[src_letter].width is not None:
                ws.column_dimensions[dst_letter].width = ws.column_dimensions[src_letter].width

            for r in range(1, copy_row_end + 1):
                _copy_cell(ws.cell(row=r, column=src_col), ws.cell(row=r, column=dst_col))

        # Write period-specific values
        # Contribution Name (row 2, marker col)
        ws.cell(row=2, column=new_marker).value = nm

        # Final Closing Date (row 2, marker+1)
        close_cell = ws.cell(row=2, column=new_marker + 1)
        if base_close_is_date:
            close_cell.value = datetime(d.year, d.month, d.day)
            # Preserve whatever the template uses (often long month format).
            close_cell.number_format = base_close_number_format
        else:
            # Template stores the date as text; match that convention.
            close_cell.value = _format_date_long(d)

        # Investment dates (all entity rows) in marker+2
        inv_col = new_marker + 2
        for r in entity_rows:
            c = ws.cell(row=r, column=inv_col)
            c.value = datetime(d.year, d.month, d.day)
            c.number_format = base_inv_number_format

        # Total SUM formula (aligned with Total IE row) in marker col
        if entity_rows:
            sum_cell = ws.cell(row=total_ie_row, column=new_marker)
            let = get_column_letter(new_marker)
            sum_cell.value = f"=SUM({let}{entity_rows[0]}:{let}{entity_rows[-1]})"

        # The template sometimes has placeholder "=" strings in footer cells to the right
        # of the amount-raised total. Clear them for generated blocks.
        ws.cell(row=total_ie_row, column=new_marker + 1).value = None
        ws.cell(row=total_ie_row, column=new_marker + 2).value = None


def _find_amount_raised_marker_cols(ws, *, min_col: int = 6) -> List[int]:
    """
    Find all columns that contain the 'Amount raised' marker (case-insensitive).
    We intentionally scan the header area (top of sheet) because the marker sits
    above the entity rows, and its exact row varies across exports.
    """
    cols: set[int] = set()
    scan_max_row = min(ws.max_row, 25)
    for r in range(1, scan_max_row + 1):
        for c in range(min_col, ws.max_column + 1):
            if _norm(ws.cell(row=r, column=c).value).lower() == "amount raised":
                cols.add(c)
    return sorted(cols)


def run_contributions_flow():
    """Contributions flow with Incomplete + Complete tabs."""
    tab_inc, tab_comp = st.tabs(["📝 Incomplete Import File", "📥 Complete Import File"])
    with tab_inc:
        run_contrib_incomplete_flow()
    with tab_comp:
        run_contrib_complete_flow()


def run_contrib_complete_flow():
    """Complete Import File flow for Contributions."""

    # --- Section 1: Upload files ------------------------------------------------
    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
    st.markdown('<h2>📥 Contributions · Complete Import File</h2>', unsafe_allow_html=True)

    saved_src_ok = st.session_state.get(CONTRIB_SHARED_SOURCE_KEY) is not None
    saved_tgt_ok = st.session_state.get(CONTRIB_SHARED_TARGET_BYTES_KEY) is not None
    default_idx = 0 if (saved_src_ok or saved_tgt_ok) else 1
    file_mode = st.radio(
        "File selection",
        ("↪️  Use files from Incomplete flow", "📂 Upload new files"),
        index=default_idx,
        horizontal=True,
        key="contrib_comp_file_mode",
    )
    use_saved = file_mode.startswith("↪️")

    src_col, tgt_col = st.columns(2)
    with src_col:
        shared = st.session_state.get(CONTRIB_SHARED_SOURCE_KEY)
        if use_saved and shared is not None:
            st.success(f"Using Source File from Incomplete flow: {shared.name}")
            source_file = shared
        else:
            st.markdown("##### Source Excel File")
            source_file = st.file_uploader("", type=["xlsx", "xls"], key="contrib_src", label_visibility="collapsed")
            if source_file is not None:
                st.session_state[CONTRIB_SHARED_SOURCE_KEY] = source_file
                use_saved = False

    with tgt_col:
        shared_bytes = st.session_state.get(CONTRIB_SHARED_TARGET_BYTES_KEY)
        shared_name = st.session_state.get(CONTRIB_SHARED_TARGET_NAME_KEY, "contrib_populated_template.xlsx")
        if use_saved and shared_bytes is not None:
            st.success(f"Using Populated Template from Incomplete flow: {shared_name}")
            target_file = io.BytesIO(shared_bytes)
            target_file.name = shared_name  # type: ignore[attr-defined]
            target_file_bytes_for_patch = shared_bytes
        else:
            st.markdown("##### Target Contribution Template")
            target_file = st.file_uploader("", type=["xlsx", "xls"], key="contrib_tgt", label_visibility="collapsed")
            if target_file is not None:
                st.session_state[CONTRIB_SHARED_TARGET_BYTES_KEY] = target_file.read()
                st.session_state[CONTRIB_SHARED_TARGET_NAME_KEY] = target_file.name
                target_file.seek(0)
                use_saved = False
            target_file_bytes_for_patch = st.session_state.get(CONTRIB_SHARED_TARGET_BYTES_KEY)

    if not (source_file and target_file):
        st.info("👆 Please upload both files to continue")
        return

    # --- Read source file -------------------------------------------------------
    df_source = pd.read_excel(source_file)

    # --- Read & parse target workbook ------------------------------------------
    target_file.seek(0)
    wb = _safe_load_workbook(target_file)
    if wb is None:
        return
    ws = wb[wb.sheetnames[0]]

    # --- Locate entity block ---------------------------------------------------
    colC_raw = [c.value for c in ws['C']]
    colC_norm = [str(v).strip() if v is not None else "" for v in colC_raw]

    try:
        ent_label_row = colC_norm.index("Investing Entity")  # zero-based
    except ValueError:
        st.error("❌ Could not locate 'Investing Entity' header in column C")
        return

    # Find footer row: prefer "Total IE" marker in column A, otherwise first blank after header in column C.
    total_ie_row_1b = _find_total_ie_row(ws)
    gp_row = None  # zero-based index into colC_norm
    if total_ie_row_1b is not None:
        total_ie_row_zb = total_ie_row_1b - 1
        if total_ie_row_zb > ent_label_row:
            gp_row = total_ie_row_zb

    if gp_row is None:
        try:
            gp_row = next(i for i in range(ent_label_row + 1, len(colC_norm)) if colC_norm[i] == "")
        except StopIteration:
            st.error("❌ Could not locate footer ('Total IE' in column A, or a blank row) after 'Investing Entity' header")
            return

    # Convert to 1-based Excel row numbers once and keep them unchanged afterwards
    entity_rows = [i + 1 for i in range(ent_label_row + 1, gp_row)]
    target_entities = [str(ws.cell(row=r, column=3).value).strip() for r in entity_rows]

    # Detect contribution period blocks (3-column width) -------------------------
    period_label_row = ent_label_row + 1  # convert to 1-based Excel row index
    block_starts = [j for j in range(6, ws.max_column + 1)
                    if str(ws.cell(row=period_label_row, column=j).value).strip() == "Amount raised"]
    if not block_starts:
        st.error("❌ No contribution blocks (Amount raised) found in the template")
        return

    blocks: List[Dict[str, Any]] = []
    for col in sorted(block_starts):
        contrib_name = _norm(ws.cell(row=2, column=col).value)
        # NOTE: period date is in the column immediately to the RIGHT of the marker.
        close_dt = _to_date(ws.cell(row=2, column=col + 1).value)
        if close_dt is not None:
            blocks.append({"date": close_dt, "name": contrib_name, "col": col})

    if not blocks:
        st.error("❌ Could not parse any contribution period dates from the template")
        return

    blocks_by_date: Dict[date, List[Dict[str, Any]]] = {}
    for b in blocks:
        blocks_by_date.setdefault(b["date"], []).append(b)
    has_duplicate_dates = any(len(v) > 1 for v in blocks_by_date.values())

    with st.expander("Detected contribution periods (debug)", expanded=False):
        st.dataframe(pd.DataFrame(blocks), use_container_width=True, hide_index=True)

    # --- Section 2: Preview & Column selection ---------------------------------
    st.markdown("##### Source Data Preview")
    st.dataframe(df_source.head(5), use_container_width=True)

    cols = df_source.columns.tolist()
    e_col, d_col, a_col = st.columns(3)
    with e_col:
        src_ent = st.selectbox("Investing Entity column", cols)
    with d_col:
        src_dt = st.selectbox("Contribution Date column", cols)
    with a_col:
        src_amt = st.selectbox("Amount column", cols)

    # Parse dates
    df_source['parsed_date'] = pd.to_datetime(df_source[src_dt], errors='coerce').dt.date

    # --- Section 3: Entity mapping --------------------------------------------
    unique_src = df_source[src_ent].dropna().astype(str).unique().tolist()
    map_df = pd.DataFrame({"source_entity": unique_src})
    target_index = build_normalized_entity_index(target_entities)
    sugg_pairs = map_df["source_entity"].apply(
        lambda x: suggest_target_entity(
            x, target_entities=target_entities, target_index=target_index, cutoff=0.6
        )
    )
    map_df[["target_entity", "match_type"]] = pd.DataFrame(sugg_pairs.tolist(), index=map_df.index)
    map_df = map_df[["source_entity", "match_type", "target_entity"]]

    st.markdown("##### Map source entities to template entities")
    edited = st.data_editor(
        map_df,
        column_config={
            'source_entity': st.column_config.TextColumn(
                "Source Entity",
                disabled=True,
            ),
            "match_type": st.column_config.TextColumn(
                "Match Type",
                help="How the Target Entity default was chosen (Exact/Fuzzy/None/Ambiguous)",
                disabled=True,
            ),
            'target_entity': st.column_config.SelectboxColumn(
                "Target Entity",
                help="Select the exact entity name from the template",
                options=[""] + target_entities,
                required=False,
            ),
        },
        hide_index=True,
        use_container_width=True,
    )

    entity_mapping = dict(zip(edited['source_entity'], edited['target_entity']))
    df_source['mapped_entity'] = df_source[src_ent].map(entity_mapping)

    # --- Contribution name matching (only needed if template has duplicate dates) ---
    valid_dates = {b["date"] for b in blocks}
    if has_duplicate_dates:
        st.markdown("##### Multiple periods share the same date — how should we match them?")
        name_match_mode = st.radio(
            "Contribution name matching",
            (
                "Match names from source sheet (requires a source Contribution Name column)",
                "Investor-separated blocks (use mapped entity as contribution name)",
            ),
            horizontal=True,
            key="contrib_name_match_mode",
        )

        if name_match_mode.startswith("Investor-separated"):
            df_source["mapped_contrib_name"] = df_source["mapped_entity"].fillna("").astype(str).str.strip()
        else:
            # Source-driven name matching (Mode D)
            extra_cols = [c for c in cols if c not in [src_ent, src_dt, src_amt]]
            if not extra_cols:
                st.error("❌ No extra columns found for Contribution Name matching. Add a name column to the source sheet or use investor-separated blocks.")
                return

            src_name_col = st.selectbox("Contribution Name column (source)", extra_cols, key="contrib_src_name_col")
            df_source["_src_contrib_name"] = df_source[src_name_col].fillna("").astype(str).str.strip()

            template_names = sorted({b["name"] for b in blocks if _norm(b.get("name")) != ""})
            unique_src_names = sorted({n for n in df_source["_src_contrib_name"].unique().tolist() if _norm(n) != ""})

            if not template_names:
                st.error("❌ Template contribution names are blank; cannot match by name. Use investor-separated blocks or regenerate the template with names.")
                return

            name_map_df = pd.DataFrame({"source_name": unique_src_names})
            name_map_df["suggestion"] = name_map_df["source_name"].apply(
                lambda x: (difflib.get_close_matches(x, template_names, n=1, cutoff=0.6) or [""])[0]
            )
            name_map_df["template_name"] = name_map_df["suggestion"]

            st.markdown("##### Map source contribution names to template contribution names")
            edited_names = st.data_editor(
                name_map_df,
                column_config={
                    "source_name": st.column_config.TextColumn("Source Name", disabled=True),
                    "suggestion": st.column_config.TextColumn("Suggested Match", disabled=True),
                    "template_name": st.column_config.SelectboxColumn(
                        "Template Name",
                        options=[""] + template_names,
                        required=False,
                    ),
                },
                hide_index=True,
                use_container_width=True,
                key="contrib_name_mapper",
            )
            mapping = dict(zip(edited_names["source_name"], edited_names["template_name"]))
            df_source["mapped_contrib_name"] = df_source["_src_contrib_name"].map(mapping).fillna("")

    else:
        df_source["mapped_contrib_name"] = ""

    # --- Resolve duplicates ----------------------------------------------------
    dup_src = df_source[(df_source['mapped_entity'] != "") & (df_source['parsed_date'].isin(valid_dates))]
    if has_duplicate_dates:
        dup_src = dup_src[dup_src["mapped_contrib_name"] != ""]

    group_cols = ["mapped_entity", "parsed_date"] + (["mapped_contrib_name"] if has_duplicate_dates else [])
    dup_groups = dup_src.groupby(group_cols)[src_amt].apply(list).reset_index(name="amounts")
    dups = dup_groups[dup_groups["amounts"].apply(len) > 1]

    chosen = {}

    if not dups.empty:
        st.markdown("##### Resolve Duplicates")
        st.markdown(
            "<small style='color: #6B7280;'>Multiple amounts found for the same entity/date. Choose which amount to keep (or SUM them). If you do nothing the first amount will be used.</small>",
            unsafe_allow_html=True,
        )

        if st.button("🔢 Sum All Duplicates", use_container_width=True, key="contrib_sum_dups"):
            for _, row in dups.iterrows():
                name_part = f"_{row.get('mapped_contrib_name','')}" if has_duplicate_dates else ""
                key = f"dup_{row['mapped_entity']}_{row['parsed_date']}{name_part}"
                st.session_state[key] = 'SUM'

        for _, row in dups.iterrows():
            ent, dt, amts = row['mapped_entity'], row['parsed_date'], row['amounts']
            nm = row.get("mapped_contrib_name", "") if has_duplicate_dates else ""
            key = f"dup_{ent}_{dt}_{nm}" if has_duplicate_dates else f"dup_{ent}_{dt}"
            options = [str(a) for a in amts] + ['SUM']
            if key not in st.session_state:
                st.session_state[key] = 'SUM'

            sel = st.radio(
                f"**{ent}** on {dt}" + (f" ({nm})" if has_duplicate_dates else ""),
                options,
                key=key,
                horizontal=True,
            )
            sel = st.session_state[key]
            chosen_key = (ent, dt, nm) if has_duplicate_dates else (ent, dt)
            chosen[chosen_key] = sum(amts) if sel == 'SUM' else float(sel)

    # --- Section 4: Fill amounts ----------------------------------------------
    unmatched = []
    for r, ent in zip(entity_rows, target_entities):
        for b in blocks:
            close_dt = b["date"]
            col_start = b["col"]
            if has_duplicate_dates:
                nm = b["name"]
                rows = df_source[
                    (df_source["mapped_entity"] == ent)
                    & (df_source["parsed_date"] == close_dt)
                    & (df_source["mapped_contrib_name"] == nm)
                ]
            else:
                rows = df_source[(df_source['mapped_entity'] == ent) &
                                  (df_source['parsed_date'] == close_dt)]
            if not rows.empty:
                chosen_key = (ent, close_dt, b["name"]) if has_duplicate_dates else (ent, close_dt)
                amt = chosen.get(chosen_key, rows[src_amt].iloc[0])
                ws.cell(row=r, column=col_start).value = amt
            else:
                if has_duplicate_dates:
                    unmatched.append((ent, close_dt, b["name"]))
                else:
                    unmatched.append((ent, close_dt))

    # --- Finalize & download ---------------------------------------------------
    buf = io.BytesIO()
    wb.save(buf)
    generated_bytes = buf.getvalue()

    # Preserve Covercy workbook-level metadata/extensions by restoring xl/workbook.xml
    # from the uploaded template (Covercy may validate this).
    if target_file_bytes_for_patch:
        try:
            generated_bytes = patch_xlsx_workbook_xml_from_template(target_file_bytes_for_patch, generated_bytes)
        except Exception:
            # Best-effort: if patching fails, still allow download for debugging.
            pass

    buf = io.BytesIO(generated_bytes)
    buf.seek(0)

    st.success("✅ Contribution import file generated!")
    st.download_button("📥 Download Import File", data=buf,
                       file_name=f"contributions_import_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)

    if unmatched:
        with st.expander(f"⚠️ {len(unmatched)} unmatched entries"):
            cols_out = ["Entity", "Date", "Contribution Name"] if has_duplicate_dates else ["Entity", "Date"]
            st.dataframe(pd.DataFrame(unmatched, columns=cols_out))


def run_contrib_incomplete_flow():
    """Generate a multi-period Contributions template from a single-period template."""
    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
    st.markdown('<h2>📝 Contributions · Incomplete Import File</h2>', unsafe_allow_html=True)

    src_col, tgt_col = st.columns(2)
    with src_col:
        st.markdown("##### Source Excel File")
        uploaded_src = st.file_uploader("", type=["xlsx", "xls"], key="contrib_inc_src", label_visibility="collapsed")
        if uploaded_src is not None:
            source_file = uploaded_src
            st.session_state[CONTRIB_SHARED_SOURCE_KEY] = uploaded_src
        else:
            shared = st.session_state.get(CONTRIB_SHARED_SOURCE_KEY)
            source_file = shared
            if shared is not None:
                st.info(f"Using previously uploaded file: {shared.name}")

    with tgt_col:
        st.markdown("##### Incomplete Target File (single period)")
        target_file = st.file_uploader("", type=["xlsx", "xls"], key="contrib_inc_tgt", label_visibility="collapsed")
        if target_file is not None:
            # Validate immediately to surface “Save As…” guidance early
            wb_test = _safe_load_workbook(io.BytesIO(target_file.getvalue()))
            if wb_test is None:
                st.stop()

    if not (source_file and target_file):
        st.info("👆 Please upload both files to continue")
        return

    df_src = pd.read_excel(source_file)
    st.markdown("##### Source Data Preview")
    st.dataframe(df_src.head(5), use_container_width=True)

    cols = df_src.columns.tolist()
    default_date_idx = 0
    for i, c in enumerate(cols):
        if "date" in str(c).lower():
            default_date_idx = i
            break

    date_col = st.selectbox(
        "Contribution Date column",
        options=cols,
        index=default_date_idx,
        key="contrib_inc_date_col",
    )

    df_src["parsed_date"] = pd.to_datetime(df_src[date_col], errors="coerce").dt.date
    invalid = df_src["parsed_date"].isna().sum()
    if invalid:
        st.warning(f"⚠️ {invalid} rows have unparseable dates and will be skipped.")

    # Load template workbook (keep raw bytes for workbook.xml patching later)
    template_bytes = target_file.read()
    wb = _safe_load_workbook(io.BytesIO(template_bytes))
    if wb is None:
        return
    ws = wb.active

    # Determine entity row span (to set Investment dates + build total formula)
    colC_raw = [c.value for c in ws['C']]
    colC_norm = [str(v).strip() if v is not None else "" for v in colC_raw]
    try:
        ent_label_row_zb = colC_norm.index("Investing Entity")
    except ValueError:
        st.error("❌ Could not locate 'Investing Entity' header in column C")
        return
    # Prefer "Total IE" marker in column A; fallback to first blank in column C.
    total_ie_row_1b = _find_total_ie_row(ws)
    gp_row_zb = None
    if total_ie_row_1b is not None:
        total_ie_row_zb = total_ie_row_1b - 1
        if total_ie_row_zb > ent_label_row_zb:
            gp_row_zb = total_ie_row_zb

    if gp_row_zb is None:
        try:
            gp_row_zb = next(i for i in range(ent_label_row_zb + 1, len(colC_norm)) if colC_norm[i] == "")
        except StopIteration:
            st.error("❌ Could not locate footer ('Total IE' in column A, or a blank row) after 'Investing Entity'")
            return
    ent_label_row = ent_label_row_zb + 1
    gp_row = gp_row_zb + 1
    entity_rows = list(range(ent_label_row + 1, gp_row))  # investor rows only

    total_ie_row = _find_total_ie_row(ws)
    if total_ie_row is None:
        # best-effort fallback: put totals right after entity block
        total_ie_row = gp_row

    # Find existing period marker column(s) on the SAME ROW as "Investing Entity" (no fallback).
    # In the Covercy contributions template, "Amount raised" appears on the same header row.
    period_label_row = ent_label_row
    marker_cols = [
        j
        for j in range(6, ws.max_column + 1)
        if _norm(ws.cell(row=period_label_row, column=j).value).lower() == "amount raised"
    ]
    if not marker_cols:
        st.error("❌ Could not find the 'Amount raised' header on the same row as 'Investing Entity' (column C).")
        return

    marker_cols = sorted(marker_cols)
    base_marker = marker_cols[0]
    last_marker = marker_cols[-1]
    width = 3  # confirmed by you

    existing_date = _to_date(ws.cell(row=2, column=base_marker + 1).value)

    uniq_dates = sorted([d for d in df_src["parsed_date"].dropna().unique().tolist() if isinstance(d, date)])
    if existing_date is not None:
        uniq_dates = [d for d in uniq_dates if d != existing_date]
    if not uniq_dates:
        st.error("❌ No valid contribution dates found in the source file (after excluding the template’s existing date).")
        return

    st.markdown("##### Select Contribution Dates")
    start_date = st.date_input("Start date", value=uniq_dates[0], min_value=uniq_dates[0], max_value=uniq_dates[-1], key="contrib_inc_start")
    end_date = st.date_input("End date", value=uniq_dates[-1], min_value=uniq_dates[0], max_value=uniq_dates[-1], key="contrib_inc_end")
    dates_in_range = [d for d in uniq_dates if start_date <= d <= end_date]

    date_strs = [d.isoformat() for d in dates_in_range]
    if "contrib_inc_sel_dates" not in st.session_state:
        st.session_state["contrib_inc_sel_dates"] = date_strs.copy()

    selected = st.multiselect(
        "Dates to add",
        options=date_strs,
        default=st.session_state.get("contrib_inc_sel_dates", date_strs),
        key="contrib_inc_sel_dates",
        format_func=lambda x: _format_date_long(datetime.strptime(x, "%Y-%m-%d").date()),
    )
    dates_to_use = sorted([datetime.strptime(s, "%Y-%m-%d").date() for s in selected])

    if not dates_to_use:
        st.warning("Select at least one date to generate periods.")
        return

    # Contribution Name options (A/B/C/D)
    st.markdown("##### Contribution Name strategy")
    name_mode = st.radio(
        "How should Contribution Name be set for new periods?",
        (
            "A — Same as date",
            "B — Write the names manually",
            "C — Separate by investor name",
            "D — Match names from source sheet",
        ),
        horizontal=True,
        key="contrib_inc_name_mode",
    )

    periods: List[Tuple[date, str]] = []

    if name_mode.startswith("A"):
        periods = [(d, _format_date_long(d)) for d in dates_to_use]

    elif name_mode.startswith("B"):
        table_key = "contrib_inc_manual_names"
        if table_key not in st.session_state:
            st.session_state[table_key] = pd.DataFrame(
                [{"date": d.isoformat(), "contribution_name": _format_date_long(d)} for d in dates_to_use]
            )
        df_names_all = st.session_state[table_key].copy()

        # Keep in sync with selected dates (preserve existing names where possible)
        selected_dates = [d.isoformat() for d in dates_to_use]
        if not df_names_all.empty and "date" in df_names_all.columns:
            df_names_all["date"] = df_names_all["date"].astype(str)
        else:
            df_names_all = pd.DataFrame(columns=["date", "contribution_name"])

        # Add newly selected dates
        existing = set(df_names_all["date"].tolist())
        for d in dates_to_use:
            ds = d.isoformat()
            if ds not in existing:
                df_names_all.loc[len(df_names_all)] = {"date": ds, "contribution_name": _format_date_long(d)}

        # Drop unselected dates
        df_names_all = df_names_all[df_names_all["date"].isin(selected_dates)].copy()
        df_names_all = df_names_all.sort_values("date").reset_index(drop=True)

        st.caption("Edits are applied when you click **Save names**.")

        with st.form("contrib_inc_names_form", border=False):
            edited_names = st.data_editor(
                df_names_all,
                column_config={
                    "date": st.column_config.TextColumn("Date", disabled=True),
                    "contribution_name": st.column_config.TextColumn("Contribution Name"),
                },
                hide_index=True,
                use_container_width=True,
                key="contrib_inc_names_editor",
            )
            save_names = st.form_submit_button("💾 Save names", use_container_width=True)

        if save_names:
            st.session_state[table_key] = edited_names
            st.session_state["contrib_inc_names_last_saved_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            try:
                st.toast("✅ Names saved", icon="✅")
            except Exception:
                pass

        last_saved = st.session_state.get("contrib_inc_names_last_saved_at")
        if last_saved:
            st.caption(f"Last saved: {last_saved}")

        # Build periods from the last saved state (or defaults if never saved)
        use_df = st.session_state.get(table_key, df_names_all)
        periods = [
            (datetime.strptime(r["date"], "%Y-%m-%d").date(), _norm(r["contribution_name"]))
            for _, r in use_df.iterrows()
            if _norm(r["contribution_name"]) != ""
        ]
        if len(periods) != len(dates_to_use):
            st.warning("Some dates are missing a Contribution Name; those periods will not be generated.")

    elif name_mode.startswith("C"):
        ent_col = st.selectbox(
            "Investor/Entity column (source)",
            options=[c for c in df_src.columns if c != date_col],
            key="contrib_inc_investor_col",
        )
        df_src["_inv_name"] = df_src[ent_col].fillna("").astype(str).str.strip()
        for d in dates_to_use:
            invs = sorted({n for n in df_src.loc[df_src["parsed_date"] == d, "_inv_name"].unique().tolist() if _norm(n) != ""})
            periods.extend([(d, inv) for inv in invs])

    else:  # D
        name_col = st.selectbox(
            "Contribution Name column (source)",
            options=[c for c in df_src.columns if c not in [date_col]],
            key="contrib_inc_source_name_col",
        )
        df_src["_src_round"] = df_src[name_col].fillna("").astype(str).str.strip()
        for d in dates_to_use:
            names = sorted({n for n in df_src.loc[df_src["parsed_date"] == d, "_src_round"].unique().tolist() if _norm(n) != ""})
            periods.extend([(d, nm) for nm in names])

    periods = [(d, n) for d, n in periods if n != ""]
    if not periods:
        st.error("❌ No periods to generate after applying the Contribution Name strategy.")
        return

    st.markdown("##### Summary")
    st.write(f"Will generate **{len(periods)}** new period blocks.")

    if st.button("🚀 Generate Populated Template", use_container_width=True, type="primary", key="contrib_inc_generate"):
        copy_row_end = max(total_ie_row, max(entity_rows) if entity_rows else total_ie_row)
        _append_contrib_period_blocks(
            ws,
            base_marker=base_marker,
            last_marker=last_marker,
            entity_rows=entity_rows,
            total_ie_row=total_ie_row,
            periods=periods,
            width=width,
            copy_row_end=copy_row_end,
        )

        buf = io.BytesIO()
        wb.save(buf)
        generated_bytes = buf.getvalue()

        # Preserve Covercy workbook-level metadata/extensions by restoring xl/workbook.xml
        # from the original Covercy template.
        try:
            generated_bytes = patch_xlsx_workbook_xml_from_template(template_bytes, generated_bytes)
        except Exception:
            pass

        buf = io.BytesIO(generated_bytes)
        buf.seek(0)

        file_name = f"contrib_populated_template_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.session_state[CONTRIB_SHARED_TARGET_BYTES_KEY] = generated_bytes
        st.session_state[CONTRIB_SHARED_TARGET_NAME_KEY] = file_name

        st.success("✅ Template populated successfully!")
        st.download_button(
            "📥 Download Populated Template",
            data=buf,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.info("Next: switch to the Complete Import File tab and choose “Use files from Incomplete flow”.")