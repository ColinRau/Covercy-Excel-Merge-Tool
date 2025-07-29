import streamlit as st
import pandas as pd
import io
from datetime import datetime
import difflib
import json
from io_helpers import _safe_load_workbook


def run_contributions_flow():
    """Complete Import File flow for Contributions (mirrors Distributions)."""

    # --- Section 1: Upload files ------------------------------------------------
    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
    st.markdown('<h2>📥 Contributions · Complete Import File</h2>', unsafe_allow_html=True)

    src_col, tgt_col = st.columns(2)
    with src_col:
        st.markdown("##### Source Excel File")
        source_file = st.file_uploader("", type=["xlsx", "xls"], key="contrib_src", label_visibility="collapsed")
    with tgt_col:
        st.markdown("##### Target Contribution Template")
        target_file = st.file_uploader("", type=["xlsx", "xls"], key="contrib_tgt", label_visibility="collapsed")

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

    # Find GP/footer row: prefer explicit 'GP', otherwise first blank after header
    try:
        gp_row = colC_norm.index("GP", ent_label_row + 1)
    except ValueError:
        try:
            gp_row = next(i for i in range(ent_label_row + 1, len(colC_norm)) if colC_norm[i] == "")
        except StopIteration:
            st.error("❌ Could not locate footer (blank row) after 'Investing Entity' header")
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

    contrib_map = {}
    for col in block_starts:
        name = str(ws.cell(row=1, column=col).value or "").strip()
        close_dt_raw = ws.cell(row=2, column=col).value
        # Normalise the closing-date value to a `datetime.date` so it matches
        # the `parsed_date` column we built from the source sheet.
        if isinstance(close_dt_raw, datetime):
            close_dt = close_dt_raw.date()
        else:
            try:
                close_dt = pd.to_datetime(str(close_dt_raw), errors="coerce").date()
            except Exception:
                close_dt = None
        if close_dt is not None:
            contrib_map[close_dt] = col

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

    # No contribution name matching – use date only
    df_source['mapped_name'] = ""

    # --- Section 3: Entity mapping --------------------------------------------
    unique_src = df_source[src_ent].dropna().astype(str).unique().tolist()
    map_df = pd.DataFrame({'source_entity': unique_src})
    map_df['suggestion'] = map_df['source_entity'].apply(lambda x: (difflib.get_close_matches(x, target_entities, n=1, cutoff=0.6) or [""])[0])
    map_df['target_entity'] = map_df['suggestion']

    st.markdown("##### Map source entities to template entities")
    edited = st.data_editor(
        map_df,
        column_config={
            'source_entity': st.column_config.TextColumn(
                "Source Entity",
                disabled=True,
            ),
            'suggestion': st.column_config.TextColumn(
                "Suggested Match",
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

    # --- Section 4: Resolve duplicates (entity + date) ------------------------
    dup_src = df_source[
        (df_source['mapped_entity'] != "") &
        df_source['parsed_date'].isin(contrib_map.keys())
    ]

    dup_groups = (
        dup_src.groupby(['mapped_entity', 'parsed_date'])[src_amt]
        .apply(list)
        .reset_index(name='amounts')
    )
    dups = dup_groups[dup_groups['amounts'].apply(len) > 1]

    chosen = {}

    if not dups.empty:
        st.markdown("##### Resolve Duplicates")
        st.markdown(
            "<small style='color: #6B7280;'>Multiple amounts found for the same entity/date. Choose which amount to keep (or SUM them). If you do nothing the first amount will be used.</small>",
            unsafe_allow_html=True,
        )

        if st.button("🔢 Sum All Duplicates", use_container_width=True, key="contrib_sum_dups"):
            for _, row in dups.iterrows():
                key = f"dup_{row['mapped_entity']}_{row['parsed_date']}"
                st.session_state[key] = 'SUM'

        for _, row in dups.iterrows():
            ent, dt, amts = row['mapped_entity'], row['parsed_date'], row['amounts']
            key = f"dup_{ent}_{dt}"
            options = [str(a) for a in amts] + ['SUM']
            if key not in st.session_state:
                st.session_state[key] = 'SUM'

            sel = st.radio(
                f"**{ent}** on {dt}",
                options,
                key=key,
                horizontal=True,
            )
            sel = st.session_state[key]
            chosen[(ent, dt)] = sum(amts) if sel == 'SUM' else float(sel)

    # --- Section 4: Fill amounts ----------------------------------------------
    unmatched = []
    for r, ent in zip(entity_rows, target_entities):
        for close_dt, col_start in contrib_map.items():
            rows = df_source[(df_source['mapped_entity'] == ent) &
                              (df_source['parsed_date'] == close_dt)]
            if not rows.empty:
                amt = chosen.get((ent, close_dt), rows[src_amt].iloc[0])
                ws.cell(row=r, column=col_start).value = amt
            else:
                unmatched.append((ent, close_dt))

    # --- Finalize & download ---------------------------------------------------
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    st.success("✅ Contribution import file generated!")
    st.download_button("📥 Download Import File", data=buf,
                       file_name=f"contributions_import_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)

    if unmatched:
        with st.expander(f"⚠️ {len(unmatched)} unmatched entries"):
            st.dataframe(pd.DataFrame(unmatched, columns=['Entity', 'Date'])) 