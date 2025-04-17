import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
import difflib
import os

# Streamlit page configuration and branding
st.set_page_config(
    page_title="Covercy‑Style Excel Merge",
    page_icon="logo.png",  # Local logo file
    layout="wide",
)

# Custom CSS for a sleek, welcoming Covercy feel
st.markdown(
    """
    <style>
      body { background-color: #ffffff; }
      .stButton>button {
        background-color: #FBBF24;
        color: #111827;
        border-radius: 0.5rem;
        padding: 0.6rem 1.2rem;
      }
      h1, h2, h3 { font-family: 'Helvetica Neue', Arial, sans-serif; }
    </style>
    """,
    unsafe_allow_html=True,
)

# Branded header with local logo
# Compute absolute path to logo so it's found regardless of current working directory
this_dir = os.path.dirname(os.path.abspath(__file__))
logo_path = os.path.join(this_dir, "logo.png")
col1, col2 = st.columns([1, 5])
with col1:
    st.image(logo_path, width=160)
with col2:
    st.markdown(
        """
        <h1 style="margin:0; color:#111827; font-family:Arial, sans-serif;">
          Covercy Excel Merge Tool
        </h1>
        """,
        unsafe_allow_html=True,
    )
st.markdown("<hr/>", unsafe_allow_html=True)

st.title("Excel Merge with Entity Mapping and Duplicate Resolution")

# 1. Upload source & target files
source_file = st.file_uploader(
    "Upload source Excel file", type=["xlsx", "xls"], key="src"
)
target_file = st.file_uploader(
    "Upload target Excel file", type=["xlsx", "xls"], key="tgt"
)

if source_file and target_file:
    # 2. Load & preview source
    df_source = pd.read_excel(source_file)
    st.subheader("Source Data Preview")
    st.dataframe(df_source.head(5))

    # 3. Column mapping for source
    cols = df_source.columns.tolist()
    source_entity_col = st.selectbox(
        "Select Investing Entity column from source", cols
    )
    source_date_col = st.selectbox(
        "Select Date column from source", cols
    )
    source_amount_col = st.selectbox(
        "Select Amount column from source", cols
    )

    # Parse dates in source
    df_source['parsed_date'] = pd.to_datetime(
        df_source[source_date_col], errors='coerce'
    ).dt.date
    invalid_dates = df_source['parsed_date'].isna().sum()
    if invalid_dates:
        st.warning(
            f"{invalid_dates} rows have unparseable dates and will be skipped."
        )

    # 4. Inspect target for entity & date axes
    df_raw = pd.read_excel(target_file, header=None)
    ent_col = 2  # Excel column C
    ent_label_row = df_raw[df_raw[ent_col] == 'Investing Entity'].index[0]
    gp_row = df_raw[df_raw[ent_col] == 'GP'].index[0]
    ent_rows = list(range(ent_label_row + 1, gp_row))
    target_entities = [
        str(df_raw.iat[r, ent_col]).strip() for r in ent_rows
    ]

    date_label_row = ent_label_row - 2
    date_cols = [
        j for j in range(df_raw.shape[1])
        if str(df_raw.iat[date_label_row, j]).strip() == 'Last Day'
    ]
    date_map = {
        j: pd.to_datetime(
            df_raw.iat[date_label_row + 1, j], errors='coerce'
        ).date()
        for j in date_cols
    }

    # 5. Unique source-entity mapping UI
    unique_src = (
        df_source[source_entity_col]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )
    map_df = pd.DataFrame({'source_entity': unique_src})
    map_df['suggestion_1'] = map_df['source_entity'].apply(
        lambda x: (
            difflib.get_close_matches(x, target_entities, n=1, cutoff=0.6)
            or [""]
        )[0]
    )
    map_df['target_entity'] = map_df['suggestion_1']

    st.subheader("Map Source Entities to Target Entities")
    edited_map = st.data_editor(
        map_df,
        column_config={
            'target_entity': {
                'editable': True,
                'type': 'dropdown',
                'options': target_entities
            }
        },
        hide_index=True,
    )

    # 6. Prepare mapping and detect duplicates
    mapping = dict(
        zip(edited_map['source_entity'], edited_map['target_entity'])
    )
    df_source['mapped_entity'] = (
        df_source[source_entity_col].map(mapping)
    )
    valid_dates = set(date_map.values())
    dup_source = df_source[
        (df_source['mapped_entity'] != "") &
        (df_source['parsed_date'].isin(valid_dates))
    ]
    dup_groups = (
        dup_source.groupby(['mapped_entity', 'parsed_date'])[source_amount_col]
                  .apply(list)
                  .reset_index(name='amounts')
    )
    dups = dup_groups[dup_groups['amounts'].apply(len) > 1]
    chosen_amounts = {}
    if not dups.empty:
        st.subheader("Resolve Duplicate Amounts")
        for _, row in dups.iterrows():
            ent, dt, amts = (
                row['mapped_entity'], row['parsed_date'], row['amounts']
            )
            key = f"dup_{ent}_{dt}"
            options = [str(a) for a in amts] + ['SUM']
            if key not in st.session_state:
                st.session_state[key] = options[-1]
            choice = st.radio(
                f"Select amount for {ent} on {dt}",
                options,
                key=key
            )
            chosen_amounts[(ent, dt)] = (
                sum(amts) if choice == 'SUM' else float(choice)
            )

    # 7. Finalize and download
    if st.button("Finalize and Download Updated Target"):
        wb = load_workbook(filename=target_file)
        ws = wb[wb.sheetnames[0]]
        unmatched = []
        for r, target_ent in zip(ent_rows, target_entities):
            for col_idx, dist_date in date_map.items():
                if pd.isna(dist_date):
                    continue
                mask = (
                    (df_source['mapped_entity'] == target_ent) &
                    (df_source['parsed_date'] == dist_date)
                )
                match = df_source.loc[mask]
                if not match.empty:
                    amt = chosen_amounts.get(
                        (target_ent, dist_date),
                        match[source_amount_col].iloc[0]
                    )
                    ws.cell(row=r+1, column=col_idx).value = amt
                else:
                    unmatched.append((target_ent, dist_date))
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        st.success("Updated target workbook successfully.")
        st.download_button(
            "Download Updated Target",
            data=buf,
            file_name="merged.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        if unmatched:
            st.warning(f"{len(unmatched)} entries were not matched.")
            st.write(unmatched)
