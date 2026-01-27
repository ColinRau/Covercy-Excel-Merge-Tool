import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import difflib
import os
from datetime import datetime, date, timedelta
import json
import base64
from math import ceil
import zipfile
from openpyxl.utils.exceptions import InvalidFileException


# Page config & branding
try:
    st.set_page_config(
        page_title="Covercy Excel Merge Tool 2.0",
        page_icon="logo.png",
        layout="wide",
        initial_sidebar_state="collapsed"
    )
except Exception:
    # If the page config is already set by the navigation layer, ignore.
    pass

# Custom CSS for Covercy branding
st.markdown(
    """
    <style>
    /* Main styles */
    .main { 
        padding: 0;
        max-width: 1400px;
        margin: 0 auto;
    }
    
    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Typography */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    /* Headers */
    h1 {
        color: #0A2540;
        font-weight: 700;
        font-size: 2.5rem;
        margin-bottom: 0.5rem;
    }
    
    h2 {
        color: #0A2540;
        font-weight: 600;
        font-size: 1.75rem;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
    
    h3 {
        color: #0A2540;
        font-weight: 600;
        font-size: 1.25rem;
        margin-top: 1.5rem;
        margin-bottom: 0.75rem;
    }
    
    /* Buttons */
    .stButton > button {
        background-color: #5B5BFF;
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        font-weight: 500;
        font-size: 1rem;
        border-radius: 8px;
        transition: all 0.2s ease;
        box-shadow: 0 2px 4px rgba(91, 91, 255, 0.2);
    }
    
    .stButton > button:hover {
        background-color: #4B4BEF;
        box-shadow: 0 4px 8px rgba(91, 91, 255, 0.3);
        transform: translateY(-1px);
    }
    
    /* Download button special styling */
    .stDownloadButton > button {
        background-color: #10B981;
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        font-weight: 500;
        border-radius: 8px;
        transition: all 0.2s ease;
    }
    
    .stDownloadButton > button:hover {
        background-color: #059669;
        transform: translateY(-1px);
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 2rem;
        background-color: transparent;
        border-bottom: 2px solid #E5E7EB;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 3rem;
        background-color: transparent;
        border: none;
        color: #6B7280;
        font-weight: 500;
        font-size: 1.1rem;
        padding: 0 1rem;
        border-radius: 0;
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        color: #0A2540;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: transparent;
        color: #5B5BFF;
        border-bottom: 3px solid #5B5BFF;
        font-weight: 600;
    }
    
    /* File uploader */
    .uploadedFile {
        background-color: #F9FAFB;
        border: 2px solid #E5E7EB;
        border-radius: 8px;
        padding: 1rem;
    }
    
    [data-testid="stFileUploader"] {
        background-color: #F9FAFB;
        border: 2px dashed #D1D5DB;
        border-radius: 12px;
        padding: 2rem;
        transition: all 0.2s ease;
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: #5B5BFF;
        background-color: #F5F5FF;
    }
    
    /* Select boxes */
    .stSelectbox > div > div {
        background-color: white;
        border: 1px solid #D1D5DB;
        border-radius: 8px;
    }
    
    .stSelectbox > div > div:hover {
        border-color: #5B5BFF;
    }
    
    /* Data editor */
    .stDataFrame {
        border: 1px solid #E5E7EB;
        border-radius: 8px;
        overflow: hidden;
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background-color: #F9FAFB;
        border: 1px solid #E5E7EB;
        border-radius: 8px;
        font-weight: 500;
        color: #0A2540;
    }
    
    .streamlit-expanderHeader:hover {
        background-color: #F3F4F6;
    }
    
    /* Info/Warning/Success boxes */
    .stAlert {
        border-radius: 8px;
        border: 1px solid;
        padding: 1rem;
    }
    
    div[data-testid="stInfo"] {
        background-color: #EFF6FF;
        border-color: #5B5BFF;
        color: #1E40AF;
    }
    
    div[data-testid="stWarning"] {
        background-color: #FEF3C7;
        border-color: #F59E0B;
        color: #92400E;
    }
    
    div[data-testid="stSuccess"] {
        background-color: #D1FAE5;
        border-color: #10B981;
        color: #065F46;
    }
    
    /* Multiselect */
    .stMultiSelect > div > div {
        background-color: white;
        border: 1px solid #D1D5DB;
        border-radius: 8px;
    }
    
    /* Radio buttons */
    .stRadio > div {
        background-color: #F9FAFB;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #E5E7EB;
    }
    
    /* Date input */
    .stDateInput > div > div {
        background-color: white;
        border: 1px solid #D1D5DB;
        border-radius: 8px;
    }
    
    /* Text input */
    .stTextInput > div > div > input {
        border: 1px solid #D1D5DB;
        border-radius: 8px;
        padding: 0.5rem 0.75rem;
    }
    
    .stTextInput > div > div > input:focus {
        border-color: #5B5BFF;
        box-shadow: 0 0 0 3px rgba(91, 91, 255, 0.1);
    }
    
    /* Section styling */
    .section-container {
        background-color: white;
        padding: 2rem;
        border-radius: 12px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
        margin-bottom: 1.5rem;
    }
    
    /* Step indicators */
    .step-indicator {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 32px;
        height: 32px;
        background-color: #5B5BFF;
        color: white;
        border-radius: 50%;
        font-weight: 600;
        margin-right: 0.75rem;
    }
    
    /* Custom divider */
    .custom-divider {
        height: 1px;
        background-color: #E5E7EB;
        margin: 2rem 0;
    }
    
    /* Header styling */
    .header-container {
        background-color: white;
        padding: 1.5rem 0;
        border-bottom: 2px solid #E5E7EB;
        margin-bottom: 2rem;
    }
    
    .logo-title-container {
        display: flex;
        align-items: center;
        gap: 1.5rem;
    }
    
    .title-text {
        color: #0A2540;
        font-size: 2rem;
        font-weight: 700;
        margin: 0;
    }
    
    /* Text area */
    .stTextArea textarea {
        border: 1px solid #D1D5DB;
        border-radius: 8px;
        font-family: 'Monaco', 'Menlo', monospace;
    }
    
    /* Code blocks */
    .stCodeBlock {
        background-color: #F9FAFB;
        border: 1px solid #E5E7EB;
        border-radius: 8px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Header with logo
st.markdown('<div class="header-container">', unsafe_allow_html=True)
col1, col2 = st.columns([1, 11])
with col1:
    this_dir = os.path.dirname(os.path.abspath(__file__))
    logo_path = os.path.join(this_dir, "logo.png")
    st.image(logo_path, width=50)
with col2:
    st.markdown('<h1 class="title-text">Excel Merge Tool 2.0</h1>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# Tab selection with better styling (show Incomplete Import first by default)
tab_incomplete, tab_complete = st.tabs(["📝 Incomplete Import File", "📊 Complete Import File"])

# Store the active tab
if tab_incomplete:
    tab = "Distributions: Incomplete Import File"
else:
    tab = "Distributions: Complete Import File"

# ---------------------------------------------------------------
# Helper to load Excel workbooks but show a friendly message when
# Covercy's raw template XML breaks openpyxl. Returns the workbook
# or None; callers should `return` early if None.
# ---------------------------------------------------------------


def _safe_load_workbook(file_like):
    """Attempt to load an XLSX; on parse failure, show guidance and return None."""
    try:
        return load_workbook(file_like, data_only=False)
    except (ValueError, InvalidFileException, zipfile.BadZipFile) as e:
        st.error("⚠️  This template can’t be read as-is. Please open it in Excel, choose “Save As…”, and upload the saved copy.")
        st.caption(f"Details: {e}")
        return None

# === Complete flow ===
def run_complete_flow():
    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
    
    # Instructions in a collapsible section
    with st.expander("📖 Instructions - How to Use This Tool", expanded=False):
        st.markdown(
            """
            ### Step-by-Step Guide
            
            1. **Prepare Your Source Data**
               - Ensure data is in the first tab of your spreadsheet
               - Distribution dates, investor names, and amounts should be in separate columns
               - Add clear headers above your data (e.g., "Date", "Investing Entity", "Amount")
               - Close the spreadsheet before uploading
            
            2. **Upload Files**
               - Upload your source data spreadsheet
               - Upload the Covercy import template
            
            3. **Map Your Data**
               - Select which columns contain your dates, entities, and amounts
               - Review and adjust entity name mappings
               - Choose how to handle distribution types
            
            4. **Review & Download**
               - Resolve any duplicate amounts
               - Download the completed import file
               - Upload to Covercy
            
            💡 **Pro Tip:** For best results, ensure investor names in your source file match Covercy's naming conventions.
            """
        )
    
    # Step 1: File Upload Section

    # --- New mode selector ---------------------------------------------------
    saved_src_ok = st.session_state.get("shared_source_file") is not None
    saved_tgt_ok = st.session_state.get("shared_target_file_bytes") is not None

    default_idx = 0 if (saved_src_ok or saved_tgt_ok) else 1
    file_mode = st.radio(
        "File selection",
        ("↪️  Use files from Incomplete flow", "📂 Upload new files"),
        index=default_idx,
        horizontal=True,
        key="comp_file_mode",
    )

    use_saved = file_mode.startswith("↪️")
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2><span class="step-indicator">1</span>Upload Files</h2>', unsafe_allow_html=True)
 
    col1, col2 = st.columns(2)
    with col1:
        shared = st.session_state.get("shared_source_file")
        if use_saved and shared is not None:
            st.success(f"Using Source File from Incomplete flow: {shared.name}")
            source_file = shared
        else:
            st.markdown("##### Source Excel File")
            st.markdown("<small style='color: #6B7280;'>Your data spreadsheet with distributions</small>", unsafe_allow_html=True)
            source_file = st.file_uploader("", type=["xlsx","xls"], key="src", label_visibility="collapsed")
            if source_file is not None:
                # Persist for future use
                st.session_state["shared_source_file"] = source_file
                # If user is uploading new files, ensure future default switches to saved mode
                use_saved = False
 
    with col2:
        shared_tgt = st.session_state.get("shared_target_file_bytes")
        shared_name = st.session_state.get("shared_target_file_name", "populated_template.xlsx")

        if use_saved and shared_tgt is not None:
            st.success(f"Using Populated Template from Incomplete flow: {shared_name}")
            target_file = io.BytesIO(shared_tgt)
            target_file.name = shared_name  # type: ignore
        else:
            st.markdown("##### Target Excel File")
            st.markdown("<small style='color: #6B7280;'>Covercy import template</small>", unsafe_allow_html=True)
            target_file = st.file_uploader("", type=["xlsx","xls"], key="tgt", label_visibility="collapsed")
            if target_file is not None:
                # Persist upload for Complete flow reuse
                st.session_state["shared_target_file_bytes"] = target_file.read()
                st.session_state["shared_target_file_name"] = target_file.name
                target_file.seek(0)
                use_saved = False

    st.markdown('</div>', unsafe_allow_html=True)
    
    if not (source_file and target_file):
        st.info("👆 Please upload both files to continue")
        return

    if source_file is None:
        st.info("👆 Please upload the source file to continue")
        return

    # Persist the choice (if user re-uploaded in Complete flow)
    st.session_state["shared_source_file"] = source_file

    # Read source file
    df_source = pd.read_excel(source_file)
    
    # Step 2: Data Preview & Column Selection
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2><span class="step-indicator">2</span>Configure Data Mapping</h2>', unsafe_allow_html=True)
    
    # Preview
    st.markdown("##### Source Data Preview")
    st.dataframe(df_source.head(5), use_container_width=True)
    
    # Column selection
    st.markdown("##### Select Data Columns")
    cols = df_source.columns.tolist()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        src_ent = st.selectbox("Investing Entity column", cols, help="Column containing investor/entity names")
    with col2:
        src_dt = st.selectbox("Date column", cols, help="Column containing distribution dates")
    with col3:
        src_amt = st.selectbox("Amount column", cols, help="Column containing distribution amounts")
    
    st.markdown('</div>', unsafe_allow_html=True)

    # Parse dates
    df_source['parsed_date'] = pd.to_datetime(df_source[src_dt], errors='coerce').dt.date
    invalid = df_source['parsed_date'].isna().sum()
    if invalid:
        st.warning(f"⚠️ {invalid} rows have unparseable dates and will be skipped.")

    # Inspect target sheet layout
    try:
        df_raw = pd.read_excel(target_file, header=None)
    except (ValueError, zipfile.BadZipFile, InvalidFileException) as e:
        st.error("⚠️  This template can’t be read as-is. Open it in Excel, “Save As…”, then upload the saved copy.")
        st.caption(f"Details: {e}")
        return
    ent_col = 2
    # Normalize Column C values for robust matching
    colC_series = df_raw[ent_col].astype(str).str.strip()
    ent_label_row = colC_series[colC_series == 'Investing Entity'].index[0]
    # Prefer explicit GP-like footer; fallback to first blank after header
    footer_slice = colC_series.iloc[ent_label_row + 1 :]
    footer_norm = footer_slice.fillna('').astype(str).str.strip()
    gp_like = footer_norm[footer_norm.str.fullmatch(r'(?i)GP/Remaining Funds')]
    if not gp_like.empty:
        gp_row = gp_like.index[0]
    else:
        blanks = footer_norm[footer_norm == ""]
        if not blanks.empty:
            gp_row = blanks.index[0]
        else:
            # As a last resort, use the last row so we don't crash
            gp_row = footer_slice.index[-1]
    ent_rows      = list(range(ent_label_row + 1, gp_row))
    target_entities = [str(df_raw.iat[r, ent_col]).strip() for r in ent_rows]

    # Build (date, type) -> starting-column map
    date_label_row = ent_label_row - 2
    date_cols = [j for j in range(df_raw.shape[1])
                 if str(df_raw.iat[date_label_row, j]).strip() == 'Last Day']

    date_type_map = {}
    for j in date_cols:
        dist_date = pd.to_datetime(df_raw.iat[date_label_row + 1, j], errors='coerce').date()
        
        try:
            dist_type_raw = str(df_raw.iat[date_label_row + 1, j + 1]).strip()
        except IndexError:
            dist_type_raw = ""

        if dist_type_raw == "" or dist_type_raw == "-":
            dist_type_raw = str(df_raw.iat[date_label_row + 1, j + 2]).strip()

        dist_type = dist_type_raw or ""
        date_type_map[(dist_date, dist_type)] = j

    # Step 3: Entity Mapping
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2><span class="step-indicator">3</span>Map Entities</h2>', unsafe_allow_html=True)
    
    unique_src = df_source[src_ent].dropna().astype(str).unique().tolist()
    map_df = pd.DataFrame({'source_entity': unique_src})
    map_df['suggestion_1'] = map_df['source_entity'].apply(
        lambda x: (difflib.get_close_matches(x, target_entities, n=1, cutoff=0.6) or [""])[0]
    )
    map_df['target_entity'] = map_df['suggestion_1']

    st.markdown("##### Match source entities to target entities")
    st.markdown("<small style='color: #6B7280;'>Review and adjust the suggested mappings below. Only mapped entities will be included in the final import.</small>", unsafe_allow_html=True)
    
    edited = st.data_editor(
        map_df,
        column_config={
            'source_entity': st.column_config.TextColumn(
                "Source Entity",
                help="Entity names from your source file",
                disabled=True,
            ),
            'suggestion_1': st.column_config.TextColumn(
                "Suggested Match",
                help="Our best guess based on name similarity",
                disabled=True,
            ),
            'target_entity': st.column_config.SelectboxColumn(
                "Target Entity",
                help="Select the exact entity name from the Covercy template",
                options=[""] + target_entities,
                required=False,
            )
        }, 
        hide_index=True,
        use_container_width=True
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # Step 4: Distribution Type Handling
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2><span class="step-indicator">4</span>Distribution Types</h2>', unsafe_allow_html=True)

    # Determine default based on presence of saved mapping tokens
    has_saved_token = bool(st.session_state.get("token_history"))
    default_radio_idx = 1 if has_saved_token else 0

    type_handling = st.radio(
        "How should distribution types be handled?",
        ("🔲 Single type - Ignore distribution types", "🔷 Multiple types - Match distribution types"),
        index=default_radio_idx,
        key="comp_type_mode",
        horizontal=True
    )

    ignore_types = type_handling.startswith("🔲")

    dist_options = [
        "Preferred Return", "Interest", "Profit", "Return of Capital",
        "Principal", "Promote", "Catch Up", "Available Cash (Profit)"
    ]

    if ignore_types:
        df_source['mapped_type'] = ""
        collapsed_map = {}
        for (d, t), col_idx in date_type_map.items():
            if (d, "") not in collapsed_map:
                collapsed_map[(d, "")] = col_idx
        date_type_map = collapsed_map
        st.info("ℹ️ Distribution types will be ignored; amounts will be matched only by entity and date.")

    else:
        type_cols = [c for c in cols if c not in [src_ent, src_dt, src_amt]]
        type_cols_display = ["<No type column>"] + type_cols

        # Pre-select any column remembered from the Incomplete flow
        default_type_col = st.session_state.get("shared_type_column", "<No type column>")
        default_idx = type_cols_display.index(default_type_col) if default_type_col in type_cols_display else 0

        chosen_type_col = st.selectbox(
            "Select Distribution Type column from source",
            options=type_cols_display,
            index=default_idx,
            key="comp_dist_type_col",
            help="Leave as '<No type column>' if your source doesn't have distribution types"
        )
        # Keep the two tabs in sync going forward
        if chosen_type_col != "<No type column>":
            st.session_state["shared_type_column"] = chosen_type_col

        # Token handling section
        st.markdown("##### Mapping Configuration")
        col1, col2 = st.columns([3, 1])

        # Ensure token_select is always defined
        token_select = ""
        
        with col1:
            saved_tokens = st.session_state.get("token_history", [])
            
            if saved_tokens and "comp_mapping_token" not in st.session_state:
                st.session_state["comp_mapping_token"] = saved_tokens[-1]

            token_input = st.text_input(
                "Paste Mapping Token (optional)",
                key="comp_mapping_token",
                placeholder="Paste a token from the Incomplete flow to reuse mappings",
                help="If you've already configured type mappings in the Incomplete flow, paste the token here"
            )

        with col2:
            if saved_tokens:
                token_options = [""] + list(reversed(saved_tokens))
                
                def _tok_label(tok: str) -> str:
                    if tok == "":
                        return "Select a token..."
                    if saved_tokens and tok == saved_tokens[-1]:
                        return "★ Most Recent"
                    idx = list(reversed(saved_tokens)).index(tok)
                    return f"Token {idx+1}"

                token_select = st.selectbox(
                    "Or select recent token",
                    token_options,
                    format_func=_tok_label,
                    key="comp_mapping_token_select",
                )

        chosen_token = token_select if token_select else token_input.strip()

        mapping_from_token = {}
        if chosen_token:
            try:
                decoded = base64.urlsafe_b64decode(chosen_token.encode()).decode()
                mapping_from_token = json.loads(decoded)
                st.success("✅ Mapping token applied successfully")
            except Exception:
                st.error("❌ Invalid mapping token")

        if chosen_type_col == "<No type column>":
            df_source['mapped_type'] = ""
        else:
            if mapping_from_token:
                type_mapping = mapping_from_token
                df_source['mapped_type'] = df_source[chosen_type_col].map(type_mapping).fillna("")
            else:
                unique_src_types = df_source[chosen_type_col].dropna().astype(str).unique().tolist()
                type_df = pd.DataFrame({'source_type': unique_src_types})
                type_df['target_type'] = type_df['source_type'].apply(
                    lambda x: (difflib.get_close_matches(x, dist_options, n=1, cutoff=0.6) or [""])[0]
                )

                st.markdown("##### Map distribution types")
                edited_type_df = st.data_editor(
                    type_df,
                    column_config={
                        'source_type': st.column_config.TextColumn(
                            "Source Type",
                            help="Distribution types from your file",
                            disabled=True
                        ),
                        'target_type': st.column_config.SelectboxColumn(
                            "Target Type",
                            help="Select the Covercy distribution type",
                            options=[""] + dist_options,
                            required=False,
                        ),
                    },
                    hide_index=True,
                    use_container_width=True,
                    key="comp_dist_type_mapper"
                )

                type_mapping = dict(zip(edited_type_df['source_type'], edited_type_df['target_type']))
                df_source['mapped_type'] = df_source[chosen_type_col].map(type_mapping).fillna("")
    
    st.markdown('</div>', unsafe_allow_html=True)

    # Step 5: Duplicate Resolution
    mapping = dict(zip(edited['source_entity'], edited['target_entity']))
    df_source['mapped_entity'] = df_source[src_ent].map(mapping)

    valid_pairs = set(date_type_map.keys())

    dup_src = df_source[
        (df_source['mapped_entity'] != "") &
        df_source.apply(lambda row: (row['parsed_date'], row['mapped_type']) in valid_pairs, axis=1)
    ]

    dup_groups = dup_src.groupby(
        ['mapped_entity', 'parsed_date', 'mapped_type']
    )[src_amt].apply(list).reset_index(name='amounts')
    dups = dup_groups[dup_groups['amounts'].apply(len) > 1]
    chosen = {}

    if not dups.empty:
        st.markdown('<div class="section-container">', unsafe_allow_html=True)
        st.markdown('<h2><span class="step-indicator">5</span>Resolve Duplicates</h2>', unsafe_allow_html=True)
        st.markdown("<small style='color: #6B7280;'>Multiple amounts found for the same entity/date/type combination. Choose how to handle each:</small>", unsafe_allow_html=True)

        if st.button("🔢 Sum All Duplicates", use_container_width=True):
            for _, row in dups.iterrows():
                key = f"dup_{row['mapped_entity']}_{row['parsed_date']}_{row['mapped_type']}"
                st.session_state[key] = 'SUM'

        for _, row in dups.iterrows():
            ent, dt, typ, amts = row['mapped_entity'], row['parsed_date'], row['mapped_type'], row['amounts']
            key = f"dup_{ent}_{dt}_{typ}"
            options = [str(a) for a in amts] + ['SUM']
            if key not in st.session_state:
                st.session_state[key] = 'SUM'
            
            type_str = f" ({typ})" if typ else ""
            sel = st.radio(
                f"**{ent}** on {dt}{type_str}", 
                options, 
                key=key,
                horizontal=True
            )
            sel = st.session_state[key]
            chosen[(ent, dt, typ)] = sum(amts) if sel == 'SUM' else float(sel)
        
        st.markdown('</div>', unsafe_allow_html=True)

    # Step 6: Finalize
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2><span class="step-indicator">6</span>Finalize & Download</h2>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown("##### Ready to generate your import file?")
        st.markdown("<small style='color: #6B7280;'>Click below to process your data and create the final import file.</small>", unsafe_allow_html=True)
    
    if st.button("✨ Generate Import File", use_container_width=True, type="primary"):
        target_file.seek(0)
        wb = _safe_load_workbook(target_file)
        if wb is None:
            return
        ws = wb[wb.sheetnames[0]]
        unmatched=[]
        
        # Progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_entities = len(target_entities)
        processed = 0
        
        for r,ent in zip(ent_rows, target_entities):
            for (dist_date, dist_type), col_idx in date_type_map.items():
                if pd.isna(dist_date):
                    continue
                m = df_source[
                    (df_source['mapped_entity'] == ent) &
                    (df_source['parsed_date'] == dist_date) &
                    (df_source['mapped_type'] == dist_type)
                ]
                if not m.empty:
                    amt = chosen.get((ent, dist_date, dist_type), m[src_amt].iloc[0])
                    ws.cell(row=r + 1, column=col_idx).value = amt
                else:
                    unmatched.append((ent, dist_date, dist_type))
            
            processed += 1
            progress_bar.progress(processed / total_entities)
            status_text.text(f"Processing entity {processed}/{total_entities}...")
        
        progress_bar.empty()
        status_text.empty()
        
        buf=io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        
        st.success("✅ Import file generated successfully!")
        
        col1, col2 = st.columns([1, 2])
        with col1:
            st.download_button(
                "📥 Download Import File",
                data=buf,
                file_name=f"covercy_import_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        if unmatched:
            with st.expander(f"⚠️ {len(unmatched)} unmatched entries", expanded=False):
                st.write("The following entity/date/type combinations were not found in the source data:")
                unmatched_df = pd.DataFrame(unmatched, columns=['Entity', 'Date', 'Type'])
                st.dataframe(unmatched_df, use_container_width=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

# === Incomplete flow ===
def run_incomplete_flow():
    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
    
    # Instructions in a collapsible section
    with st.expander("📖 Instructions - How to Use This Tool", expanded=False):
        st.markdown(
            """
            ### Overview
            This tool helps you create a complete import file when you only have a template with a single distribution date.
            
            ### Step-by-Step Guide
            
            1. **Prepare Template**
               - Generate a Covercy import file with a single distribution date
               - Use a date BEFORE your actual distributions to avoid duplicates
            
            2. **Prepare Source Data**
               - Ensure dates, entities, and amounts are in separate columns
               - Add clear headers to your data
               - Close the file before uploading
            
            3. **Upload & Configure**
               - Upload your source data and single-date template
               - Select which dates to include
               - Configure distribution types
            
            4. **Generate & Use**
               - Download the populated template
               - Use it in the "Complete Import File" tab
            
            ⚠️ **Important:** Due to Covercy limitations, only ~60% of periods may import at once. You may need to repeat the process for remaining dates.
            """
        )
    
    # Step 1: File Upload
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2><span class="step-indicator">1</span>Upload Files</h2>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("##### Source Excel File")
        st.markdown("<small style='color: #6B7280;'>Your data with all distribution dates</small>", unsafe_allow_html=True)

        uploaded_src = st.file_uploader("", type=["xlsx","xls"], key="inc_src", label_visibility="collapsed")

        if uploaded_src is not None:
            # User picked a new file – use it and cache for Complete flow
            source_file = uploaded_src
            st.session_state["shared_source_file"] = uploaded_src
        else:
            shared = st.session_state.get("shared_source_file")
            if shared is not None:
                st.info(f"Using previously uploaded file: {shared.name}")
                source_file = shared
            else:
                source_file = None
    
    with col2:
        st.markdown("##### Incomplete Target File")
        st.markdown("<small style='color: #6B7280;'>Covercy template with single date - please Save As before uploading</small>", unsafe_allow_html=True)
        target_file = st.file_uploader("", type=["xlsx","xls"], key="inc_tgt", label_visibility="collapsed")

        # Immediate validation so any parse error is shown here (no scrolling needed)
        if target_file is not None:
            wb_test = _safe_load_workbook(io.BytesIO(target_file.getvalue()))
            if wb_test is None:
                st.stop()
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    if not (source_file and target_file):
        st.info("👆 Please upload both files to continue")
        return

    # Save source file to share with Complete flow
    st.session_state["shared_source_file"] = source_file

    # Read & parse source
    df_src = pd.read_excel(source_file)
    
    # Step 2: Configure Source Data
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2><span class="step-indicator">2</span>Configure Source Data</h2>', unsafe_allow_html=True)
    
    st.markdown("##### Source Data Preview")
    st.dataframe(df_src.head(5), use_container_width=True)

    st.markdown("##### Select Date Column")
    inc_date_col = st.selectbox(
        "Which column contains distribution dates?",
        options=df_src.columns.tolist(),
        key="inc_date_col",
        help="Select the column that contains the distribution dates"
    )

    # Parse dates
    df_src['parsed_date'] = pd.to_datetime(
        df_src[inc_date_col], errors='coerce'
    ).dt.date

    invalid = df_src['parsed_date'].isna().sum()
    if invalid:
        st.warning(f"⚠️ {invalid} rows have unparseable dates and will be skipped.")
    
    st.markdown('</div>', unsafe_allow_html=True)

    # Load target workbook
    data = target_file.read()
    wb = _safe_load_workbook(io.BytesIO(data))
    if wb is None:
        return
    ws = wb.active

    # Locate entity rows
    colC_raw = [c.value for c in ws['C']]
    colC_norm = [str(v).strip() if v is not None else "" for v in colC_raw]
    # Zero-based indices for searching
    ent_label_row_zb = None
    try:
        ent_label_row_zb = colC_norm.index("Investing Entity")
    except ValueError:
        st.error("❌ Could not locate 'Investing Entity' header in column C")
        return
    # Find footer row (zero-based): prefer explicit "GP/Remaining Funds", else first blank
    try:
        gp_row_zb = next(
            i for i in range(ent_label_row_zb + 1, len(colC_norm))
            if colC_norm[i].strip().lower() == "gp/remaining funds"
        )
    except StopIteration:
        try:
            gp_row_zb = next(
                i for i in range(ent_label_row_zb + 1, len(colC_norm))
                if colC_norm[i] == ""
            )
        except StopIteration:
            st.error("❌ Could not locate footer (GP/Remaining Funds or blank row) after 'Investing Entity'")
            return
    # Convert to 1-based Excel row numbers
    ent_label_row = ent_label_row_zb + 1
    gp_row        = gp_row_zb + 1
    entity_rows   = list(range(ent_label_row+1, gp_row+1))

    # Copy first block headers
    # Determine first block start dynamically:
    # The header row that contains 'Last Day' is two rows above the entity header.
    width = 7
    last_day_row = max(1, ent_label_row - 2)
    last_day_cols = [
        j for j in range(1, ws.max_column + 1)
        if str(ws.cell(row=last_day_row, column=j).value).strip() == "Last Day"
    ]
    if last_day_cols:
        # Amounts start one column to the left of 'Last Day'
        first_col = max(1, last_day_cols[0] - 1)
    else:
        # Fallback to legacy starting column (F)
        first_col = 6
    hdr1 = [ws.cell(row=1, column=c).value for c in range(first_col, first_col+width)]
    hdr3 = [ws.cell(row=3, column=c).value for c in range(first_col, first_col+width)]
    hdr5 = [ws.cell(row=5, column=c).value for c in range(first_col, first_col+width)]

    # Build new-dates list
    # Use the date directly under the first detected 'Last Day' if available
    if last_day_cols:
        existing = ws.cell(row=last_day_row + 1, column=last_day_cols[0]).value
    else:
        existing = ws.cell(row=4, column=first_col+1).value
    uniq = df_src['parsed_date'].unique()
    new_dates = sorted([d for d in uniq if pd.notna(d) and d != existing])

    # Step 3: Select Dates
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2><span class="step-indicator">3</span>Select Distribution Dates</h2>', unsafe_allow_html=True)
    
    if not new_dates:
        st.error("No valid dates found in the source file!")
        return
    
    # Date range selector
    st.markdown("##### Filter by Date Range")
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input(
            "Start date",
            value=new_dates[0],
            min_value=new_dates[0],
            max_value=new_dates[-1],
            key='inc_start_date'
        )
    with col2:
        end_date = st.date_input(
            "End date",
            value=new_dates[-1],
            min_value=new_dates[0],
            max_value=new_dates[-1],
            key='inc_end_date'
        )
    
    # Filter dates based on range
    dates_in_range = [d for d in new_dates if start_date <= d <= end_date]
    
    st.info(f"📅 Found {len(dates_in_range)} dates in the selected range")
    
    # Date selection buttons
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("✓ Select All", use_container_width=True):
            st.session_state['sel_dates'] = [d.strftime("%Y-%m-%d") for d in dates_in_range]
    with col2:
        if st.button("✗ Clear All", use_container_width=True):
            st.session_state['sel_dates'] = []
    
    # Initialize selection
    date_strs = [d.strftime("%Y-%m-%d") for d in dates_in_range]
    if 'sel_dates' not in st.session_state:
        st.session_state['sel_dates'] = date_strs.copy()
    
    # Multiselect for fine-tuning
    with st.expander("Fine-tune date selection", expanded=False):
        selected = st.multiselect(
            "Select specific dates:",
            options=date_strs,
            default=st.session_state.get('sel_dates', date_strs),
            key='sel_dates',
            format_func=lambda x: datetime.strptime(x, "%Y-%m-%d").strftime("%b %d, %Y")
        )
    
    dates_to_use = sorted([d for d in dates_in_range if d.strftime("%Y-%m-%d") in selected])
    
    st.markdown('</div>', unsafe_allow_html=True)

    # Step 4: Distribution Types
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2><span class="step-indicator">4</span>Configure Distribution Types</h2>', unsafe_allow_html=True)

    type_handling = st.radio(
        "How should distribution types be determined?",
        ("🔲 Single type for all distributions", "🔷 Map types from source column"),
        index=0,
        key="dist_type_handling",
        horizontal=True
    )

    dist_options = [
        "Preferred Return", "Interest", "Profit", "Return of Capital",
        "Principal", "Promote", "Catch Up", "Available Cash (Profit)"
    ]

    dates_and_types_to_use = []

    if type_handling.startswith("🔲"):
        dist_type = st.selectbox(
            "Select Distribution Type for all periods",
            options=dist_options,
            index=0,
            help="This type will be applied to all distribution periods"
        )
        dates_and_types_to_use = [(d, dist_type) for d in dates_to_use]
        
    else:  # Map types from source
        # Select type column
        type_cols = [c for c in df_src.columns if c not in [inc_date_col, 'parsed_date']]
        
        if not type_cols:
            st.error("No additional columns available for distribution types!")
            return
            
        inc_dist_type_col = st.selectbox(
            "Select Distribution Type column",
            options=type_cols,
            key="inc_dist_type_col",
            help="Column that specifies the distribution type for each row"
        )
        # Store selected column so the Complete flow can pick it up automatically
        if inc_dist_type_col:
            st.session_state["shared_type_column"] = inc_dist_type_col

        if inc_dist_type_col:
            # --- Build initial mapping df ---
            unique_src_types = df_src[inc_dist_type_col].dropna().astype(str).unique().tolist()

            # Retrieve existing in-session mapping (persists across reruns)
            mapping_dict = st.session_state.get("inc_type_mapping", {})

            def _initial_target(src: str) -> str:
                # Prefer saved mapping; otherwise suggest best match
                if src in mapping_dict:
                    return mapping_dict[src]
                return (difflib.get_close_matches(src, dist_options, n=1, cutoff=0.6) or [""])[0]

            type_map_df = pd.DataFrame({
                'source_type': unique_src_types,
                'target_type': [ _initial_target(s) for s in unique_src_types ]
            })

            # --- Bulk-assign helper UI ---
            st.markdown("##### Bulk assign distribution types")
            fil_col, assign_col = st.columns([3,2])

            with fil_col:
                filter_text = st.text_input(
                    "Filter source types (case-insensitive contains)",
                    key="inc_type_filter",
                    placeholder="e.g. pref"
                )

            with assign_col:
                bulk_choice = st.selectbox(
                    "Set filtered to…",
                    [""] + dist_options,
                    key="inc_bulk_type_choice"
                )

            if st.button("Apply to filtered rows", key="inc_apply_bulk"):
                if filter_text and bulk_choice:
                    matches = [s for s in unique_src_types if filter_text.lower() in s.lower()]
                    for s in matches:
                        mapping_dict[s] = bulk_choice
                    # Reflect change in current session and dataframe for immediate feedback
                    st.session_state["inc_type_mapping"] = mapping_dict
                    type_map_df.loc[type_map_df['source_type'].isin(matches), 'target_type'] = bulk_choice
                    st.success(f"Assigned '{bulk_choice}' to {len(matches)} source types")
 
            st.markdown("##### Map source types to Covercy types")
            # Apply live filter to table view
            display_df = type_map_df if not filter_text else type_map_df[type_map_df['source_type'].str.contains(filter_text, case=False)]

            edited_type_map = st.data_editor(
                display_df,
                column_config={
                    'source_type': st.column_config.TextColumn(
                        "Source Type",
                        help="Types found in your file",
                        disabled=True
                    ),
                    'target_type': st.column_config.SelectboxColumn(
                        "Target Type",
                        help="Select the Covercy type",
                        options=[""] + dist_options,
                        required=False,
                    )
                },
                hide_index=True,
                use_container_width=True,
                key="inc_dist_type_mapper"
            )

            # Merge manual edits into mapping_dict
            for src, tgt in zip(edited_type_map['source_type'], edited_type_map['target_type']):
                if tgt:
                    mapping_dict[src] = tgt
                elif src in mapping_dict:
                    del mapping_dict[src]

            st.session_state["inc_type_mapping"] = mapping_dict

            # Final mapping used downstream
            type_mapping = mapping_dict.copy()
            df_src['mapped_type'] = df_src[inc_dist_type_col].map(type_mapping)

            
            # Generate token
            mapping_token = base64.urlsafe_b64encode(json.dumps(type_mapping).encode()).decode()
            
            st.markdown("##### Mapping Token")
            st.text_area(
                "Copy this token to reuse these mappings in the Complete flow",
                mapping_token,
                height=80,
                help="Save this token to avoid remapping types when using the Complete Import File tab"
            )

            # Save token to history
            hist = st.session_state.get("token_history", [])
            if mapping_token not in hist:
                hist.append(mapping_token)
                st.session_state["token_history"] = hist[-10:]

            # Build date-type pairs
            dates_to_use_set = set(dates_to_use)
            relevant_rows = df_src[
                df_src['parsed_date'].isin(dates_to_use_set)
                & df_src['mapped_type'].notna()
                & (df_src['mapped_type'] != "")
            ]
            unique_pairs = relevant_rows[['parsed_date', 'mapped_type']].drop_duplicates()
            dates_and_types_to_use = sorted(
                [tuple(x) for x in unique_pairs.to_numpy()],
                key=lambda x: x[0],
            )
    
    st.markdown('</div>', unsafe_allow_html=True)

    # Step 5: Generate Template
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2><span class="step-indicator">5</span>Generate Template</h2>', unsafe_allow_html=True)
    
    # 56% rule handling
    total_needed = ceil(len(dates_and_types_to_use) / 0.56)
    extra_needed = total_needed - len(dates_and_types_to_use)
    
    if extra_needed > 0:
        placeholder_date = date(2040, 1, 1)
        placeholder_type = "Preferred Return"
        dates_and_types_to_use.extend([(placeholder_date, placeholder_type)] * extra_needed)
        st.info(f"ℹ️ Added {extra_needed} placeholder periods to accommodate Covercy's import limitations")
    
    st.markdown("##### Summary")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Selected Dates", len([d for d in dates_and_types_to_use if d[0].year != 2040]))
    with col2:
        st.metric("Placeholder Dates", extra_needed)
    with col3:
        st.metric("Total Periods", len(dates_and_types_to_use))
    
    if st.button("🚀 Generate Populated Template", use_container_width=True, type="primary"):
        # Make newest token the default for the Complete flow
        if 'mapping_token' in locals() and mapping_token and mapping_token != "e30=":
            st.session_state["comp_mapping_token"] = mapping_token
            st.session_state.pop("comp_mapping_token_select", None)

        # Progress tracking
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # Append blocks
        for idx, (last_day, dist_type) in enumerate(dates_and_types_to_use):
            base = first_col + width*(idx+1)
            
            # Update progress
            progress = (idx + 1) / len(dates_and_types_to_use)
            progress_bar.progress(progress)
            status_text.text(f"Adding distribution period {idx+1}/{len(dates_and_types_to_use)}...")
            
            # Headers
            for r,vals in zip([1,3,5],[hdr1,hdr3,hdr5]):
                for j,v in enumerate(vals):
                    ws.cell(row=r, column=base+j).value = v
            
            # Row 2 - date range
            short = f"{last_day.day} {last_day.strftime('%b')} {last_day.year}"
            ws.cell(row=2, column=base   ).value = f"{short} - {short}"
            ws.cell(row=2, column=base+1 ).value = "Custom"
            ws.cell(row=2, column=base+2 ).value = "-"
            ws.cell(row=2, column=base+3 ).value = datetime.now().year
            
            # Row 4 - full date and type
            full = f"{last_day.day} {last_day.strftime('%B')} {last_day.year}"
            ws.cell(row=4, column=base   ).value = full
            ws.cell(row=4, column=base+1 ).value = full
            ws.cell(row=4, column=base+2 ).value = dist_type
            ws.cell(row=4, column=base+3 ).value = "USD"
            
            # Payment dates
            pay_col = base+5
            dt_val  = datetime(last_day.year,last_day.month,last_day.day)
            for r in entity_rows:
                cell_pd = ws.cell(row=r, column=pay_col)
                cell_pd.value = dt_val
                cell_pd.number_format = 'm/d/yyyy'
            
            # GP formula
            prom = base+2
            let = get_column_letter(prom)
            s,e = entity_rows[0],entity_rows[-2]
            ws.cell(row=entity_rows[-1], column=base).value = f"=SUM({let}{s}:{let}{e})"
            
            # Net formulas
            g,t,p,a,n = base,base+1,base+2,base+3,base+4
            for r in entity_rows[:-1]:
                expr = (f"=SUM({get_column_letter(g)}{r},-"
                        f"{get_column_letter(t)}{r},"
                        f"{get_column_letter(a)}{r},-"
                        f"{get_column_letter(p)}{r})")
                ws.cell(row=r, column=n).value = expr
            r_gp = entity_rows[-1]
            expr_gp = (f"=SUM({get_column_letter(g)}{r_gp},-"
                      f"{get_column_letter(t)}{r_gp},"
                      f"{get_column_letter(a)}{r_gp})")
            ws.cell(row=r_gp, column=n).value = expr_gp
        
        # Clear progress
        progress_bar.empty()
        status_text.empty()
        
        # Save workbook to buffer
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        # Store in session so Complete flow can use automatically
        file_name = f"populated_template_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.session_state["shared_target_file_bytes"] = buf.getvalue()
        st.session_state["shared_target_file_name"] = file_name
 
        st.success("✅ Template populated successfully!")
        
        # Download button
        st.download_button(
            "📥 Download Populated Template",
            data=buf,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        st.info("💡 **Next Step:** Use this file as your 'Target Excel File' in the Complete Import File tab")

    
    st.markdown('</div>', unsafe_allow_html=True)

# Dispatch based on selected tab
with tab_incomplete:
    run_incomplete_flow()

with tab_complete:
    run_complete_flow()
