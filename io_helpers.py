import streamlit as st
import zipfile
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


def _safe_load_workbook(file_like):
    """Attempt to load an XLSX; on parse failure, show guidance and return None."""
    try:
        return load_workbook(file_like, data_only=False)
    except (ValueError, InvalidFileException, zipfile.BadZipFile) as e:
        st.error("⚠️  This template can’t be read as-is. Please open it in Excel, choose “Save As…”, and upload the saved copy.")
        st.caption(f"Details: {e}")
        return None 