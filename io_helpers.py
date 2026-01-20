import streamlit as st
import zipfile
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import io


def _safe_load_workbook(file_like):
    """Attempt to load an XLSX; on parse failure, show guidance and return None."""
    try:
        return load_workbook(file_like, data_only=False)
    except (ValueError, InvalidFileException, zipfile.BadZipFile) as e:
        st.error("⚠️  This template can’t be read as-is. Please open it in Excel, choose “Save As…”, and upload the saved copy.")
        st.caption(f"Details: {e}")
        return None 


def patch_xlsx_workbook_xml_from_template(template_xlsx_bytes: bytes, generated_xlsx_bytes: bytes) -> bytes:
    """
    Covercy appears to validate workbook-level metadata/extensions in `xl/workbook.xml`.
    `openpyxl` can drop unknown extensions on save, causing Covercy to reject the file.

    This function replaces `xl/workbook.xml` in the generated XLSX with the one from
    the original Covercy template, while keeping all worksheet data from the generated file.
    """
    part_name = "xl/workbook.xml"

    with zipfile.ZipFile(io.BytesIO(template_xlsx_bytes), "r") as tzip:
        template_part = tzip.read(part_name)

    with zipfile.ZipFile(io.BytesIO(generated_xlsx_bytes), "r") as gzip:
        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w") as outzip:
            for info in gzip.infolist():
                data = gzip.read(info.filename)
                if info.filename == part_name:
                    data = template_part
                outzip.writestr(info, data)

    return out_buf.getvalue()