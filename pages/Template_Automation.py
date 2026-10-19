"""Clean source headers for the Template Automation workflow."""

import io
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from EmailBlasting_Template import (
    EXCEL_MIME, MANILA_TIMEZONE, MAX_UPLOAD_BYTES, ValidationError,
    automate_template, export_table, output_filename, read_source,
)


def clear_source():
    st.session_state.pop("template_automation_result", None)
    st.session_state["template_automation_generation"] += 1


def main():
    st.set_page_config(page_title="Template Automation", layout="wide")
    st.title("Template Automation")
    st.caption(
        "Remove spaces from headers, rename chname to FullName, Agent to Monthly, "
        "and Code to Terms. Missing Monthly and Terms columns are added blank."
    )
    st.session_state.setdefault("template_automation_generation", 0)
    generation = st.session_state["template_automation_generation"]
    uploaded = st.file_uploader(
        "Source file", type=["xlsx", "xls", "csv"],
        key=f"template_automation_upload_{generation}",
        help="Excel or UTF-8 CSV. Maximum 10 MB and 50,000 rows.",
    )
    if not uploaded:
        return
    st.button("Clear source file", on_click=clear_source)
    scope = f"{generation}_{uploaded.file_id}"
    try:
        content = uploaded.getvalue()
        if len(content) > MAX_UPLOAD_BYTES:
            raise ValidationError("Upload a source file that is 10 MB or smaller.")
        sheet_name = None
        with st.expander("File options"):
            if Path(uploaded.name).suffix.lower() in (".xlsx", ".xls"):
                try:
                    engine = "openpyxl" if uploaded.name.lower().endswith(".xlsx") else "xlrd"
                    with pd.ExcelFile(io.BytesIO(content), engine=engine) as workbook:
                        sheets = workbook.sheet_names
                    sheet_name = st.selectbox(
                        "Source sheet", sheets, key=f"template_automation_sheet_{scope}",
                    ) if len(sheets) > 1 else sheets[0]
                except Exception:
                    raise ValidationError(
                        "Could not open this workbook. Choose a valid, unprotected Excel file."
                    ) from None
            header_row = st.number_input(
                "Header row", min_value=1, max_value=1000, value=1, step=1,
                key=f"template_automation_header_{scope}",
            )
        signature = (uploaded.file_id, sheet_name, header_row)
        result = st.session_state.get("template_automation_result")
        if result is None or result["signature"] != signature:
            st.session_state.pop("template_automation_result", None)
            with st.spinner("Preparing file..."):
                source = read_source(content, uploaded.name, sheet_name, header_row)
                output = automate_template(source)
                filename = output_filename(
                    "Template", uploaded.name, datetime.now(MANILA_TIMEZONE),
                ).replace("EmailBlasting_(Template)", "TemplateAutomation", 1)
                result = {
                    "signature": signature, "output": output,
                    "content": export_table(output), "filename": filename,
                }
                st.session_state["template_automation_result"] = result
    except ValidationError as error:
        st.error(str(error))
        return

    output = result["output"]
    st.subheader("Preview")
    st.caption(f"{len(output):,} rows · Showing the first {min(len(output), 100):,}")
    st.dataframe(output.head(100), hide_index=True, width="stretch")
    st.download_button(
        "Download Excel", result["content"], result["filename"], EXCEL_MIME,
        type="primary", on_click="ignore",
    )
    st.caption(result["filename"])


if __name__ == "__main__":
    main()
