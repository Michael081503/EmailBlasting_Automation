"""Fill the supplied email-blasting workbook layout from an uploaded source file."""

import csv
import io
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell


TEMPLATE_COLUMNS = [
    "email", "primary_key", "indexoneuid", "cardnumber", "chname", "phone",
    "OB", "AGENT", "CODE", "doclass", "Level", "Product", "EndoDate",
    "ChristmasAmnesty", "NewEndo",
]
OPTIONAL_COLUMNS = {
    "ChristmasAmnesty", "NewEndo", "primary_key", "indexoneuid", "CODE",
    "doclass", "Level", "EndoDate",
}
REQUIRED_COLUMNS = [column for column in TEMPLATE_COLUMNS if column not in OPTIONAL_COLUMNS]
EXPORT_VERSION = 4
ALIASES = {
    "email": ["emailaddress", "emailid", "debtoremail"],
    "primary_key": ["primaryid"],
    "phone": ["phonenumber", "phone1", "mobile", "mobilenumber", "telephone"],
    "indexoneuid": ["indexoneid"],
    "cardnumber": ["cardno", "cardnum"],
    "chname": ["cardholdername", "fullname", "name", "customername", "clientname"],
    "OB": ["outstandingbalance", "obmasterlist"],
    "doclass": ["doclassification"],
    "EndoDate": ["endorsementdate"],
    "CODE": ["areacode"],
    "AGENT": ["agentname"],
}
MAX_UPLOAD_BYTES = 10 * 1024 * 1024
MAX_SOURCE_ROWS = 50_000
CHOOSE_COLUMN = "Choose a source column"
LEAVE_BLANK = "Leave blank"
FIXED_VALUE = "Use a fixed value"
EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MANILA_TIMEZONE = timezone(timedelta(hours=8))
AMNESTY_CATEGORIES = ("AMNESTY NEGATIVE", "AMNESTY")
AMNESTY_AREAS = ("MINDANAO", "R3")


class ValidationError(ValueError):
    """A file or mapping error that is safe to display to the user."""


def normalize_header(value):
    return re.sub(r"[^a-z0-9]", "", str(value).casefold())


def excel_text(cell):
    """Keep text identifiers intact and honor simple Excel zero-padding formats."""
    value = cell.value
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        number_format = cell.number_format.split(";")[0]
        if isinstance(value, int) and re.fullmatch(r"0+", number_format):
            return str(value).zfill(len(number_format))
    return str(value)


def frame_from_rows(headers, rows):
    """Keep source values and row order; ignore completely empty rows/columns."""
    headers = [str(value).strip() if value is not None else "" for value in headers]
    active = []
    for index, header in enumerate(headers):
        has_data = any(index < len(row) and str(row[index]).strip() for row in rows)
        if header:
            active.append(index)
        elif has_data:
            raise ValidationError(f"Source column {index + 1} contains data but has no header. Add a column name or choose the correct header row.")
    selected_headers = [headers[index] for index in active]
    if not selected_headers:
        raise ValidationError("No column headers were found. Choose the row containing the headers.")
    if len(set(selected_headers)) != len(selected_headers):
        raise ValidationError("The source has duplicate column headers. Give each column a distinct name.")
    values = [
        [row[index] if index < len(row) else "" for index in active]
        for row in rows if any(str(value).strip() for value in row)
    ]
    if not values:
        raise ValidationError("No data rows were found below the selected header row.")
    if len(values) > MAX_SOURCE_ROWS:
        raise ValidationError(f"Use a source with {MAX_SOURCE_ROWS:,} data rows or fewer.")
    return pd.DataFrame(values, columns=selected_headers).fillna("").astype(str)


def read_source(content, filename, sheet_name=None, header_row=1):
    if len(content) > MAX_UPLOAD_BYTES:
        raise ValidationError("Upload a source file that is 10 MB or smaller.")
    if header_row < 1:
        raise ValidationError("The header row must be 1 or higher.")
    try:
        suffix = Path(filename).suffix.lower()
        if suffix == ".csv":
            reader = csv.reader(io.StringIO(content.decode("utf-8-sig")))
            for _ in range(header_row - 1):
                next(reader)
            headers = next(reader)
            rows = []
            for row in reader:
                if not row or not any(value.strip() for value in row):
                    continue
                if len(row) != len(headers):
                    raise ValidationError("CSV data rows must have the same number of columns as the header. Check commas and quoted values.")
                rows.append(row)
                if len(rows) > MAX_SOURCE_ROWS:
                    raise ValidationError(f"Use a source with {MAX_SOURCE_ROWS:,} data rows or fewer.")
        elif suffix == ".xlsx":
            workbook = load_workbook(io.BytesIO(content), data_only=True, read_only=True)
            try:
                sheet = workbook[sheet_name] if sheet_name else workbook.worksheets[0]
                if sheet.max_row is not None and sheet.max_row > MAX_SOURCE_ROWS + header_row:
                    raise ValidationError(f"Use a sheet with {MAX_SOURCE_ROWS:,} data rows or fewer. Remove unused formatted rows if necessary.")
                source_rows = sheet.iter_rows(min_row=header_row)
                headers = [cell.value for cell in next(source_rows)]
                rows = []
                for row in source_rows:
                    values = [excel_text(cell) for cell in row]
                    if not any(value.strip() for value in values):
                        continue
                    rows.append(values)
                    if len(rows) > MAX_SOURCE_ROWS:
                        raise ValidationError(f"Use a source with {MAX_SOURCE_ROWS:,} data rows or fewer.")
            finally:
                workbook.close()
        elif suffix == ".xls":
            raw = pd.read_excel(
                io.BytesIO(content), sheet_name=sheet_name or 0, header=None,
                dtype=str, keep_default_na=False, engine="xlrd",
            )
            headers = raw.iloc[header_row - 1].tolist()
            rows = raw.iloc[header_row:].values.tolist()
        else:
            raise ValidationError("Upload an XLSX, XLS, or CSV source file.")
        return frame_from_rows(headers, rows)
    except ValidationError:
        raise
    except Exception:
        raise ValidationError("Could not read the source. Check the file format, selected sheet, and header row. CSV files must use UTF-8.") from None


def suggest_mapping(source_columns):
    """Match the revised layout, including optional columns when present."""
    result = {}
    for target in TEMPLATE_COLUMNS:
        exact = [column for column in source_columns if normalize_header(column) == normalize_header(target)]
        aliases = [column for column in source_columns if normalize_header(column) in ALIASES.get(target, [])]
        # Phone 1 is the confirmed phone source, distinct from phone_1.
        preferred = [column for column in source_columns if re.sub(r"\s+", " ", str(column).strip()).casefold() == "phone 1"] if target == "phone" else []
        matches = preferred if preferred else exact if exact else aliases
        result[target] = matches[0] if len(matches) == 1 else None
    return result


def populate_template(source, mapping, fixed_values=None):
    """Copy mapped fields; the two optional columns may contain values or blanks."""
    fixed_values = fixed_values or {}
    missing = [field for field in REQUIRED_COLUMNS if not mapping.get(field) and not str(fixed_values.get(field, "")).strip()]
    if missing:
        raise ValidationError("Choose a source column or fixed value for: " + ", ".join(missing))
    output = pd.DataFrame("", index=range(len(source)), columns=TEMPLATE_COLUMNS)
    for field in TEMPLATE_COLUMNS:
        column = mapping.get(field)
        if column:
            if column not in source.columns:
                raise ValidationError(f"The selected source column for {field} was not found.")
            output[field] = source[column].fillna("").astype(str).to_numpy()
        else:
            output[field] = str(fixed_values.get(field, ""))
    return output


def canonical_area(value):
    normalized = normalize_header(value)
    if normalized == "mindanao":
        return "MINDANAO"
    if normalized in ("r3", "region3"):
        return "R3"
    return ""


def split_amnesty(source, output, area_column, amnesty_column):
    """Group rows by source AREA and whether the original Amnesty cell is blank."""
    if area_column not in source.columns:
        raise ValidationError("Choose the source AREA column.")
    if amnesty_column not in source.columns:
        raise ValidationError("Choose the Christmas Amnesty source column to split the files.")
    if len(source) != len(output):
        raise ValidationError("Source rows and template rows must have the same length.")
    areas = source[area_column].fillna("").astype(str).map(canonical_area).reset_index(drop=True)
    has_amnesty = source[amnesty_column].fillna("").astype(str).str.strip().ne("").reset_index(drop=True)
    output = output.reset_index(drop=True)
    groups = {}
    for area in AMNESTY_AREAS:
        for category in AMNESTY_CATEGORIES:
            mask = areas.eq(area) & (has_amnesty if category == "AMNESTY" else ~has_amnesty)
            if mask.any():
                groups[f"{category} ({area})"] = output.loc[mask].reset_index(drop=True)
    return groups, int(areas.eq("").sum())


def automate_template(source):
    """Clean headers and supply the Monthly and Terms fields without changing rows."""
    renames = {"chname": "FullName", "agent": "Monthly", "code": "Terms",
               "monthly": "Monthly", "terms": "Terms"}
    headers = []
    for column in source.columns:
        cleaned = re.sub(r"\s+", "", str(column))
        if not cleaned:
            raise ValidationError("Every source column must have a header.")
        headers.append(renames.get(cleaned.casefold(), cleaned))
    if len({header.casefold() for header in headers}) != len(headers):
        raise ValidationError(
            "Header cleanup would create duplicate columns. Use distinct headers "
            "and include only one of Agent/Monthly and Code/Terms."
        )
    output = source.copy()
    output.columns = headers
    for field in ("Monthly", "Terms"):
        if field not in output.columns:
            output[field] = ""
    return output


def export_table(output):
    """Export table headers and text values, without workbook styling."""
    workbook = Workbook(write_only=True)
    workbook.properties.creator = ""
    sheet = workbook.create_sheet("Sheet1")
    sheet.append(output.columns.tolist())
    for values in output.itertuples(index=False, name=None):
        cells = []
        for value in values:
            cell = WriteOnlyCell(sheet, value="" if pd.isna(value) else str(value))
            cell.data_type = "s"
            cells.append(cell)
        sheet.append(cells)
    buffer = io.BytesIO()
    workbook.save(buffer)
    workbook.close()
    return buffer.getvalue()


def export_workbook(output):
    """Export only the 15 template columns in their original order."""
    if output.columns.tolist() != TEMPLATE_COLUMNS:
        raise ValidationError("The output must use the 15 template columns in their original order.")
    return export_table(output)


def output_filename(area, source_name, generated_at, category=None):
    """Identify the area, source, and export time in each downloadable filename."""
    def clean_part(value, fallback, limit):
        cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value)
        cleaned = re.sub(r"\s+", "_", cleaned).strip(" ._")[:limit].rstrip(" ._")
        return cleaned or fallback

    area_part = clean_part(area, "Type_of_Area", 40)
    source_part = clean_part(Path(source_name.replace("\\", "/")).stem, "Source", 80)
    timestamp = generated_at.astimezone(MANILA_TIMEZONE).strftime("%Y%m%d_%H%M%S_%f")
    if category is not None:
        if category not in AMNESTY_CATEGORIES:
            raise ValidationError("Choose AMNESTY or AMNESTY NEGATIVE as the file category.")
        return f"{category} ({area_part})_{source_part}_{timestamp}.xlsx"
    return f"EmailBlasting_({area_part})_{source_part}_{timestamp}.xlsx"


def prepare_group_downloads(groups, source_name, generated_at):
    """Build separate plain Excel files and their ZIP archive in memory."""
    downloads = {}
    archive_buffer = io.BytesIO()
    with ZipFile(archive_buffer, "w", compression=ZIP_DEFLATED) as archive:
        for label, frame in groups.items():
            category, area = label.rsplit(" (", 1)
            filename = output_filename(area.rstrip(")"), source_name, generated_at, category)
            content = export_workbook(frame)
            downloads[label] = {"filename": filename, "content": content}
            archive.writestr(filename, content)
    return downloads, archive_buffer.getvalue()


def clear_source():
    for key in ("email_source", "email_source_signature", "email_result", "email_result_signature", "email_result_created_at", "email_group_downloads"):
        st.session_state.pop(key, None)
    st.session_state["email_upload_generation"] += 1


def main():
    st.set_page_config(page_title="Email Blasting Template", layout="wide")
    st.title("Email Blasting Template")
    st.caption("Copy source values into your Excel template.")
    st.session_state.setdefault("email_upload_generation", 0)
    generation = st.session_state["email_upload_generation"]
    uploaded = st.file_uploader(
        "Source file", type=["xlsx", "xls", "csv"], key=f"email_source_upload_{generation}",
        help="Excel or UTF-8 CSV. Maximum 10 MB and 50,000 rows.",
    )
    if not uploaded:
        return
    st.button("Clear source file", on_click=clear_source)
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
                    if len(sheets) > 1:
                        sheet_name = st.selectbox("Source sheet", sheets, key=f"email_sheet_{generation}_{uploaded.file_id}")
                    else:
                        sheet_name = sheets[0]
                except Exception:
                    raise ValidationError("Could not open this workbook. Choose a valid, unprotected Excel file.") from None
            header_row = st.number_input("Header row", min_value=1, max_value=1000, value=1, step=1, key=f"email_header_row_{generation}_{uploaded.file_id}")
        source_signature = (uploaded.file_id, sheet_name, header_row)
        if st.session_state.get("email_source_signature") != source_signature:
            for key in ("email_source", "email_source_signature", "email_result", "email_result_signature", "email_result_created_at", "email_group_downloads"):
                st.session_state.pop(key, None)
            source = read_source(content, uploaded.name, sheet_name, header_row)
            st.session_state["email_source"] = source
            st.session_state["email_source_signature"] = source_signature
        source = st.session_state["email_source"]
    except ValidationError as error:
        st.error(str(error))
        return

    suggested = suggest_mapping(source.columns)
    area_candidates = [column for column in source.columns if normalize_header(column) == "area"]
    grouped = bool(area_candidates)
    area_column = None
    mapping, fixed_values = {}, {}
    missing_required = (
        any(suggested[field] is None for field in REQUIRED_COLUMNS)
        or (grouped and suggested["ChristmasAmnesty"] is None)
        or len(area_candidates) > 1
    )
    with st.expander("Column mapping", expanded=missing_required):
        columns = source.columns.tolist()
        # Widget values are scoped to the file, selected sheet, and header row.
        mapping_scope = f"{generation}_{uploaded.file_id}_{sheet_name}_{header_row}"
        if grouped:
            area_options = [CHOOSE_COLUMN, *range(len(columns))]
            area_default = columns.index(area_candidates[0]) if len(area_candidates) == 1 else CHOOSE_COLUMN
            area_choice = st.selectbox(
                "AREA", area_options, index=area_options.index(area_default),
                format_func=lambda value: columns[value] if isinstance(value, int) else value,
                key=f"email_area_column_{mapping_scope}",
            )
            area_column = columns[area_choice] if isinstance(area_choice, int) else None
        mapping_columns = st.columns(2)
        for index, target in enumerate(TEMPLATE_COLUMNS):
            needs_source = grouped and target == "ChristmasAmnesty"
            blank_option = LEAVE_BLANK if target in OPTIONAL_COLUMNS and not needs_source else CHOOSE_COLUMN
            options = [blank_option, *range(len(columns))]
            if not needs_source:
                options.append(FIXED_VALUE)
            default = columns.index(suggested[target]) if suggested[target] is not None else blank_option
            with mapping_columns[index % 2]:
                choice = st.selectbox(
                    f"{target}{' (optional)' if target in OPTIONAL_COLUMNS and not needs_source else ''}", options, index=options.index(default),
                    format_func=lambda value: columns[value] if isinstance(value, int) else value,
                    key=f"email_map_{mapping_scope}_{target}",
                )
                mapping[target] = columns[choice] if isinstance(choice, int) else None
                if choice == FIXED_VALUE:
                    fixed_values[target] = st.text_input(f"{target} value for all rows", key=f"email_fixed_{mapping_scope}_{target}")
    try:
        output = populate_template(source, mapping, fixed_values)
        groups, excluded = split_amnesty(source, output, area_column, mapping["ChristmasAmnesty"]) if grouped else ({}, 0)
    except ValidationError as error:
        st.warning(str(error))
        return
    if grouped:
        if excluded:
            st.warning(f"{excluded:,} row(s) excluded: AREA is blank or is not MINDANAO or R3.")
        if not groups:
            st.warning("No rows have AREA MINDANAO or R3.")
            return
        st.subheader("Preview")
        preview_label = st.selectbox("File", list(groups), key=f"email_group_preview_{mapping_scope}")
        preview = groups[preview_label]
        st.caption(f"{len(preview):,} rows · Showing the first {min(len(preview), 100):,}")
        st.dataframe(preview.head(100), hide_index=True, width="stretch")
        result_signature = (EXPORT_VERSION, source_signature, tuple(mapping.items()), tuple(fixed_values.items()), area_column)
        if st.session_state.get("email_result_signature") != result_signature:
            with st.spinner("Preparing files..."):
                generated_at = datetime.now(MANILA_TIMEZONE)
                downloads, archive = prepare_group_downloads(groups, uploaded.name, generated_at)
                st.session_state["email_group_downloads"] = downloads
                st.session_state["email_result"] = archive
                st.session_state["email_result_created_at"] = generated_at
                st.session_state["email_result_signature"] = result_signature
        archive_name = output_filename("AMNESTY", uploaded.name, st.session_state["email_result_created_at"]).replace(".xlsx", ".zip")
        st.download_button("Download all files (ZIP)", st.session_state["email_result"], archive_name, "application/zip", type="primary", on_click="ignore")
        download_columns = st.columns(2)
        for index, (label, download) in enumerate(st.session_state["email_group_downloads"].items()):
            with download_columns[index % 2]:
                st.download_button(label, download["content"], download["filename"], EXCEL_MIME, key=f"email_download_{label}", on_click="ignore", width="stretch")
                st.caption(f"{len(groups[label]):,} rows")
        return
    st.subheader("Preview")
    st.caption(f"{len(output):,} rows · Showing the first {min(len(output), 100):,}")
    st.dataframe(output.head(100), hide_index=True, width="stretch")
    area = st.text_input("Type of area", value="Type of Area", key="email_area")
    result_signature = (EXPORT_VERSION, source_signature, tuple(mapping.items()), tuple(fixed_values.items()))
    if st.session_state.get("email_result_signature") != result_signature or "email_result_created_at" not in st.session_state:
        with st.spinner("Preparing file..."):
            st.session_state["email_result"] = export_workbook(output)
            st.session_state["email_result_signature"] = result_signature
            st.session_state["email_result_created_at"] = datetime.now(MANILA_TIMEZONE)
    download_name = output_filename(area, uploaded.name, st.session_state["email_result_created_at"])
    st.download_button(
        "Download Excel", st.session_state["email_result"],
        download_name, EXCEL_MIME, type="primary", on_click="ignore",
    )
    st.caption(download_name)


if __name__ == "__main__":
    from streamlit.runtime.scriptrunner import get_script_run_ctx

    if get_script_run_ctx(suppress_warning=True) is None:
        from streamlit.web import cli as streamlit_cli

        sys.argv = ["streamlit", "run", str(Path(__file__).resolve()), *sys.argv[1:]]
        raise SystemExit(streamlit_cli.main())
    main()
