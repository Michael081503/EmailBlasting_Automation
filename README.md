# EmailBlastingAutomation

`EmailBlasting_Template.py` copies values from an uploaded Excel or CSV file into the 15-column template.

Choose **Template Automation** in the sidebar for the separate header-cleanup workflow. Upload an Excel or UTF-8 CSV file, select the sheet/header row if needed, preview the result, and download Excel. This page removes all whitespace from headers and renames `chname` → `FullName`, `Agent` → `Monthly`, and `Code` → `Terms`, ignoring capitalization. Missing `Monthly` and `Terms` columns are added with blank values; existing columns with those names are kept. Other columns and data values retain their order. Conflicting headers after cleanup are rejected so no columns are overwritten.

Output columns:

```text
email | primary_key | indexoneuid | cardnumber | chname | phone | OB | AGENT | CODE | doclass | Level | Product | EndoDate | ChristmasAmnesty | NewEndo
```

Upload the source, check the preview, and download the completed Excel file. File options are collapsed; column mapping opens when required columns are unmatched. ChristmasAmnesty, NewEndo, primary_key, indexoneuid, CODE, doclass, Level, and EndoDate can contain values or remain blank. Missing enrichment fields remain blank until a later source is supplied. `Phone 1` is the confirmed source for template `phone`, and takes precedence over `phone_1`.

Download names include the area, source filename, and Manila export date/time, including microseconds: `EmailBlasting_(NORTH)_source_20261013_203045_123456.xlsx`. The same prepared export keeps its timestamp; a new upload or changed mapping generates a new timestamp.

When the source has AREA, rows are split into `AMNESTY NEGATIVE (MINDANAO)`, `AMNESTY (MINDANAO)`, `AMNESTY NEGATIVE (R3)`, and `AMNESTY (R3)`. Empty or whitespace-only Christmas Amnesty values go to NEGATIVE; any value, including zero, goes to AMNESTY. REGION 3 is treated as R3. Only the selected source sheet is used. Other or blank areas are excluded with a visible count. Download groups separately or together as a ZIP. Grouping requires the original Christmas Amnesty source column; required template fields must be mapped or given confirmed fixed values.

The download contains one sheet with headers and values only. It has no added styling, custom widths, extra rows, or instructions. Identifiers and phone numbers are stored as text to preserve leading zeros. Source rows retain their order; duplicate rows are kept and completely empty rows are skipped.

Run locally:

```powershell
py -3.13 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run EmailBlasting_Template.py --server.address=127.0.0.1
```

For GitHub and Streamlit, include `EmailBlasting_Template.py`, `pages/Template_Automation.py`, `requirements.txt`, and `.streamlit/config.toml`. Set the main file path to `EmailBlasting_Template.py` and use Python 3.13. No credentials or separate template workbook are required. Keep source data out of GitHub and use the project's private sharing settings.

To deploy on [Streamlit Community Cloud](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy):

1. Click **Create app** at [share.streamlit.io](https://share.streamlit.io).
2. Select repository `Michael081503/EmailBlasting_Automation`, branch `main`, and main file path `EmailBlasting_Template.py`.
3. In **Advanced settings**, select Python **3.13**, save, and deploy. Leave Secrets empty; this app does not use them.

The pinned packages in `requirements.txt` are installed automatically. Use the local virtual environment above for verification as well; a different installed Streamlit version may not support the upload API used by the interface tests. The upload server enforces the 10 MB limit, and Excel export writes rows progressively to reduce memory use for larger files.

Verify:

```powershell
.\.venv\Scripts\python.exe -B -m unittest discover -s tests -v
```
