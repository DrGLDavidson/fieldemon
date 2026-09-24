# Fieldemon

Convert Davidson Group ringing-data workbooks into DemOn upload tabs.

Fieldemon is a Python script that reads the `Ringing data` sheet and
automatically generates three tabs:

- `Errors` - flags records needing attention before upload
- `Outstanding Comments` - lists records with comments that may need adding to DemOn
- `DemOn Upload` - contains cleaned, correctly formatted data ready to export as CSV

The script never modifies the original `Ringing data` sheet.

## Repository Structure

- `build_demon_tabs.py` - the main script
- `README.md` - this quick-start guide
- `Instructions.txt` - detailed conversion and upload guidance
- `requirements.txt` - Python dependency list
- `input/` - local source workbooks (ignored by Git)
- `output/` - generated workbooks (ignored by Git)

## Setup

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
```

The script requires Python 3.x and `openpyxl`; there are no other runtime
dependencies.

## Data Safety

Source spreadsheets can contain personal or sensitive ringing data and should
never be committed to the repository. Keep workbooks and generated files in
the `input/` and `output/` folders, which are ignored by Git. The ignore rules
also cover common spreadsheet formats, Python cache files, virtual environments,
and local system files.

The repository's `.gitignore` should include:

```gitignore
*.xlsx
*.xls
*.csv
*.xlsm
__pycache__/
*.pyc
.venv/
.DS_Store
```

## Run

Save exactly one `.xlsx` source workbook in `input/`, then run:

```powershell
py build_demon_tabs.py
```

The converted workbook is saved to `output/` using the local run time, for
example `Davidson_group_ringing_data_withDemonTabs_2026-09-24-14-30-05.xlsx`.
The script stops with an error if `input/` is empty or contains more than one
`.xlsx` workbook.

## Correct and Rerun

If the generated workbook identifies errors that require manual correction,
edit the source workbook outside the project, then replace the workbook in
`input/` with the corrected version. Keep only that one `.xlsx` file in `input/`
and run the script again. Each run creates a new timestamped workbook in
`output/`.

## Export for DemOn Upload

1. Open the `DemOn Upload` tab in the generated workbook.
2. Delete the `SOURCE_ROW` column, which is the last column and is highlighted
	in brown. It is a helper column and is not a valid DemOn column.
3. Use **File > Save As > CSV (Comma delimited)**.
4. Log in to DemOn and go to **Enter Data > Ringing > Bulk Upload**.

See [`Instructions.txt`](Instructions.txt) for the full conversion rules and
upload guidance.