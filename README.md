# Fieldemon

Convert Davidson Group ringing-data workbooks into DemOn upload tabs and
maintain the Individuals table used by the master ringing workbook.

## Scripts

- `build_individuals_table.py` — rebuilds the Individuals tab; run after
  every ringing session
- `build_demon_tabs.py` — prepares data for DemOn upload; run a couple of
  times a year when ready to submit

Both scripts read from `input/` and write to `output/`. Neither modifies
the source workbook.

## Repository Structure

- `build_demon_tabs.py` - DemOn upload preparation script
- `build_individuals_table.py` - Individuals table script
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

The scripts require Python 3.x and `openpyxl`; there are no other runtime
dependencies.

## Data Safety

Source spreadsheets can contain personal or sensitive ringing data and should
never be committed to the repository. Keep workbooks and generated files in
the `input/` and `output/` folders, which are ignored by Git.

The master macro workbook (`.xlsm`) should also be kept outside the
repository entirely — store it somewhere secure such as a shared drive.

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

---

## Workflow A: After Every Ringing Session

Run this to keep the Individuals table in the master macro workbook up to
date. The Individuals tab is used by the VBA macros for live error checking.

1. Open the master macro workbook (`.xlsm`) in Excel.
2. Use **File > Save As** and save a copy as **Excel Workbook (`.xlsx`)**.
   This strips the macros and produces a clean data file.
3. Place that `.xlsx` file in the `input/` folder. Make sure `input/`
   contains only that one `.xlsx` file.
4. Run the script:
   ```powershell
   py build_individuals_table.py
   ```
5. The output is saved to `output/` with a timestamp, for example
   `Davidson_group_ringing_data_withIndividuals_2026-09-24-14-30-05.xlsx`.
6. Open the output workbook and the master macro workbook side by side.
7. Right-click the **Individuals** tab in the output workbook, choose
   **Move or Copy**, select the master macro workbook as the destination,
   tick **Create a copy**, and place it after the existing sheets.
8. If an Individuals tab already exists in the master macro workbook,
   delete the old one first, then repeat step 7.
9. Save the master macro workbook.

The Individuals tab is now current and the VBA macros will use it for
live error checking during the next ringing session.

---

## Workflow B: DemOn Upload (a few times a year)

Run this when you are ready to submit records to BTO DemOn.

1. Open the master macro workbook (`.xlsm`) in Excel.
2. Use **File > Save As** and save a copy as **Excel Workbook (`.xlsx`)**.
3. Place that `.xlsx` file in the `input/` folder. Make sure `input/`
   contains only that one `.xlsx` file.
4. Run the script:
   ```powershell
   py build_demon_tabs.py
   ```
5. The output is saved to `output/` with a timestamp, for example
   `Davidson_group_ringing_data_withDemonTabs_2026-09-24-14-30-05.xlsx`.
6. Open the output workbook and review the **Errors** and
   **Outstanding Comments** tabs. Resolve any issues in the master macro
   workbook, save a fresh `.xlsx` copy, and rerun the script if needed.
7. When ready to export, go to the **DemOn Upload** tab.
8. Delete the `SOURCE_ROW` column (last column, highlighted in brown).
   It is a helper column and is not a valid DemOn column.
9. Use **File > Save As > CSV (Comma delimited)**.
10. Log in to DemOn and go to **Enter Data > Ringing > Bulk Upload**.

See [`Instructions.txt`](Instructions.txt) for the full conversion rules,
location code table, record type logic, and upload guidance.

---

## Correct and Rerun

If either script identifies errors requiring manual correction, edit the
master macro workbook, save a fresh `.xlsx` copy to `input/`, and run the
script again. Each run creates a new timestamped file in `output/` so
previous outputs are never overwritten.
