"""
build_individuals_table.py
Part of: fieldemon
Builds an 'Individuals' sheet in the ringing data workbook.
Each row = one unique ringed bird, with full history derived from Ringing data.

Usage:
    Save exactly one .xlsx source workbook in input/, then run:
        py build_individuals_table.py
    The output is saved to output/ with a timestamp.
"""

import datetime
import collections
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Locate input/output folders ───────────────────────────────────────────────
PROJECT_DIR = Path(__file__).resolve().parent
INPUT_DIR   = PROJECT_DIR / 'input'
OUTPUT_DIR  = PROJECT_DIR / 'output'

source_workbooks = sorted(
    path for path in INPUT_DIR.glob('*.xlsx') if not path.name.startswith('~$')
)

if len(source_workbooks) != 1:
    raise RuntimeError(
        f'Expected exactly one .xlsx workbook in {INPUT_DIR}, '
        f'found {len(source_workbooks)}.'
    )

SRC = source_workbooks[0]
OUTPUT_DIR.mkdir(exist_ok=True)
OUT = OUTPUT_DIR / (
    f"Davidson_group_ringing_data_withIndividuals_"
    f"{datetime.datetime.now():%Y-%m-%d-%H-%M-%S}.xlsx"
)

# ── Constants ─────────────────────────────────────────────────────────────────
SKIP_SPECIES  = {'WINTER 2026 PLACEHOLDER', 'LOST', 'OVERLAPPED REMOVED', None}
SKIP_RINGTYPES = {'NOTRINGED', 'LOST', 'OVERLAPPED REMOVED', 'X', None}

# ── Load workbook ─────────────────────────────────────────────────────────────
wb = load_workbook(SRC)

if 'Individuals' in wb.sheetnames:
    del wb['Individuals']

ws_data = wb['Ringing data']
rows    = list(ws_data.iter_rows(values_only=True))
header  = rows[0]
data    = rows[1:]

def col(name):
    try:
        return header.index(name)
    except ValueError:
        raise ValueError(
            f'Column "{name}" not found in Ringing data sheet. '
            f'Available columns: {list(header)}'
        )

C_LOC      = col('location')
C_DATE     = col('date')
C_SPECIES  = col('species')
C_RINGNO   = col('ringNo')
C_RINGTYPE = col('ringType')
C_AGE      = col('age')
C_SEX      = col('sex')
C_SEXMTD   = col('sexingMethod')
C_WING     = col('wing')
C_WEIGHT   = col('weight')
C_TIME     = col('time')
C_INITIALS = col('initials')
C_PITTAG   = col('pittagNo')
C_COMMENT  = col('comment')
C_NESTBOX  = col('nestbox')
C_FEATHER  = col('featherLength')
C_BROOD    = col('broodSize')
C_LETTERID = col('nestlingLetterID')
C_UPLOAD   = col('NEEDS UPLOADING TO DEMON?')

# ── Helper functions ──────────────────────────────────────────────────────────
def get_year(val):
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.year
    return None

def format_date(val):
    if isinstance(val, datetime.datetime):
        return val.strftime('%d/%m/%Y')
    if isinstance(val, datetime.date):
        return val.strftime('%d/%m/%Y')
    return str(val) if val else ''

def estimate_hatch_year(first_date, first_age):
    """
    Estimate hatch year from first seen date and age code.
    Returns (hatch_year, is_exact).
    Age codes: 1=pullus, 3=first year, 4=second year, 5/6=adult.
    """
    if not first_date or first_age is None:
        return None, False
    year = get_year(first_date)
    if year is None:
        return None, False
    try:
        age = int(str(first_age)[0])
    except (ValueError, TypeError):
        return None, False

    if age == 1:
        return year, True    # pullus — hatched this year, exact
    elif age == 3:
        return year, True    # first year — hatched this year, exact
    elif age == 4:
        return year - 1, True  # second year — hatched last year, exact
    elif age in (5, 6):
        return year - 1, False  # adult — hatched no later than last year
    else:
        return None, False

# ── Build individuals dictionary ──────────────────────────────────────────────
individuals   = {}   # ring_no → dict
nestling_link = {}   # (nestbox, letterID, year) → ring_no
ambiguous_links = set()  # (nestbox, letterID, year) with >1 ring

for i, row in enumerate(data):
    species  = row[C_SPECIES]
    ringno   = row[C_RINGNO]
    ringtype = row[C_RINGTYPE]
    date     = row[C_DATE]
    nestbox  = row[C_NESTBOX]
    letterid = row[C_LETTERID]
    year     = get_year(date)
    row_num  = i + 2  # 1-based index + header row

    # Skip placeholder and invalid rows
    if species in SKIP_SPECIES:
        continue
    if not ringno:
        continue
    if str(ringtype) in SKIP_RINGTYPES:
        continue

    if ringno not in individuals:
        hatch_year, hatch_exact = estimate_hatch_year(date, row[C_AGE])
        individuals[ringno] = {
            'ring_no':          ringno,
            'species':          species,
            'sex':              row[C_SEX] if row[C_SEXMTD] != 'G' else '',
            'first_seen':       date,
            'last_seen':        date,
            'first_ringtype':   ringtype,
            'first_age':        row[C_AGE],
            'hatch_year':       hatch_year,
            'hatch_exact':      hatch_exact,
            'first_location':   row[C_LOC],
            'first_nestbox':    nestbox,
            'nestling_letter':  letterid,
            'pit_tag':          row[C_PITTAG],
            'n_records':        1,
            'sex_conflict':     False,
            'species_conflict': False,
            'first_row':        row_num,
            'conflict_rows':    [],
            'ambiguous_link':   False,
        }
    else:
        ind = individuals[ringno]
        ind['n_records'] += 1

        # Update last seen date
        if date and (ind['last_seen'] is None or date > ind['last_seen']):
            ind['last_seen'] = date

        # Sex conflict (ignore G sexing method — that is provisional only)
        if row[C_SEXMTD] != 'G' and row[C_SEX] and ind['sex']:
            if row[C_SEX] != ind['sex']:
                ind['sex_conflict'] = True
                ind['conflict_rows'].append(
                    f'Sex conflict row {row_num} '
                    f'(recorded {row[C_SEX]}, expected {ind["sex"]})'
                )

        # Species conflict
        if species != ind['species']:
            ind['species_conflict'] = True
            ind['conflict_rows'].append(
                f'Species conflict row {row_num}: '
                f'{ind["species"]} vs {species}'
            )

        # Pick up pit tag if not already recorded
        if row[C_PITTAG] and not ind['pit_tag']:
            ind['pit_tag'] = row[C_PITTAG]

        # Pick up nestling letter if not already recorded
        if letterid and not ind['nestling_letter']:
            ind['nestling_letter'] = letterid

    # Build nestling link lookup for future pre-ringing rows
    if letterid and nestbox and year:
        key = (nestbox, letterid, year)
        if key in nestling_link:
            if nestling_link[key] != ringno:
                ambiguous_links.add(key)
                individuals[ringno]['ambiguous_link'] = True
                if nestling_link[key] in individuals:
                    individuals[nestling_link[key]]['ambiguous_link'] = True
        else:
            nestling_link[key] = ringno

# ── Sort by first seen date then ring number ──────────────────────────────────
sorted_individuals = sorted(
    individuals.values(),
    key=lambda x: (x['first_seen'] or datetime.datetime.min, x['ring_no'])
)

# ── Styles ────────────────────────────────────────────────────────────────────
FONT_HEADER   = Font(name='Arial', bold=True, color='FFFFFF', size=10)
FONT_NORMAL   = Font(name='Arial', size=10)
FONT_CONFLICT = Font(name='Arial', size=10, color='CC0000')
FONT_NOTE     = Font(name='Arial', size=9, italic=True, color='666666')

FILL_HEADER   = PatternFill('solid', fgColor='1A5276')
FILL_CONFLICT = PatternFill('solid', fgColor='FADBD8')
FILL_AMBIG    = PatternFill('solid', fgColor='FEF9E7')

thin   = Side(style='thin', color='CCCCCC')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def style_header(cell):
    cell.font      = FONT_HEADER
    cell.fill      = FILL_HEADER
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border    = BORDER

def style_cell(cell, conflict=False, ambig=False):
    cell.font      = FONT_CONFLICT if conflict else FONT_NORMAL
    cell.fill      = FILL_CONFLICT if conflict else (FILL_AMBIG if ambig else PatternFill())
    cell.alignment = Alignment(vertical='center')
    cell.border    = BORDER

# ── Build Individuals sheet ───────────────────────────────────────────────────
ws_ind = wb.create_sheet('Individuals')

IND_COLS = [
    'Ring No',
    'Species',
    'Sex',
    'First seen',
    'Last seen',
    'No. records',
    'Ring type\n(first)',
    'Age\n(first)',
    'Hatch year\n(est.)',
    'Hatch year\nexact?',
    'First location',
    'First nestbox',
    'Nestling\nletter ID',
    'PIT tag',
    'Sex\nconflict?',
    'Species\nconflict?',
    'Ambiguous\nnestling link?',
    'Conflict details',
    'First row in\nRinging data',
]

COL_WIDTHS = {
    'Ring No': 12,
    'Species': 10,
    'Sex': 6,
    'First seen': 13,
    'Last seen': 13,
    'No. records': 10,
    'Ring type\n(first)': 10,
    'Age\n(first)': 8,
    'Hatch year\n(est.)': 12,
    'Hatch year\nexact?': 12,
    'First location': 20,
    'First nestbox': 13,
    'Nestling\nletter ID': 12,
    'PIT tag': 16,
    'Sex\nconflict?': 10,
    'Species\nconflict?': 12,
    'Ambiguous\nnestling link?': 14,
    'Conflict details': 55,
    'First row in\nRinging data': 14,
}

for c, h in enumerate(IND_COLS, 1):
    cell = ws_ind.cell(row=1, column=c, value=h)
    style_header(cell)
    ws_ind.column_dimensions[get_column_letter(c)].width = COL_WIDTHS.get(h, 14)

ws_ind.row_dimensions[1].height = 35

for r_idx, ind in enumerate(sorted_individuals, 2):
    conflict = ind['sex_conflict'] or ind['species_conflict']
    ambig    = ind['ambiguous_link']

    values = [
        ind['ring_no'],
        ind['species'],
        ind['sex'] or '',
        format_date(ind['first_seen']),
        format_date(ind['last_seen']),
        ind['n_records'],
        ind['first_ringtype'],
        ind['first_age'],
        ind['hatch_year'],
        'Yes' if ind['hatch_exact'] else 'Est.',
        ind['first_location'],
        ind['first_nestbox'] or '',
        ind['nestling_letter'] or '',
        ind['pit_tag'] or '',
        'YES' if ind['sex_conflict'] else '',
        'YES' if ind['species_conflict'] else '',
        'YES' if ambig else '',
        '; '.join(ind['conflict_rows']) if ind['conflict_rows'] else '',
        ind['first_row'],
    ]

    for c_idx, val in enumerate(values, 1):
        cell = ws_ind.cell(row=r_idx, column=c_idx, value=val)
        style_cell(cell, conflict=conflict, ambig=ambig)

ws_ind.freeze_panes = 'A2'
ws_ind.auto_filter.ref = f'A1:{get_column_letter(len(IND_COLS))}1'

# ── Summary note ──────────────────────────────────────────────────────────────
note_row = len(sorted_individuals) + 3
note     = ws_ind.cell(
    row=note_row, column=1,
    value=(
        f'Generated {datetime.datetime.now().strftime("%d/%m/%Y %H:%M")} '
        f'from {len(data)} rows in {SRC.name}. '
        f'{len(sorted_individuals)} unique ringed birds. '
        f'Red rows = sex or species conflict. '
        f'Yellow rows = ambiguous nestling letter ID link '
        f'(same nestbox + letter + year on more than one ring number). '
        f'Regenerate by running build_individuals_table.py.'
    )
)
note.font = FONT_NOTE
ws_ind.merge_cells(
    start_row=note_row, start_column=1,
    end_row=note_row,   end_column=len(IND_COLS)
)

# ── Save ──────────────────────────────────────────────────────────────────────
wb.save(OUT)

# ── Console report ────────────────────────────────────────────────────────────
conflicts = [i for i in sorted_individuals if i['sex_conflict'] or i['species_conflict']]
ambiguous = [i for i in sorted_individuals if i['ambiguous_link']]

print(f'Source:           {SRC.name}')
print(f'Output:           {OUT.name}')
print(f'Unique birds:     {len(sorted_individuals)}')
print(f'Conflicts:        {len(conflicts)}')
print(f'Ambiguous links:  {len(ambiguous)}')

if conflicts:
    print('\nConflicts to review:')
    for ind in conflicts:
        print(f'  {ind["ring_no"]} ({ind["species"]}), '
              f'first row {ind["first_row"]}: '
              f'{"; ".join(ind["conflict_rows"])}')

if ambiguous:
    print('\nAmbiguous nestling links:')
    for ind in ambiguous:
        print(f'  {ind["ring_no"]} ({ind["species"]}), '
              f'nestbox {ind["first_nestbox"]}, '
              f'letter {ind["nestling_letter"]}')
