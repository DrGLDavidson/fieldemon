import collections
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime
from pathlib import Path

# ── Load source workbook ──────────────────────────────────────────────────────
PROJECT_DIR = Path(__file__).resolve().parent
INPUT_DIR = PROJECT_DIR / 'input'
OUTPUT_DIR = PROJECT_DIR / 'output'
source_workbooks = sorted(
    path for path in INPUT_DIR.glob('*.xlsx') if not path.name.startswith('~$')
)

if len(source_workbooks) != 1:
    raise RuntimeError(
        f'Expected exactly one .xlsx workbook in {INPUT_DIR}, found {len(source_workbooks)}.'
    )

SRC = source_workbooks[0]
OUTPUT_DIR.mkdir(exist_ok=True)
RUN_TIMESTAMP = datetime.datetime.now()
OUT = OUTPUT_DIR / (
    f"Davidson_group_ringing_data_withDemonTabs_{RUN_TIMESTAMP:%Y-%m-%d-%H-%M-%S}.xlsx"
)
DATE_REPORTED = RUN_TIMESTAMP.date().isoformat()

wb = load_workbook(SRC)

# Remove existing tabs if re-running
for name in ['Errors', 'DemOn Upload', 'DemOn Upload 2', 'Outstanding Comments']:
    if name in wb.sheetnames:
        del wb[name]

ws_data = wb['Ringing data']

# ── Read all data rows ────────────────────────────────────────────────────────
rows = list(ws_data.iter_rows(values_only=True))
header = tuple(h.strip() if isinstance(h, str) else h for h in rows[0])
data = rows[1:]

def col(name):
    return header.index(name)

# Column indices
C_LOC       = col('location')
C_DATE      = col('date')
C_SPECIES   = col('species')
C_RINGNO    = col('ringNo')
C_RINGTYPE  = col('ringType')
C_AGE       = col('age')
C_SEX       = col('sex')
C_SEXMTD    = col('sexingMethod')
C_WING      = col('wing')
C_WEIGHT    = col('weight')
C_TIME      = col('time')
C_INITIALS  = col('initials')
C_PITTAG    = col('pittagNo')
C_PITTYPE   = col('pitType')
C_COMMENT   = col('comment')
C_NESTBOX   = col('nestbox')
C_FEATHER   = col('featherLength')
C_BROOD     = col('broodSize')
C_PULLUS_NOT_RINGED = col('pullusNotRinged')
C_LETTERID  = header.index('nestlingLetterID') if 'nestlingLetterID' in header else None
C_BREEDING = header.index('breedingCondition') if 'breedingCondition' in header else None
_header_lower = [h.lower() if isinstance(h, str) else h for h in header]
C_SPECIES_RINGSIZE_WARNING = (
    _header_lower.index('warning_species_ringsize_conflict')
    if 'warning_species_ringsize_conflict' in _header_lower else None
)
C_UPLOAD    = col('NEEDS UPLOADING TO DEMON?')
C_ERRORS    = col('ERRORS FOUND')

# ── Lookup tables ─────────────────────────────────────────────────────────────
LOCATION_MAP = {
    'rabbitenclosure':      ('RABBIT-UEA',       'A3'),
    'blackdaleplantation':  ('UEA-BD',            'A1'),
    'greenhouse':           ('UEA-GREEN',          'A3'),
    'newwoods':             ('UEA-COL-LN-WOOD',   'A3'),
    'universitycampus':     ('MAIN-UEA',           'F1'),
    'earlhamhall':          ('EARL-UEA',           'A1'),
    'universitywoods':      ('UEA-NESTBOXES',      'A1'),
    'ueawoods':             ('UEA-NESTBOXES',      'A1'),
    'charterwoods':         ('Carter Wood',        'A1'),
    'm_woods':              ('MAD-WOODS',          'A3'),
    'm_village':            ('MAD-VILLAGE',        'A6'),
}

SEXING_METHOD_MAP = {
    'BP': 'B',
    'C':  'C',
    'P':  'P',
}

NESTBOX_C5 = {'NW8', 'NW9', 'NW15'}

FEATHER_MAP = {
    'IP': 'IP',
    'S':  'FS',
    'M':  'FM',
    'L':  'FL',
}

SKIP_RINGTYPES = {'X', 'NOTRINGED', 'LOST', 'OVERLAPPED REMOVED'}
SKIP_SPECIES = {'WINTER 2026 PLACEHOLDER', 'LOST', 'OVERLAPPED REMOVED', None}

# ── Helper functions ──────────────────────────────────────────────────────────
def get_location_and_habitat(row):
    """Return (location_code, habitat_1, warning) for a row."""
    nestbox = row[C_NESTBOX]
    location = str(row[C_LOC]).strip() if row[C_LOC] else ''

    if nestbox:
        nestbox = str(nestbox).strip()
        loc_code = nestbox
        loc_key = location.lower()
        if loc_key in LOCATION_MAP:
            habitat = LOCATION_MAP[loc_key][1]
        else:
            habitat = ''
        if nestbox in NESTBOX_C5:
            habitat = 'C5'
        warning = ''
        if nestbox == 'UW2U':
            warning = 'Outdated nestbox code UW2U — check DemOn location code'
        return loc_code, habitat, warning
    else:
        loc_key = location.lower()
        if loc_key in LOCATION_MAP:
            demon_code, habitat = LOCATION_MAP[loc_key]
            return demon_code, habitat, ''
        elif location in ('LOST', ''):
            return location, '', f'Location "{location}" is invalid or missing — check DemOn location code'
        else:
            return location, '', f'Unknown location "{location}" — check DemOn location code'


def get_record_type_fields(row):
    """Return dict of DemOn record-type related fields, or None if row should be skipped."""
    ringtype = str(row[C_RINGTYPE]).strip() if row[C_RINGTYPE] else ''
    pittag = row[C_PITTAG]
    age = row[C_AGE]
    nestbox = row[C_NESTBOX]
    pittype = str(row[C_PITTYPE]).strip() if row[C_PITTYPE] else '' 
    has_pit = bool(pittag and str(pittag).strip()) and pittype == 'N'

    # Nestbox pulli
    if nestbox and age == 1:
        return {
            'RECORD_TYPE': 'N',
            'CONDITION': 'N',
            'METAL_MARK_INFO': 'N',
            'CAPTURE_METHOD': 'N',
            'FINDING_CONDITION': '',
            'FINDING_CIRCUMSTANCES': '',
            'WARNING_FC_SPECIAL_METHOD': '',
            'WARNING_C_SPECIAL_METHOD': '',
        }

    capture_method = 'N' if age in (1, 5, 6) else 'M'

    if ringtype == 'N':
        if has_pit:
            return {
                'RECORD_TYPE': 'M',
                'CONDITION': 'M',
                'METAL_MARK_INFO': 'N',
                'CAPTURE_METHOD': capture_method,
                'FINDING_CONDITION': '',
                'FINDING_CIRCUMSTANCES': '',
                'WARNING_FC_SPECIAL_METHOD': 'PTL',
                'WARNING_C_SPECIAL_METHOD': '',
            }
        else:
            return {
                'RECORD_TYPE': 'N',
                'CONDITION': 'N',
                'METAL_MARK_INFO': 'N',
                'CAPTURE_METHOD': capture_method,
                'FINDING_CONDITION': '',
                'FINDING_CIRCUMSTANCES': '',
                'WARNING_FC_SPECIAL_METHOD': '',
                'WARNING_C_SPECIAL_METHOD': '',
            }
    elif ringtype == 'R':
        if has_pit:
            return {
                'RECORD_TYPE': 'I',
                'CONDITION': 'M',
                'METAL_MARK_INFO': 'O',
                'CAPTURE_METHOD': capture_method,
                'FINDING_CONDITION': '8',
                'FINDING_CIRCUMSTANCES': '20',
                'WARNING_FC_SPECIAL_METHOD': '',
                'WARNING_C_SPECIAL_METHOD': 'PTL',
            }
        else:
            return {
                'RECORD_TYPE': 'S',
                'CONDITION': 'N',
                'METAL_MARK_INFO': 'O',
                'CAPTURE_METHOD': capture_method,
                'FINDING_CONDITION': '8',
                'FINDING_CIRCUMSTANCES': '20',
                'WARNING_FC_SPECIAL_METHOD': '',
                'WARNING_C_SPECIAL_METHOD': '',
            }
    else:
        return None


def format_date(val, row_num):
    if val is None or val == '':
        return ''
    if isinstance(val, datetime.datetime):
        return val.date().isoformat()
    if isinstance(val, datetime.date):
        return val.isoformat()
    if isinstance(val, str):
        date_text = val.strip()
        try:
            parsed_date = datetime.date.fromisoformat(date_text)
            if parsed_date.isoformat() == date_text:
                return parsed_date.isoformat()
        except ValueError:
            pass
    raise ValueError(
        f'Invalid date in Ringing data row {row_num}: {val!r}. '
        'Expected YYYY-MM-DD.'
    )


def format_time(val):
    if isinstance(val, datetime.time):
        return val.strftime('%H:%M')
    if isinstance(val, str) and val:
        return val
    return ''


def ringing_data_row_num(data_index):
    """Convert 0-based data index to row number in the Ringing data sheet (1-based + 1 for header)."""
    return data_index + 2


# ── Styles ────────────────────────────────────────────────────────────────────
FONT_HEADER = Font(name='Arial', bold=True, color='FFFFFF', size=10)
FONT_NORMAL = Font(name='Arial', size=10)
FONT_BOLD   = Font(name='Arial', bold=True, size=10)

FILL_RED_HEADER    = PatternFill('solid', fgColor='8B0000')
FILL_PURPLE_HEADER = PatternFill('solid', fgColor='5B2C6F')
FILL_BLUE_HEADER   = PatternFill('solid', fgColor='1F4E79')
FILL_ORANGE_HEADER = PatternFill('solid', fgColor='7D3C00')
FILL_WARN_ROW      = PatternFill('solid', fgColor='FFF2CC')

thin = Side(style='thin', color='CCCCCC')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def style_header_cell(cell, fill):
    cell.font = FONT_HEADER
    cell.fill = fill
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = BORDER

def style_cell(cell, bold=False):
    cell.font = FONT_BOLD if bold else FONT_NORMAL
    cell.alignment = Alignment(vertical='center')
    cell.border = BORDER


# ── Build Errors tab ──────────────────────────────────────────────────────────
ws_err = wb.create_sheet('Errors')

err_headers = [
    'Ringing data\nrow number',
    'Ring No',
    'Species',
    'Date',
    'Issue type',
    'Details',
    'date_reported',
    'date_corrected',
]

for c, h in enumerate(err_headers, 1):
    cell = ws_err.cell(row=1, column=c, value=h)
    style_header_cell(cell, FILL_RED_HEADER)

ws_err.row_dimensions[1].height = 35
ws_err.column_dimensions['A'].width = 16
ws_err.column_dimensions['B'].width = 14
ws_err.column_dimensions['C'].width = 12
ws_err.column_dimensions['D'].width = 14
ws_err.column_dimensions['E'].width = 30
ws_err.column_dimensions['F'].width = 65
ws_err.column_dimensions['G'].width = 14
ws_err.column_dimensions['H'].width = 14

error_note = ws_err.cell(
    row=2,
    column=1,
    value='NOTE: SOURCE_ROW = row number in "Ringing data" sheet (use to trace records back). '
          'Delete SOURCE_ROW column before saving as CSV for DemOn upload. '
          'CAPTURE_METHOD assumes M (mist net) for all non-nestbox/non-pullus records — review Potter trap records manually.',
)
error_note.font = Font(name='Arial', size=10, bold=True)
error_note.fill = FILL_WARN_ROW
error_note.alignment = Alignment(vertical='center', wrap_text=True)
error_note.border = BORDER
ws_err.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
ws_err.row_dimensions[2].height = 45

err_row = 3

def add_error(data_idx, row, issue_type, details, ring_no=None):
    global err_row
    vals = [
        ringing_data_row_num(data_idx),
        row[C_RINGNO] if ring_no is None else ring_no,
        row[C_SPECIES],
        format_date(row[C_DATE], ringing_data_row_num(data_idx)),
        issue_type,
        details,
        DATE_REPORTED,
        '',
    ]
    for c, v in enumerate(vals, 1):
        cell = ws_err.cell(row=err_row, column=c, value=v)
        style_cell(cell)
    err_row += 1


def find_ambiguous_nestling_links():
    if C_LETTERID is None:
        return {}

    records_by_key = collections.defaultdict(dict)
    for data_idx, row in enumerate(data):
        species = row[C_SPECIES]
        ringno = row[C_RINGNO]
        ringtype = row[C_RINGTYPE]
        nestbox = row[C_NESTBOX]
        letterid = str(row[C_LETTERID]).strip() if row[C_LETTERID] is not None else ''
        nestbox = str(nestbox).strip() if nestbox else ''
        age = row[C_AGE]

        if age != 1:
            continue
        if species in SKIP_SPECIES or not ringno or str(ringtype) in SKIP_RINGTYPES:
            continue
        if not nestbox or letterid.upper() in ('', 'N/A', 'NA'):
            continue

        row_num = ringing_data_row_num(data_idx)
        date_text = format_date(row[C_DATE], row_num)
        if not date_text:
            continue

        key = (nestbox, letterid, int(date_text[:4]))
        records_by_key[key].setdefault(ringno, {
            'data_idx': data_idx,
            'species': species,
        })

    return {
        key: records
        for key, records in records_by_key.items()
        if len(records) > 1
    }


def find_ring_conflicts():
    """Flag ring numbers on nestbox records that clash across nests, species or rows.

    Only rows with a nestbox are considered, so later retraps elsewhere are ignored.
    Returns {ringNo: {'rows': [data_idx, ...], 'problems': [str, ...]}}.
    """
    by_ring = collections.defaultdict(list)
    for data_idx, row in enumerate(data):
        ringno = row[C_RINGNO]
        species = row[C_SPECIES]
        nestbox = str(row[C_NESTBOX]).strip() if row[C_NESTBOX] else ''
        if not ringno or not str(ringno).strip() or not nestbox:
            continue
        if species in SKIP_SPECIES or str(row[C_RINGTYPE]).strip() in SKIP_RINGTYPES:
            continue
        date_text = format_date(row[C_DATE], ringing_data_row_num(data_idx))
        by_ring[str(ringno).strip()].append({
            'data_idx': data_idx,
            'nestbox': nestbox,
            'species': str(species).strip(),
            'year': int(date_text[:4]) if date_text else None,
            'age': row[C_AGE],
            'ringtype': str(row[C_RINGTYPE]).strip(),
        })

    def rows_text(records):
        return ', '.join(str(ringing_data_row_num(r['data_idx'])) for r in records)

    conflicts = {}
    for ringno, records in by_ring.items():
        if len(records) < 2:
            continue
        problems = []

        species_set = sorted({r['species'] for r in records})
        if len(species_set) > 1:
            problems.append(f'used for different species ({", ".join(species_set)})')

        by_year = collections.defaultdict(list)
        for r in records:
            by_year[r['year']].append(r)
        for year, year_records in sorted(by_year.items(), key=lambda kv: str(kv[0])):
            boxes = sorted({r['nestbox'] for r in year_records})
            if len(boxes) > 1:
                problems.append(f'used in different nestboxes in {year} ({", ".join(boxes)})')

        chick_boxes = sorted({r['nestbox'] for r in records if r['age'] == 1})
        if len(chick_boxes) > 1:
            problems.append(f'used for age 1 chicks in different nestboxes ({", ".join(chick_boxes)})')

        chicks = [r for r in records if r['age'] == 1]
        new_rings = [r for r in records if r['ringtype'] == 'N']
        if len(chicks) > 1:
            problems.append(f'appears {len(chicks)} times as an age 1 chick (rows {rows_text(chicks)})')
        elif len(new_rings) > 1:
            problems.append(f'is marked as newly ringed (ringType N) {len(new_rings)} times (rows {rows_text(new_rings)})')

        if problems:
            conflicts[ringno] = {
                'rows': [r['data_idx'] for r in records],
                'problems': problems,
            }
    return conflicts


def find_nestbox_consistency_issues():
    records_by_key = collections.defaultdict(list)
    for data_idx, row in enumerate(data):
        species = row[C_SPECIES]
        ringtype = row[C_RINGTYPE]
        nestbox = str(row[C_NESTBOX]).strip() if row[C_NESTBOX] else ''
        age = row[C_AGE]

        if age != 1:
            continue
        if species in SKIP_SPECIES or str(ringtype) in SKIP_RINGTYPES:
            continue
        if not nestbox or not species:
            continue
        if row[C_AGE] != 1:
            continue

        row_num = ringing_data_row_num(data_idx)
        date_text = format_date(row[C_DATE], row_num)
        if not date_text:
            continue

        key = (nestbox, int(date_text[:4]))
        records_by_key[key].append({
            'data_idx': data_idx,
            'species_code': str(species).strip(),
            'date': date_text,
        })

    issues = {}
    for key, records in records_by_key.items():
        species_codes = sorted({record['species_code'] for record in records})
        dates = sorted({record['date'] for record in records})
        if len(species_codes) == 1 and len(dates) == 1:
            continue

        details = []
        if len(species_codes) > 1:
            details.append(f'multiple species codes: {", ".join(species_codes)}')
        if len(dates) > 1:
            details.append(f'multiple dates: {", ".join(dates)}')
        issues[key] = {
            'records': records,
            'details': f'Nestbox {key[0]}, year {key[1]} has ' + '; '.join(details) + '.',
        }

    return issues


def find_nestbox_brood_size_issues():
    def normalize_count(value):
        try:
            numeric_value = float(value)
        except (TypeError, ValueError):
            return str(value).strip()
        return int(numeric_value) if numeric_value.is_integer() else numeric_value

    records_by_key = collections.defaultdict(list)
    for data_idx, row in enumerate(data):
        species = row[C_SPECIES]
        ringtype = row[C_RINGTYPE]
        nestbox = str(row[C_NESTBOX]).strip() if row[C_NESTBOX] else ''

        if species in SKIP_SPECIES or str(ringtype) in SKIP_RINGTYPES:
            continue
        if not nestbox or not species or row[C_AGE] != 1:
            continue

        row_num = ringing_data_row_num(data_idx)
        date_text = format_date(row[C_DATE], row_num)
        if not date_text:
            continue

        key = (nestbox, int(date_text[:4]))
        records_by_key[key].append({
            'data_idx': data_idx,
            'brood_size': row[C_BROOD],
            'pullus_not_ringed': row[C_PULLUS_NOT_RINGED],
        })

    issues = {}
    for key, records in records_by_key.items():
        brood_sizes = {
            normalize_count(record['brood_size'])
            for record in records
            if record['brood_size'] is not None
            and str(record['brood_size']).strip()
        }
        pullus_not_ringed_values = {
            normalize_count(record['pullus_not_ringed'])
            for record in records
            if record['pullus_not_ringed'] is not None
            and str(record['pullus_not_ringed']).strip()
        }
        has_missing_brood_size = any(
            record['brood_size'] is None
            or not str(record['brood_size']).strip()
            for record in records
        )

        if (
            len(brood_sizes) == 1
            and not has_missing_brood_size
            and len(pullus_not_ringed_values) <= 1
        ):
            brood_size = next(iter(brood_sizes))
            pullus_not_ringed = (
                next(iter(pullus_not_ringed_values))
                if pullus_not_ringed_values
                else 0
            )
            if not (
                isinstance(brood_size, (int, float))
                and isinstance(pullus_not_ringed, (int, float))
            ):
                brood_size_values = ', '.join(
                    sorted((str(value) for value in brood_sizes))
                ) or 'none'
                pullus_not_ringed_values_text = ', '.join(
                    sorted((str(value) for value in pullus_not_ringed_values))
                ) or 'none (treated as 0)'
                details = (
                    f'Nestbox {key[0]}, year {key[1]} has {len(records)} entries, '
                    f'but broodSize values are missing or inconsistent '
                    f'({brood_size_values}) or pullusNotRinged values are inconsistent '
                    f'({pullus_not_ringed_values_text}); number of entries cannot be verified.'
                )
            elif len(records) == brood_size - pullus_not_ringed:
                continue
            else:
                details = (
                    f'Nestbox {key[0]}, year {key[1]} has {len(records)} entries, '
                    f'but broodSize is {brood_size} and pullusNotRinged is '
                    f'{pullus_not_ringed}; expected '
                    f'{brood_size - pullus_not_ringed} ringed entries.'
                )
        else:
            brood_size_values = ', '.join(
                sorted((str(value) for value in brood_sizes))
            ) or 'none'
            pullus_not_ringed_values_text = ', '.join(
                sorted((str(value) for value in pullus_not_ringed_values))
            ) or 'none (treated as 0)'
            details = (
                f'Nestbox {key[0]}, year {key[1]} has {len(records)} entries, '
                f'but broodSize values are missing or inconsistent '
                f'({brood_size_values}) or pullusNotRinged values are inconsistent '
                f'({pullus_not_ringed_values_text}); number of entries cannot be verified.'
            )

        issues[key] = {
            'records': records,
            'details': details,
        }

    return issues


# ── Build Outstanding Comments tab ───────────────────────────────────────────
ws_com = wb.create_sheet('Outstanding Comments')

com_headers = [
    'Ringing data\nrow number',
    'Ring No',
    'Species',
    'Date',
    'Comment',
    'Action needed',
]

for c, h in enumerate(com_headers, 1):
    cell = ws_com.cell(row=1, column=c, value=h)
    style_header_cell(cell, FILL_PURPLE_HEADER)

ws_com.row_dimensions[1].height = 35
ws_com.column_dimensions['A'].width = 16
ws_com.column_dimensions['B'].width = 14
ws_com.column_dimensions['C'].width = 12
ws_com.column_dimensions['D'].width = 14
ws_com.column_dimensions['E'].width = 40
ws_com.column_dimensions['F'].width = 55

com_row = 2

def add_comment(data_idx, row, comment):
    global com_row
    vals = [
        ringing_data_row_num(data_idx),
        row[C_RINGNO],
        row[C_SPECIES],
        format_date(row[C_DATE], ringing_data_row_num(data_idx)),
        comment,
        'Review comment — check if data needs adding to DemOn (e.g. as OWN column or WARNING column). EN0 = egg number, for your records only.',
    ]
    for c, v in enumerate(vals, 1):
        cell = ws_com.cell(row=com_row, column=c, value=v)
        style_cell(cell)
    com_row += 1


# ── Process all rows ──────────────────────────────────────────────────────────
for i, row in enumerate(data):
    upload_flag = str(row[C_UPLOAD]).strip() if row[C_UPLOAD] else ''
    ringtype = str(row[C_RINGTYPE]).strip() if row[C_RINGTYPE] else ''
    errors_found = row[C_ERRORS]
    comment = row[C_COMMENT]

    # Flag skip types regardless of upload flag
    if ringtype in SKIP_RINGTYPES:
        add_error(i, row, 'Excluded from upload',
                  f'ringType "{ringtype}" — needs manual review before uploading to DemOn')

    if upload_flag != 'Y':
        continue

    # Known errors
    if errors_found:
        add_error(i, row, 'Known error recorded',
                  f'ERRORS FOUND: {errors_found}')

    # Missing mandatory fields
    for field_name, field_idx in [
        ('Ring No', C_RINGNO), ('Species', C_SPECIES),
        ('Age', C_AGE), ('Date', C_DATE), ('Time', C_TIME),
    ]:
        if not row[field_idx]:
            add_error(i, row, f'Missing {field_name}',
                      f'"{field_name}" is blank — required for DemOn upload')

    # Location
    _, _, loc_warning = get_location_and_habitat(row)
    if loc_warning:
        add_error(i, row, 'Location code issue', loc_warning)

    # Outdated nestbox
    if row[C_NESTBOX] and str(row[C_NESTBOX]).strip() == 'UW2U':
        add_error(i, row, 'Outdated nestbox code',
                  'UW2U may be outdated — verify current DemOn location code')

    # Cannot map record type
    rt_fields = get_record_type_fields(row)
    if rt_fields is None:
        add_error(i, row, 'Cannot determine record type',
                  f'ringType "{ringtype}" could not be mapped to a DemOn RECORD_TYPE')

    # Comments → Outstanding Comments tab (not Errors)
    if comment and str(comment).strip():
        add_comment(i, row, str(comment).strip())

# Check age-1 same-year nestbox entries for inconsistent species codes or dates.
nestbox_consistency_issues = find_nestbox_consistency_issues()
for issue in nestbox_consistency_issues.values():
    for record in issue['records']:
        data_idx = record['data_idx']
        add_error(
            data_idx,
            data[data_idx],
            'Nestbox species/date mismatch',
            issue['details'],
        )


# Check the number of ringed nestling entries against broodSize less unringed pulli.
nestbox_brood_size_issues = find_nestbox_brood_size_issues()
for issue in nestbox_brood_size_issues.values():
    record = issue['records'][0]
    data_idx = record['data_idx']
    ring_numbers = sorted({
        str(data[entry['data_idx']][C_RINGNO]).strip()
        for entry in issue['records']
        if data[entry['data_idx']][C_RINGNO]
        and str(data[entry['data_idx']][C_RINGNO]).strip()
    })
    ring_number_range = ''
    if ring_numbers:
        ring_number_range = ring_numbers[0]
        if len(ring_numbers) > 1:
            ring_number_range = f'{ring_numbers[0]} - {ring_numbers[-1]}'
    add_error(
        data_idx,
        data[data_idx],
        'Nestbox brood size mismatch',
        issue['details'],
        ring_no=ring_number_range,
    )

# Ambiguous age-1 nestling links need review before linking pre-ringing records.

ambiguous_nestling_links = find_ambiguous_nestling_links()
for (nestbox, letterid, year), records in ambiguous_nestling_links.items():
    ring_numbers = sorted(records, key=str)
    for ringno in ring_numbers:
        data_idx = records[ringno]['data_idx']
        others = [str(r) for r in ring_numbers if r != ringno]
        details = (
            f'Nestbox "{nestbox}" + nestlingLetterID "{letterid}" + year {year} '
            f'match {len(ring_numbers)} different ring numbers (rows '
            f'{", ".join(str(records[r]["data_idx"] + 2) for r in ring_numbers)}). '
            f'Each letter should identify one nestling per nestbox per year. '
            f'This ring also shares it with: {", ".join(others)}. '
            f'Check nestlingLetterID (or nestbox) on these rows.'
        )
        add_error(data_idx, data[data_idx], 'Ambiguous nestling link', details)

ring_conflicts = find_ring_conflicts()
for ringno, conflict in ring_conflicts.items():
    all_rows = ', '.join(str(ringing_data_row_num(i)) for i in conflict['rows'])
    details = (
        f'Ring {ringno} on nestbox records (rows {all_rows}): '
        + '; '.join(conflict['problems']) + '.'
    )
    for data_idx in conflict['rows']:
        add_error(data_idx, data[data_idx], 'Duplicate ring number', details)

if err_row == 3:
    ws_err.cell(row=3, column=1, value='No errors found').font = Font(name='Arial', size=10, italic=True)

if com_row == 2:
    ws_com.cell(row=2, column=1, value='No comments found').font = Font(name='Arial', size=10, italic=True)

ws_err.freeze_panes = 'A3'
ws_com.freeze_panes = 'A2'


# ── Build DemOn Upload tab ────────────────────────────────────────────────────
DEMON_COLS = [
    'SCHEME', 'RECORD_TYPE', 'RING_NO', 'SPECIES', 'AGE', 'SEX',
    'PROVISIONAL_SEX', 'SEXING_METHOD', 'CONDITION', 'METAL_MARK_INFO',
    'CAPTURE_METHOD', 'VISIT_DATE', 'DATE_MEASURED', 'CAPTURE_TIME',
    'TIME_MEASURED', 'LOCATION_CODE', 'HABITAT_1', 'WING_LENGTH', 'WEIGHT',
    'FINDING_CONDITION', 'FINDING_CIRCUMSTANCES',
    'WARNING_FC_SPECIAL_METHOD', 'WARNING_C_SPECIAL_METHOD',
    'RINGER_INITIALS', 'PROCESSOR_INITIALS',
    'PULLUS_STAGE', 'PULLI_ALIVE', 'PULLI_RINGED', 'BREEDING_CONDITION',
    'WARNING_AGE_CODE', 'WARNING_SEX',
]
if C_SPECIES_RINGSIZE_WARNING is not None:
    DEMON_COLS.append('WARNING_SPECIES_RINGSIZE_CONFLICT')
DEMON_COLS.append('SOURCE_ROW')

col_widths = {
    'SCHEME': 10, 'RECORD_TYPE': 13, 'RING_NO': 12, 'SPECIES': 10,
    'AGE': 6, 'SEX': 6, 'PROVISIONAL_SEX': 16, 'SEXING_METHOD': 14,
    'CONDITION': 11, 'METAL_MARK_INFO': 15, 'CAPTURE_METHOD': 15,
    'VISIT_DATE': 12, 'DATE_MEASURED': 14, 'CAPTURE_TIME': 13,
    'TIME_MEASURED': 13, 'LOCATION_CODE': 18, 'HABITAT_1': 11,
    'WING_LENGTH': 12, 'WEIGHT': 9,
    'FINDING_CONDITION': 18, 'FINDING_CIRCUMSTANCES': 22,
    'WARNING_FC_SPECIAL_METHOD': 24, 'WARNING_C_SPECIAL_METHOD': 24,
    'RINGER_INITIALS': 15, 'PROCESSOR_INITIALS': 18,
    'PULLUS_STAGE': 13, 'PULLI_ALIVE': 11, 'PULLI_RINGED': 12,         'BREEDING_CONDITION': 18,
    'WARNING_AGE_CODE': 18, 'WARNING_SEX': 13,
    'WARNING_SPECIES_RINGSIZE_CONFLICT': 34,
    'SOURCE_ROW': 14,
}

def create_upload_sheet(title):
    worksheet = wb.create_sheet(title)
    for c, h in enumerate(DEMON_COLS, 1):
        cell = worksheet.cell(row=1, column=c, value=h)
        fill = FILL_ORANGE_HEADER if h == 'SOURCE_ROW' else FILL_BLUE_HEADER
        style_header_cell(cell, fill)
        worksheet.column_dimensions[get_column_letter(c)].width = col_widths.get(h, 14)
    worksheet.row_dimensions[1].height = 30
    worksheet.freeze_panes = 'A2'
    return worksheet


ws_up = create_upload_sheet('DemOn Upload')
upload_sheets = [ws_up]
upload_sheet = ws_up
upload_sheet_row = 2
upload_record_count = 0

for i, row in enumerate(data):
    upload_flag = str(row[C_UPLOAD]).strip() if row[C_UPLOAD] else ''
    if upload_flag != 'Y':
        continue

    ringtype = str(row[C_RINGTYPE]).strip() if row[C_RINGTYPE] else ''
    if ringtype in SKIP_RINGTYPES:
        continue

    rt_fields = get_record_type_fields(row)
    if rt_fields is None:
        continue

    loc_code, habitat, _ = get_location_and_habitat(row)

    sex_raw = str(row[C_SEX]).strip() if row[C_SEX] else ''
    sexmtd_raw = str(row[C_SEXMTD]).strip() if row[C_SEXMTD] else ''
    if sexmtd_raw == 'G':
        sex_out = ''
        provisional_sex = sex_raw
        sexmtd_out = ''
    else:
        sex_out = sex_raw
        provisional_sex = ''
        sexmtd_out = SEXING_METHOD_MAP.get(sexmtd_raw, sexmtd_raw)

    fl_raw = str(row[C_FEATHER]).strip() if row[C_FEATHER] else ''
    pullus_stage = FEATHER_MAP.get(fl_raw, '')

    visit_date = format_date(row[C_DATE], ringing_data_row_num(i))
    capture_time = format_time(row[C_TIME])
    brood_size = row[C_BROOD]
    pullus_not_ringed = row[C_PULLUS_NOT_RINGED]
    pullus_ringed = (
        brood_size - (pullus_not_ringed or 0)
        if brood_size not in (None, '')
        else ''
    )

    record = {
        'SCHEME':                    'GBT',
        'RECORD_TYPE':               rt_fields['RECORD_TYPE'],
        'RING_NO':                   row[C_RINGNO],
        'SPECIES':                   row[C_SPECIES],
        'AGE':                       row[C_AGE],
        'SEX':                       sex_out,
        'PROVISIONAL_SEX':           provisional_sex,
        'SEXING_METHOD':             sexmtd_out,
        'CONDITION':                 rt_fields['CONDITION'],
        'METAL_MARK_INFO':           rt_fields['METAL_MARK_INFO'],
        'CAPTURE_METHOD':            rt_fields['CAPTURE_METHOD'],
        'VISIT_DATE':                visit_date,
        'DATE_MEASURED':             visit_date,
        'CAPTURE_TIME':              capture_time,
        'TIME_MEASURED':             capture_time,
        'LOCATION_CODE':             loc_code,
        'HABITAT_1':                 habitat,
        'WING_LENGTH':               row[C_WING],
        'WEIGHT':                    row[C_WEIGHT],
        'FINDING_CONDITION':         rt_fields['FINDING_CONDITION'],
        'FINDING_CIRCUMSTANCES':     rt_fields['FINDING_CIRCUMSTANCES'],
        'WARNING_FC_SPECIAL_METHOD': rt_fields['WARNING_FC_SPECIAL_METHOD'],
        'WARNING_C_SPECIAL_METHOD':  rt_fields['WARNING_C_SPECIAL_METHOD'],
        'RINGER_INITIALS':           row[C_INITIALS] if ringtype == 'N' else '',
        'PROCESSOR_INITIALS':        row[C_INITIALS],
        'PULLUS_STAGE':              pullus_stage,
        'PULLI_ALIVE':               brood_size,
        'PULLI_RINGED':              pullus_ringed,
        'BREEDING_CONDITION':            row[C_BREEDING] if C_BREEDING is not None else '',
        'WARNING_AGE_CODE':          '',
        'WARNING_SEX':               '',
        'WARNING_SPECIES_RINGSIZE_CONFLICT': (
            row[C_SPECIES_RINGSIZE_WARNING] if C_SPECIES_RINGSIZE_WARNING is not None else ''
        ),
        'SOURCE_ROW':                ringing_data_row_num(i),
    }

    if upload_record_count == 500:
        upload_sheet = create_upload_sheet('DemOn Upload 2')
        upload_sheets.append(upload_sheet)
        upload_sheet_row = 2

    for c, col_name in enumerate(DEMON_COLS, 1):
        val = record.get(col_name, '')
        cell = upload_sheet.cell(row=upload_sheet_row, column=c, value=val)
        style_cell(cell)

    upload_sheet_row += 1
    upload_record_count += 1

# ── Save ──────────────────────────────────────────────────────────────────────
wb.save(OUT)
print(f'Saved: {OUT}')
print(f'Upload rows: {upload_record_count}')
print(f'Error rows: {err_row - 3}')
print(f'Comment rows: {com_row - 2}')

if err_row > 3:
    print('\nErrors to review:')
    for error in ws_err.iter_rows(min_row=3, values_only=True):
        if error[4]:
            print(
                f'  row {error[0]}: {error[1]} ({error[2]}) - '
                f'{error[4]}: {error[5]}'
            )

if ambiguous_nestling_links:
    print('\nAmbiguous nestling links:')
    for (nestbox, letterid, year), records in ambiguous_nestling_links.items():
        for ringno, record in records.items():
            print(
                f'  {ringno} ({record["species"]}), nestbox {nestbox}, '
                f'letter {letterid}, year {year}'
            )

if nestbox_consistency_issues:
    print('\nNestbox species/date inconsistencies:')
    for (nestbox, year), issue in nestbox_consistency_issues.items():
        for record in issue['records']:
            print(
                f'  {data[record["data_idx"]][C_RINGNO]} '
                f'({record["species_code"]}), nestbox {nestbox}, '
                f'date {record["date"]}, year {year}'
            )
