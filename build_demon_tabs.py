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
OUT = OUTPUT_DIR / (
    f"Davidson_group_ringing_data_withDemonTabs_{datetime.datetime.now():%Y-%m-%d-%H-%M-%S}.xlsx"
)

wb = load_workbook(SRC)

# Remove existing tabs if re-running
for name in ['Errors', 'DemOn Upload', 'Outstanding Comments']:
    if name in wb.sheetnames:
        del wb[name]

ws_data = wb['Ringing data']

# ── Read all data rows ────────────────────────────────────────────────────────
rows = list(ws_data.iter_rows(values_only=True))
header = rows[0]
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
}

NESTBOX_C5 = {'NW8', 'NW9', 'NW15'}

FEATHER_MAP = {
    'IP': 'IP',
    'S':  'FS',
    'M':  'FM',
    'L':  'FL',
}

SKIP_RINGTYPES = {'X', 'NOTRINGED', 'LOST', 'OVERLAPPED REMOVED'}

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
    has_pit = bool(pittag and str(pittag).strip())

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


def format_date(val):
    if isinstance(val, datetime.datetime):
        return val.strftime('%d/%m/%Y')
    if isinstance(val, datetime.date):
        return val.strftime('%d/%m/%Y')
    return str(val) if val else ''


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
FILL_BATCH_WARN    = PatternFill('solid', fgColor='FFE0E0')

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

err_row = 2
upload_count = 0

def add_error(data_idx, row, issue_type, details):
    global err_row
    vals = [
        ringing_data_row_num(data_idx),
        row[C_RINGNO],
        row[C_SPECIES],
        format_date(row[C_DATE]),
        issue_type,
        details,
    ]
    for c, v in enumerate(vals, 1):
        cell = ws_err.cell(row=err_row, column=c, value=v)
        style_cell(cell)
    err_row += 1


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
        format_date(row[C_DATE]),
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

    upload_count += 1

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

# Batch size warning — insert at top of errors if needed
if upload_count > 500:
    ws_err.insert_rows(2)
    warn_vals = [
        '—', '—', '—', '—',
        'Batch size warning',
        f'{upload_count} records marked Y — DemOn limit is 500 per file. Split into multiple uploads before exporting.'
    ]
    for c, v in enumerate(warn_vals, 1):
        cell = ws_err.cell(row=2, column=c, value=v)
        cell.font = Font(name='Arial', bold=True, color='8B0000', size=10)
        cell.fill = FILL_BATCH_WARN
        cell.border = BORDER

if err_row == 2:
    ws_err.cell(row=2, column=1, value='No errors found').font = Font(name='Arial', size=10, italic=True)

if com_row == 2:
    ws_com.cell(row=2, column=1, value='No comments found').font = Font(name='Arial', size=10, italic=True)

ws_err.freeze_panes = 'A2'
ws_com.freeze_panes = 'A2'


# ── Build DemOn Upload tab ────────────────────────────────────────────────────
ws_up = wb.create_sheet('DemOn Upload')

DEMON_COLS = [
    'SCHEME', 'RECORD_TYPE', 'RING_NO', 'SPECIES', 'AGE', 'SEX',
    'PROVISIONAL_SEX', 'SEXING_METHOD', 'CONDITION', 'METAL_MARK_INFO',
    'CAPTURE_METHOD', 'VISIT_DATE', 'DATE_MEASURED', 'CAPTURE_TIME',
    'TIME_MEASURED', 'LOCATION_CODE', 'HABITAT_1', 'WING_LENGTH', 'WEIGHT',
    'FINDING_CONDITION', 'FINDING_CIRCUMSTANCES',
    'WARNING_FC_SPECIAL_METHOD', 'WARNING_C_SPECIAL_METHOD',
    'RINGER_INITIALS', 'PROCESSOR_INITIALS',
    'PULLUS_STAGE', 'PULLI_ALIVE',
    'WARNING_AGE_CODE', 'WARNING_SEX',
    'SOURCE_ROW',
]

for c, h in enumerate(DEMON_COLS, 1):
    cell = ws_up.cell(row=1, column=c, value=h)
    fill = FILL_ORANGE_HEADER if h == 'SOURCE_ROW' else FILL_BLUE_HEADER
    style_header_cell(cell, fill)

ws_up.row_dimensions[1].height = 30

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
    'PULLUS_STAGE': 13, 'PULLI_ALIVE': 11,
    'WARNING_AGE_CODE': 18, 'WARNING_SEX': 13,
    'SOURCE_ROW': 14,
}
for c, h in enumerate(DEMON_COLS, 1):
    ws_up.column_dimensions[get_column_letter(c)].width = col_widths.get(h, 14)

up_row = 2

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
        sexmtd_out = sexmtd_raw

    fl_raw = str(row[C_FEATHER]).strip() if row[C_FEATHER] else ''
    pullus_stage = FEATHER_MAP.get(fl_raw, '')

    visit_date = format_date(row[C_DATE])
    capture_time = format_time(row[C_TIME])

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
        'RINGER_INITIALS':           row[C_INITIALS],
        'PROCESSOR_INITIALS':        row[C_INITIALS],
        'PULLUS_STAGE':              pullus_stage,
        'PULLI_ALIVE':               row[C_BROOD],
        'WARNING_AGE_CODE':          '',
        'WARNING_SEX':               '',
        'SOURCE_ROW':                ringing_data_row_num(i),
    }

    for c, col_name in enumerate(DEMON_COLS, 1):
        val = record.get(col_name, '')
        cell = ws_up.cell(row=up_row, column=c, value=val)
        style_cell(cell)

    up_row += 1

# Note row
note_row = up_row + 1
note = ws_up.cell(row=note_row, column=1,
    value='NOTE: SOURCE_ROW = row number in "Ringing data" sheet (use to trace records back). '
          'Delete SOURCE_ROW column before saving as CSV for DemOn upload. '
          'CAPTURE_METHOD assumes M (mist net) for all non-nestbox/non-pullus records — review Potter trap records manually.')
note.font = Font(name='Arial', size=9, italic=True, color='666666')
ws_up.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=12)

ws_up.freeze_panes = 'A2'

# ── Save ──────────────────────────────────────────────────────────────────────
wb.save(OUT)
print(f'Saved: {OUT}')
print(f'Upload rows: {up_row - 2}')
print(f'Error rows: {err_row - 2}')
print(f'Comment rows: {com_row - 2}')
