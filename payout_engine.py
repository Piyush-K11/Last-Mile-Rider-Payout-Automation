"""
Rider Payout Converter — supports multiple vendors (see COMPANY_CONFIG) and
multiple stores within each. Converts a raw rider trip/attendance file into
that vendor's final payout workbook format.

NOTE: every vendor name, store name and amount in this project is fictional
sample data. See COMPANY_CONFIG for how to plug in your own.

Usage:
    python payout_engine.py <raw_file.xlsx> <output_file.xlsx> --company ALPHA --store "NORTH HUB"
    python payout_engine.py <raw_file.xlsx> <output_file.xlsx> --company BETA --store RIVERSIDE

--company defaults to ALPHA if omitted.

Override a store's numbers manually (new/unlisted store):
    python payout_engine.py <raw_file.xlsx> <output_file.xlsx> --company ALPHA \
        --store LAKESIDE --rate 40 --insurance 20 --tiers "7:45:640,6:45:420"

===============================================================================
ALPHA format — output has 3 tabs: PAYOUT / PAYOUT DATA / FINAL PAYOUT
===============================================================================
- PAYOUT tab = raw file's Sheet1 AS-IS (whatever its exact column names/order
  are, unchanged), with one new column appended: PAYOUT = Delivered Orders * rate.
- PAYOUT DATA tab = untouched copy of raw file's Sheet2 (values, formulas,
  formatting) — kept for reference only, not used in any calculation. Skipped
  if the raw file has no Sheet2.
- FINAL PAYOUT tab is computed independently from Sheet1 only:
    ATTENDANCE             = count of distinct days the rider appears on
                              (same rider, same day, multiple trips = 1 day)
    TOTAL DELIVERED ORDERS = sum of delivered orders
    INCENTIVE               = looked up from tiers (needs BOTH min login days
                              AND min delivered orders for that tier)
    PAYOUT                  = delivered orders * rate
    TOTAL PAYOUT             = PAYOUT + INCENTIVE
    Insurance Deduction     = flat value per store (0 if store doesn't deduct it)
    FINAL PAYOUT             = TOTAL PAYOUT - Insurance Deduction
    DCPO                     = ROUND(TOTAL PAYOUT / TOTAL DELIVERED ORDERS, 0)

===============================================================================
BETA format — output has 3 tabs: PAYOUT / INCENTIVE / INSURANCE DEDUCTION
===============================================================================
- PAYOUT tab = raw file's Sheet1 AS-IS, with PAYOUT column appended (same as ALPHA).
- INCENTIVE tab = pivot-style summary: RIDER NAME, ATTENDANCE, TOTAL DELIVERED
  ORDERS, INCENTIVE — computed from Sheet1 only. Riders with fewer login days
  than the smallest tier threshold (e.g. 1-5 days when tiers start at 6) are
  EXCLUDED from this tab entirely (not shown with 0, just not listed).
- INSURANCE DEDUCTION tab = RIDER NAME, INSURANCE DEDUCTION — every rider gets
  the flat per-store deduction, listed here regardless of attendance.
- No combined "Final Payout" / DCPO tab for BETA.

===============================================================================
Shared formatting (all sheets, all companies):
===============================================================================
- Header row: bold white text on green fill, centered, thin borders.
- Data rows: thin borders on every cell; FINAL PAYOUT / INCENTIVE / INSURANCE
  DEDUCTION data is center-aligned.
- TOTAL row: bold text, light-yellow fill, thin borders.
(The ALPHA "PAYOUT DATA" untouched Sheet2 copy is exempt — it is left exactly
as the raw file has it, no formatting changes, by design.)

Column detection is keyword-based (not exact-name based), so small naming
differences between weeks/stores/companies (e.g. "RIDER ID" vs
"RIDER ID / Emp. Code", sheet named "Sheet1" vs "Per Order Data") do not
break the script.
"""
import argparse
import copy
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill('solid', fgColor='1B5E20')  # dark green
TOTAL_FILL = PatternFill('solid', fgColor='FFF2CC')  # light yellow
HEADER_FONT = Font(name='Arial', size=11, bold=True, color='FFFFFF')
BODY_FONT = Font(name='Arial', size=10)
TOTAL_FONT = Font(name='Arial', size=10, bold=True)
CENTER = Alignment(horizontal='center', vertical='center')
HEADER_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=False)
THIN = Side(style='thin')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# ---------------------------------------------------------------------------
# Sample vendor rate cards, grouped by vendor.
#
# ALL VENDOR NAMES, STORE NAMES AND AMOUNTS BELOW ARE FICTIONAL SAMPLE DATA.
# Replace them with your own. Add new vendors/stores here — then
# `--company X --store Y` is all that's needed on the command line, and the
# web UI picks them up automatically.
#
# Each store entry describes ONE payout model. The keys present decide which
# model applies:
#   rate / tiers / insurance   per-order rate, attendance+volume incentive
#                              tiers, flat insurance deduction
#   mg                         minimum guarantee: flat base up to N orders,
#                              then a per-order rate beyond that
#   mg_per_trip                same MG formula, applied to every trip row
#   distance_mg                flat amount per trip + per-km distance bonus
#   flat_per_day               flat amount per rider-day, not order-driven
#   cancelled_rate             also pay for cancelled orders
#   dcpo_range                 expected (low, high) delivery-cost-per-order,
#                              used to flag outlier riders in the preview
# ---------------------------------------------------------------------------
COMPANY_CONFIG = {
    'ALPHA': {
        'NORTH HUB':   {'rate': 45, 'insurance': 20, 'tiers': [(7, 45, 640), (6, 45, 420)], 'dcpo_range': (47, 59)},
        'EAST HUB':    {'rate': 45, 'insurance': 20, 'tiers': [(7, 45, 640), (6, 45, 420)], 'dcpo_range': (47, 59)},
        'RIVERSIDE':   {'rate': 40, 'insurance': 20, 'tiers': [(7, 45, 640), (6, 45, 420)]},
        'WEST HUB':    {'rate': 40, 'insurance': 20, 'tiers': [(7, 45, 640), (6, 45, 420)], 'dcpo_range': (42, 54)},
        'LAKESIDE':    {'rate': 40, 'insurance': 20, 'tiers': [(7, 45, 640), (6, 45, 420)], 'dcpo_range': (42, 54)},
        'HILLTOP':     {'rate': 40, 'insurance': 20, 'tiers': [(7, 45, 640), (6, 45, 420)], 'dcpo_range': (42, 54)},
        'CENTRAL HUB': {'rate': 40, 'insurance': 20, 'tiers': [(7, 45, 640), (6, 45, 420)], 'dcpo_range': (42, 54)},
    },
    'BETA': {
        'RIVERSIDE': {'rate': 40, 'insurance': 20, 'tiers': [(7, 45, 640), (6, 45, 420)]},
        'NORTH HUB': {'rate': 45, 'insurance': 20, 'tiers': [(7, 45, 640), (6, 45, 420)],
                      'mg': {'base': 620, 'threshold': 12, 'rate': 45}},
    },
    'GAMMA': {
        # distance-based model: no per-order rate/incentive at all.
        # Total per trip = flat MG amount + (distance * per-km rate).
        'PINE STREET': {'insurance': 20, 'distance_mg': {'mg_flat': 600, 'per_km_rate': 2}},
        'MAPLE CIRCLE': {'insurance': 20, 'distance_mg': {'mg_flat': 550, 'per_km_rate': 0}},
        # Per-order rate + incentive tiers + insurance. The trip tab keeps the
        # raw sheet's own name ('Per Order Data') instead of being renamed to
        # 'PAYOUT', to match the vendor's own file conventions.
        'CITY SQUARE': {'rate': 40, 'insurance': 20, 'tiers': [(7, 45, 640), (6, 45, 420)],
                        'trip_tab_title': 'Per Order Data'},
        # The raw export mixes in a worker paid under a different vendor at
        # the same store — excluded via the attendance sheet's Vendor column.
        'OAKRIDGE': {'insurance': 20, 'distance_mg': {'mg_flat': 520, 'per_km_rate': 2},
                     'restrict_vendor': 'Gamma'},
    },
    'DELTA': {
        # no insurance deduction tab for this vendor at all, on any store.
        'NORTH HUB': {'rate': 45, 'mg': {'base': 620, 'threshold': 12, 'rate': 45}},
        'RIVERSIDE GARDEN': {'rate': 40, 'tiers': [(7, 45, 640), (6, 45, 420)]},
        'GREENFIELD PARK': {'rate': 50, 'tiers': [(6, 100, 880), (6, 85, 620), (6, 70, 360)]},
    },
    'EPSILON GROUP': {
        # Every trip (rider-day row) is paid its own minimum-guarantee formula
        # directly — no per-order rate, no incentive tiers, no aggregation
        # across the week first.
        'TECH PARK': {'insurance': 20, 'mg_per_trip': {'base': 680, 'threshold': 12, 'rate': 52}},
    },
    'ZETA': {
        # Payout = (delivered orders * rate) + (cancelled orders * cancelled_rate).
        # No insurance, no incentive tiers for this vendor.
        'HILL CITY': {'rate': 50, 'cancelled_rate': 20},
    },
    'THETA': {
        # Payout = (delivered orders * rate) + (cancelled orders * cancelled_rate),
        # same formula as ZETA, but THETA ALSO pays an attendance/delivery
        # incentive on top (see tiers). No insurance deduction tab.
        'HILL CITY PARK': {
            'rate': 50, 'cancelled_rate': 20,
            'tiers': [(7, 100, 880), (7, 85, 620), (6, 70, 360)],
        },
    },
    'IOTA': {
        'RIVERSIDE GARDEN': {'rate': 40, 'insurance': 20, 'tiers': [(7, 45, 640), (6, 45, 420)]},
    },
    'KAPPA': {
        # Flat per-day amount for every day the rider is present — payout
        # isn't order-driven at all, so no rate/tiers and no shipment check.
        # The raw file also carries an attendance sheet + bank-details sheet
        # that get copied through untouched; any other trailing sheet is
        # simply not carried over.
        'ORCHARD BAY': {'insurance': 12, 'flat_per_day': 660},
    },
    'LAMBDA': {
        # Flat amount per rider-day, no threshold/escalation at all — simpler
        # than the other MG stores (there's no "after X orders" tier).
        # No insurance deduction tab for this vendor.
        'NORTH HUB': {'flat_per_day': 720},
    },
    'OMEGA': {
        # Flat amount per trip/row, regardless of delivered orders — same
        # shape as LAMBDA, but this vendor DOES deduct insurance.
        'HARBOR POINT': {'insurance': 20, 'flat_per_day': 840},
    },
    'SIGMA': {
        # Same flat-MG + per-km-distance model as GAMMA: every trip pays a
        # flat amount plus a per-km bonus, no per-order rate, no tiers.
        'CENTRAL HUB': {'insurance': 20, 'distance_mg': {'mg_flat': 600, 'per_km_rate': 2}},
    },
    'TAU': {
        # Plain rate + insurance, no tiers, no FINAL PAYOUT rollup — see
        # build_per_order_insured for the no-trip-row filtering.
        'MILL TOWN': {'rate': 50, 'insurance': 20},
    },
}


def tiers_to_list(tiers):
    return [f'{days}d & {orders}+ orders → ₹{int(amount)}' for days, orders, amount in tiers]


def mg_summary(cfg):
    """Short display string for a store's minimum-guarantee scheme, or '-' if none."""
    if cfg.get('mg'):
        m = cfg['mg']
        return f"≤{m['threshold']} orders: ₹{m['base']} · +₹{m['rate']}/order after"
    if cfg.get('mg_per_trip'):
        m = cfg['mg_per_trip']
        return f"per trip: ≤{m['threshold']} orders: ₹{m['base']} · +₹{m['rate']}/order after"
    if cfg.get('distance_mg'):
        m = cfg['distance_mg']
        if m['per_km_rate'] == 0:
            return f"₹{m['mg_flat']} flat (no distance bonus)"
        return f"₹{m['mg_flat']} flat + ₹{m['per_km_rate']}/km"
    if cfg.get('flat_per_day'):
        return f"₹{cfg['flat_per_day']}/day flat (no order-based calc)"
    return '-'


def build_rate_card_rows():
    """Flattened list of every configured company/store, for the rate card UI."""
    rate_card = []
    for company, stores in COMPANY_CONFIG.items():
        for store, cfg in sorted(stores.items()):
            rate_card.append({
                'company': company,
                'store': store,
                'rate': cfg['rate'] if 'rate' in cfg else '-',
                'insurance': cfg['insurance'] if 'insurance' in cfg else '-',
                'tiers': tiers_to_list(cfg['tiers']) if cfg.get('tiers') else [],
                'mg': mg_summary(cfg),
            })
    rate_card.sort(key=lambda r: (r['company'], r['store']))
    return rate_card


def store_has_mg(company, store):
    """True if this store's payout uses any minimum-guarantee scheme (mg,
    mg_per_trip, or distance_mg). Shipment cross-checking isn't meaningful
    for these — payout isn't driven purely by reported delivered orders the
    way a plain rate x orders scheme is."""
    company_key = (company or '').strip().upper()
    store_key = (store or '').strip().upper()
    cfg = COMPANY_CONFIG.get(company_key, {}).get(store_key)
    if cfg is None:
        return False
    return bool(cfg.get('mg') or cfg.get('mg_per_trip') or cfg.get('distance_mg') or cfg.get('flat_per_day'))


def build_store_mg_flags():
    """Map of 'COMPANY|STORE' -> bool(has_mg), for every configured store.
    Used by the frontend to hide the shipment-check upload where it doesn't
    apply."""
    flags = {}
    for company, stores in COMPANY_CONFIG.items():
        for store, cfg in stores.items():
            flags[f'{company}|{store}'] = bool(
                cfg.get('mg') or cfg.get('mg_per_trip') or cfg.get('distance_mg') or cfg.get('flat_per_day')
            )
    return flags


def find_col(columns, *keywords):
    """Return the first column whose uppercased name contains all keywords."""
    for col in columns:
        up = str(col).upper()
        if all(kw in up for kw in keywords):
            return col
    return None


class MissingColumnError(ValueError):
    """Raised when a required column can't be found in a raw sheet. Carries
    structured detail (not just a message string) so the UI can show exactly
    which columns were missing and what was actually in the file."""
    def __init__(self, sheet_id, missing, found):
        self.sheet_id = sheet_id
        self.missing = missing
        self.found = found
        message = (f"Could not find required column(s) in raw sheet '{sheet_id}': {missing}. "
                   f"Actual columns found: {found}")
        super().__init__(message)


ABSENCE_MARKERS = {'A', 'ABSENT', 'LEAVE', 'NA', 'N/A'}


def filter_absent_rows(df):
    """Drop rows where IN TIME or OUT TIME is a placeholder like 'A' (Absent)
    rather than a real time value. Applies to every raw sheet, every company —
    an absent row isn't a trip and shouldn't count toward attendance or payout.
    Doesn't touch rows with 0 or blank times, since those represent real
    (if incomplete) trip records, not marked absences."""
    in_col = find_col(df.columns, 'IN', 'TIME')
    out_col = find_col(df.columns, 'OUT', 'TIME')
    cols = [c for c in (in_col, out_col) if c is not None]
    if not cols:
        return df
    mask = pd.Series(True, index=df.index)
    for col in cols:
        is_absent = df[col].astype(str).str.strip().str.upper().isin(ABSENCE_MARKERS)
        mask &= ~is_absent
    return df[mask].reset_index(drop=True)


def normalize_date_col(df, date_col):
    """Raw files sometimes mix real date/datetime cells with plain text dates
    in the same DATE column (e.g. some rows '13/07/2026' as text, others as
    an actual Excel date) — pandas then reads the column as mixed str/
    Timestamp objects, which makes nunique() (attendance/login-day counts)
    overcount the same calendar day as two different values. Coercing the
    whole column to real datetimes up front fixes attendance counts and also
    makes the dd-mm-yyyy display formatting apply consistently."""
    df[date_col] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
    return df


def load_sheet(raw_path, sheet_id=0):
    """sheet_id can be a position (0 = first sheet) or a sheet name."""
    df = pd.read_excel(raw_path, sheet_name=sheet_id)
    df.columns = [str(c).strip() for c in df.columns]
    df = filter_absent_rows(df)

    date_col = find_col(df.columns, 'DATE')
    rider_col = find_col(df.columns, 'RIDER', 'NAME')
    delivered_col = find_col(df.columns, 'DELIVERED', 'ORDER')
    store_col = find_col(df.columns, 'STORE')

    missing = [n for n, c in [('DATE', date_col), ('RIDER NAME', rider_col),
                               ('TOTAL DELIVERED ORDERS', delivered_col)] if c is None]
    if missing:
        raise MissingColumnError(sheet_id, missing, list(df.columns))
    df = normalize_date_col(df, date_col)

    return df, date_col, rider_col, delivered_col, store_col


def load_sheet1(raw_path):
    return load_sheet(raw_path, 0)


SHIPMENT_CHECK_FILL_OK = PatternFill('solid', fgColor='2E7D32')   # light-ish green
SHIPMENT_CHECK_FILL_BAD = PatternFill('solid', fgColor='C62828')  # red
SHIPMENT_CHECK_FONT = Font(name='Arial', size=10, bold=False, color='FFFFFF')


def _normalize_trip_id(val):
    """Make trip IDs from different sources comparable: '16905', 16905,
    16905.0, ' 16905 ', and '16,905' should all normalize to the same key."""
    if val is None:
        return None
    s = str(val).strip().replace(',', '')
    if s == '' or s.lower() == 'nan':
        return None
    try:
        return str(int(float(s)))
    except (ValueError, TypeError):
        return s


def load_shipment_lookup(shipment_path):
    """Read a shipment/tracking file and build {trip_id: delivered_count}.
    Looks for a sheet with a header row containing both a 'Trip ID' column
    and a 'Total' (or similar count) column — this matches the ready-made
    pivot format riders' shipment files commonly use. Scans every sheet,
    and within each sheet scans the first several rows for the header,
    since these pivot exports often have a few blank/label rows above it.

    Uses read_only streaming mode: shipment files are often tens of MB with
    a large "detailed" sheet we never actually need, and a normal
    load_workbook() call eagerly parses every sheet up front even if we only
    read a few rows from one of them. read_only mode + iter_rows avoids that
    entirely and is dramatically faster on large files."""
    wb = load_workbook(shipment_path, data_only=True, read_only=True)
    try:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            row_iter = ws.iter_rows(min_row=1, max_row=15, values_only=True)
            for r, row_values in enumerate(row_iter, start=1):
                headers = {}
                for c, v in enumerate(row_values, start=1):
                    if v:
                        headers[str(v).strip().upper()] = c
                trip_col = next((c for h, c in headers.items() if 'TRIP' in h and 'ID' in h), None)
                total_col = next((c for h, c in headers.items()
                                   if h in ('TOTAL', 'COUNT') or ('TOTAL' in h and 'QUANTITY' not in h)), None)
                if not (trip_col and total_col):
                    continue

                lookup = {}
                blank_streak = 0
                for row_values in ws.iter_rows(min_row=r + 1, values_only=True):
                    trip_val = row_values[trip_col - 1] if len(row_values) >= trip_col else None
                    if trip_val is None:
                        blank_streak += 1
                        if blank_streak >= 3:
                            break
                        continue
                    blank_streak = 0
                    key = _normalize_trip_id(trip_val)
                    total_val = row_values[total_col - 1] if len(row_values) >= total_col else None
                    if key is not None and total_val is not None:
                        try:
                            lookup[key] = lookup.get(key, 0) + float(total_val)
                        except (ValueError, TypeError):
                            pass
                if lookup:
                    return lookup
    finally:
        wb.close()
    raise ValueError("Could not find a 'Trip ID' + count column in the shipment file "
                      "(checked every sheet's first 15 rows for a matching header).")


def shipment_check_diff(trip_id_raw, delivered_value, lookup):
    """Returns (value, status) where status is one of:
    - None: not applicable (no shipment data provided, or trip id blank)
    - 'ok': delivered matches or is LESS than shipment's confirmed count
      (no overpayment risk) -> shown in green, value = diff (delivered - shipment)
    - 'mismatch': delivered EXCEEDS what shipment confirms for a trip ID that
      WAS found in the shipment data -> shown in red, value = diff
    - 'not_found': trip ID (or a part of a combined 'A/B' trip ID) does not
      appear in the shipment data at all -> shown in red, value = 'NOT FOUND'
      (deliberately NOT a number, so it's never confused with a genuine
      confirmed discrepancy — if you're seeing 'NOT FOUND' everywhere, the
      trip IDs in your raw file aren't matching your shipment file at all,
      which usually means a format mismatch, not a real payout problem)

    Handles combined trip IDs like '28538/28388' by summing shipment counts
    across every part; if ANY part isn't found, the whole trip is 'not_found'."""
    if lookup is None or trip_id_raw is None:
        return None, None
    raw_str = str(trip_id_raw).strip()
    if raw_str == '' or raw_str.lower() == 'nan':
        return None, None
    parts = [p.strip() for p in raw_str.split('/') if p.strip()]
    if not parts:
        return None, None

    shipment_total = 0
    any_missing = False
    for p in parts:
        key = _normalize_trip_id(p)
        if key is None or key not in lookup:
            any_missing = True
        else:
            shipment_total += lookup[key]

    if any_missing:
        return 'NOT FOUND', 'not_found'

    try:
        delivered = float(delivered_value) if delivered_value is not None else 0
    except (ValueError, TypeError):
        delivered = 0

    diff = delivered - shipment_total
    # round to int when it's a whole number, for a clean cell value
    if diff == int(diff):
        diff = int(diff)
    return diff, ('ok' if diff <= 0 else 'mismatch')


def apply_shipment_check_cell(cell, trip_val, delivered_val, shipment_lookup):
    """Writes the SHIPMENT CHECK cell value + green/red styling, used by every
    build_* function so the behavior (including the not-found case) stays
    consistent everywhere instead of being duplicated four times."""
    value, status = shipment_check_diff(trip_val, delivered_val, shipment_lookup)
    if status is None:
        cell.value = 'NO TRIP ID'
        cell.font = BODY_FONT
    elif status == 'not_found':
        cell.value = 'NOT FOUND'
        cell.fill = SHIPMENT_CHECK_FILL_BAD
        cell.font = SHIPMENT_CHECK_FONT
    elif status == 'ok':
        cell.value = value
        cell.fill = SHIPMENT_CHECK_FILL_OK
        cell.font = SHIPMENT_CHECK_FONT
    else:  # 'mismatch'
        cell.value = value
        cell.fill = SHIPMENT_CHECK_FILL_BAD
        cell.font = SHIPMENT_CHECK_FONT


def find_mg_sheet_name(raw_path):
    """If the raw file has a sheet whose name contains 'MG', return its name;
    otherwise None (store has no minimum-guarantee scheme / single-sheet raw file)."""
    wb = load_workbook(raw_path, data_only=False)
    for name in wb.sheetnames:
        if 'MG' in name.upper():
            return name
    return None


def find_non_mg_sheet_name(raw_path, mg_sheet_name):
    wb = load_workbook(raw_path, data_only=False)
    for name in wb.sheetnames:
        if name != mg_sheet_name:
            return name
    return wb.sheetnames[0]


def copy_sheet2_untouched(raw_path, target_wb, target_name='PAYOUT DATA'):
    src_wb = load_workbook(raw_path, data_only=False)
    if len(src_wb.sheetnames) < 2:
        return False  # no Sheet2 present, skip silently
    src_ws = src_wb[src_wb.sheetnames[1]]
    dst_ws = target_wb.create_sheet(target_name)
    for row in src_ws.iter_rows():
        for cell in row:
            new_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy.copy(cell.font)
                new_cell.fill = copy.copy(cell.fill)
                new_cell.border = copy.copy(cell.border)
                new_cell.alignment = copy.copy(cell.alignment)
                new_cell.number_format = cell.number_format
    for col_letter, dim in src_ws.column_dimensions.items():
        if dim.width:
            dst_ws.column_dimensions[col_letter].width = dim.width
    return True


def find_sheet_by_keyword(raw_path, *keywords):
    """Return the name of the first sheet whose name contains all keywords
    (case-insensitive), or None if no sheet matches."""
    wb = load_workbook(raw_path, data_only=False)
    for name in wb.sheetnames:
        up = name.upper()
        if all(kw in up for kw in keywords):
            return name
    return None


def copy_sheet_untouched(raw_path, sheet_name, target_wb, target_name):
    """Copy one named sheet from the raw file into target_wb, values +
    formatting preserved exactly, under a new tab name. Returns False (no-op)
    if sheet_name isn't present. Generalizes copy_sheet2_untouched to an
    arbitrary sheet name instead of always 'the 2nd sheet'."""
    src_wb = load_workbook(raw_path, data_only=False)
    if sheet_name not in src_wb.sheetnames:
        return False
    src_ws = src_wb[sheet_name]
    dst_ws = target_wb.create_sheet(target_name)
    for row in src_ws.iter_rows():
        for cell in row:
            new_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy.copy(cell.font)
                new_cell.fill = copy.copy(cell.fill)
                new_cell.border = copy.copy(cell.border)
                new_cell.alignment = copy.copy(cell.alignment)
                new_cell.number_format = cell.number_format
    for col_letter, dim in src_ws.column_dimensions.items():
        if dim.width:
            dst_ws.column_dimensions[col_letter].width = dim.width
    return True


def preview_distance_mg(raw_path, insurance, mg_flat, per_km_rate,
                         distance_col_override=None, restrict_vendor=None):
    """Shared preview-stats computation for any store using the flat MG +
    per-km distance model (Gamma, SIGMA, etc.) — every row/trip pays
    mg_flat plus distance*per_km_rate, no per-order rate or incentive tiers."""
    df = pd.read_excel(raw_path, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    df = filter_absent_rows(df)
    date_col = find_col(df.columns, 'DATE')
    rider_col = find_col(df.columns, 'RIDER', 'NAME')
    if distance_col_override:
        distance_col = next((c for c in df.columns
                              if str(c).strip().upper() == distance_col_override.strip().upper()), None)
    else:
        distance_col = find_col(df.columns, 'DISTANCE')
    delivered_col = find_col(df.columns, 'DELIVERED', 'ORDER')

    missing = [n for n, c in [('DATE', date_col), ('RIDER NAME', rider_col),
                               ('DISTANCE', distance_col)] if c is None]
    if missing:
        raise MissingColumnError(0, missing, list(df.columns))
    df = normalize_date_col(df, date_col)
    if restrict_vendor:
        df = restrict_riders_by_attendance_vendor(raw_path, df, rider_col, restrict_vendor)

    grp = df.groupby(rider_col)
    days_by_rider = grp[date_col].nunique()
    delivered_by_rider = grp[delivered_col].sum() if delivered_col else None
    riders = sorted(days_by_rider.index.tolist())

    rider_rows = []
    total_delivered = total_payout = total_final = 0
    for name in riders:
        rows = df[df[rider_col] == name]
        trips = len(rows)
        payout = float(rows[distance_col].fillna(0).astype(float).sum() * per_km_rate + mg_flat * trips)
        final = float(payout - insurance)
        d = int(delivered_by_rider[name]) if delivered_by_rider is not None else 0
        rider_rows.append({
            'name': str(name), 'days': int(days_by_rider[name]), 'delivered': d,
            'incentive': 0, 'payout': payout, 'total_payout': payout,
            'insurance': insurance, 'final_payout': final, 'scheme': 'MG+Distance',
        })
        total_delivered += d
        total_payout += payout
        total_final += final

    try:
        date_min = str(pd.to_datetime(df[date_col], errors='coerce').min().date())
        date_max = str(pd.to_datetime(df[date_col], errors='coerce').max().date())
    except Exception:
        date_min = date_max = None

    return {
        'rate': None, 'insurance': insurance,
        'rider_count': len(riders), 'row_count': len(df),
        'date_min': date_min, 'date_max': date_max,
        'riders': rider_rows,
        'dcpo_range': None,
        'totals': {
            'delivered': total_delivered, 'incentive': 0,
            'payout': total_payout, 'total_payout': total_payout,
            'insurance': insurance * len(riders), 'final_payout': total_final,
        },
    }


def parse_tiers(s):
    """'7:45:640,6:45:420' -> [(7,45,640), (6,45,420)]"""
    tiers = []
    for part in s.split(','):
        d, orders, amt = part.split(':')
        tiers.append((int(d), int(orders), float(amt)))
    tiers.sort(key=lambda t: (t[0], t[1]), reverse=True)
    return tiers


def login_amount(days, orders, tiers):
    for min_days, min_orders, amount in tiers:
        if days >= min_days and orders >= min_orders:
            return amount
    return 0


def min_tier_days(tiers):
    return min(t[0] for t in tiers) if tiers else 0


def style_header_row(ws, headers):
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=j, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = BORDER
        c.alignment = HEADER_ALIGN


def mg_payout_value(delivered, base, threshold, rate):
    """Python equivalent of the MG rider formula, e.g.
    IF(delivered<=12, 620, 620+(delivered-12)*rate). Used for preview calcs."""
    if delivered <= threshold:
        return base
    return base + (delivered - threshold) * rate


def mg_formula_str(delivered_cell, base, threshold, rate):
    """Excel formula string equivalent, for writing into the workbook."""
    return f'=IF({delivered_cell}<={threshold},{base},{base}+({delivered_cell}-{threshold})*{rate})'


def build_trip_sheet(wb, title, df, rider_col, date_col, delivered_col, formula_fn, is_first=False,
                      payout_label='PAYOUT', trip_id_col=None, shipment_lookup=None):
    """Builds a trip-level tab: raw sheet as-is + a payout column, where each
    row's formula is produced by formula_fn(delivered_cell_ref) -> formula string.
    Used for ALPHA's PAYOUT tab, BETA/Delta's PER ORDER and MG ORDERS tabs.

    If shipment_lookup is given, an extra 'SHIPMENT CHECK' column is appended
    after the payout column: green 'CORRECT' if the trip's delivered count
    matches the shipment data, red 'INCORRECT' otherwise. Doesn't affect any
    payout math — purely a flag for manual review."""
    ws = wb.active if is_first else wb.create_sheet()
    ws.title = title
    orig_cols = list(df.columns)
    style_header_row(ws, orig_cols)
    payout_col_idx = len(orig_cols) + 1
    c = ws.cell(row=1, column=payout_col_idx, value=payout_label)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.border = BORDER
    c.alignment = HEADER_ALIGN

    check_col_idx = None
    if shipment_lookup is not None:
        check_col_idx = payout_col_idx + 1
        c = ws.cell(row=1, column=check_col_idx, value='SHIPMENT CHECK')
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = BORDER
        c.alignment = HEADER_ALIGN

    df_sorted = df.sort_values([rider_col, date_col]).reset_index(drop=True)
    delivered_idx = orig_cols.index(delivered_col) + 1
    date_idx = orig_cols.index(date_col) + 1
    trip_idx = orig_cols.index(trip_id_col) + 1 if trip_id_col and trip_id_col in orig_cols else None

    for i, row in df_sorted.iterrows():
        r = i + 2
        for j, col in enumerate(orig_cols, start=1):
            val = row.get(col, None)
            if pd.isna(val):
                val = None
            cell = ws.cell(row=r, column=j, value=val)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = CENTER
            if j == date_idx:
                cell.number_format = 'dd-mm-yyyy'
        delivered_cell = f'{get_column_letter(delivered_idx)}{r}'
        pay_cell = ws.cell(row=r, column=payout_col_idx, value=formula_fn(delivered_cell))
        pay_cell.font = BODY_FONT
        pay_cell.border = BORDER
        pay_cell.alignment = CENTER

        if check_col_idx:
            trip_val = row.get(trip_id_col) if trip_idx else None
            delivered_val = row.get(delivered_col)
            check_cell = ws.cell(row=r, column=check_col_idx)
            check_cell.border = BORDER
            check_cell.alignment = CENTER
            apply_shipment_check_cell(check_cell, trip_val, delivered_val, shipment_lookup)

    last_col = check_col_idx or payout_col_idx
    for j in range(1, last_col + 1):
        ws.column_dimensions[get_column_letter(j)].width = 22

    return ws


def build_payout_sheet(wb, df, rider_col, date_col, delivered_col, rate, trip_id_col=None, shipment_lookup=None):
    """Backwards-compatible wrapper: builds the shared PAYOUT tab (raw data +
    PAYOUT = delivered * rate). Used by ALPHA and single-sheet BETA stores."""
    return build_trip_sheet(
        wb, 'PAYOUT', df, rider_col, date_col, delivered_col,
        formula_fn=lambda cell: f'={cell}*{rate}',
        is_first=True, trip_id_col=trip_id_col, shipment_lookup=shipment_lookup,
    )


def build_tiered_summary(raw_path, out_path, rate, insurance, tiers, shipment_lookup=None):
    df, date_col, rider_col, delivered_col, store_col = load_sheet1(raw_path)
    trip_id_col = find_col(df.columns, 'TRIP', 'ID')

    grp = df.groupby(rider_col)
    delivered_by_rider = grp[delivered_col].sum()
    login_days_by_rider = grp[date_col].nunique()
    riders = sorted(delivered_by_rider.index.tolist())

    wb = Workbook()
    build_payout_sheet(wb, df, rider_col, date_col, delivered_col, rate,
                        trip_id_col=trip_id_col, shipment_lookup=shipment_lookup)

    # PAYOUT DATA: untouched copy of raw Sheet2, for reference only
    copy_sheet2_untouched(raw_path, wb, 'PAYOUT DATA')

    # FINAL PAYOUT: computed from Sheet1 only
    ws3 = wb.create_sheet('FINAL PAYOUT')
    headers3 = ['RIDERS NAME', 'ATTENDANCE', 'TOTAL DELIVERED ORDERS', 'INCENTIVE', 'PAYOUT',
                'TOTAL PAYOUT', 'Insurance Deduction', 'FINAL PAYOUT', 'DCPO']
    style_header_row(ws3, headers3)

    for i, name in enumerate(riders):
        r = i + 2
        d = int(delivered_by_rider[name])
        ld = int(login_days_by_rider[name])
        la = login_amount(ld, d, tiers)
        payout = rate * d
        ws3.cell(row=r, column=1, value=name).font = BODY_FONT
        ws3.cell(row=r, column=2, value=ld).font = BODY_FONT
        ws3.cell(row=r, column=3, value=d).font = BODY_FONT
        ws3.cell(row=r, column=4, value=la).font = BODY_FONT
        ws3.cell(row=r, column=5, value=payout).font = BODY_FONT
        ws3.cell(row=r, column=6, value=f'=E{r}+D{r}').font = BODY_FONT
        ws3.cell(row=r, column=7, value=insurance).font = BODY_FONT
        ws3.cell(row=r, column=8, value=f'=F{r}-G{r}').font = BODY_FONT
        ws3.cell(row=r, column=9, value=f'=ROUND(F{r}/C{r},0)').font = BODY_FONT
        for col in range(1, 10):
            cell = ws3.cell(row=r, column=col)
            cell.alignment = CENTER
            cell.border = BORDER

    total_row3 = len(riders) + 2
    ws3.cell(row=total_row3, column=1, value='TOTAL').font = TOTAL_FONT
    for col_letter in ['C', 'D', 'E', 'F', 'G', 'H']:
        ws3.cell(row=total_row3, column=ord(col_letter) - 64,
                 value=f'=SUM({col_letter}2:{col_letter}{total_row3-1})').font = TOTAL_FONT
    ws3.cell(row=total_row3, column=9, value=f'=ROUND(F{total_row3}/C{total_row3},0)').font = TOTAL_FONT
    for col in range(1, 10):
        cell = ws3.cell(row=total_row3, column=col)
        cell.fill = TOTAL_FILL
        cell.alignment = CENTER
        cell.border = BORDER

    widths3 = [16.3, 14.9, 29.3, 11.8, 9.4, 16.8, 21.3, 16, 6.8]
    for j, w in enumerate(widths3, start=1):
        ws3.column_dimensions[get_column_letter(j)].width = w

    wb.save(out_path)
    return out_path


def build_tiered_split(raw_path, out_path, rate, insurance, tiers, mg_config=None, shipment_lookup=None,
                trip_tab_title=None):
    mg_sheet_name = find_mg_sheet_name(raw_path) if mg_config else None
    po_sheet_name = find_non_mg_sheet_name(raw_path, mg_sheet_name) if mg_sheet_name else 0

    df_po, date_col, rider_col, delivered_col, store_col = load_sheet(raw_path, po_sheet_name)
    trip_id_col = find_col(df_po.columns, 'TRIP', 'ID')

    grp_po = df_po.groupby(rider_col)
    delivered_by_rider_po = grp_po[delivered_col].sum()
    login_days_by_rider_po = grp_po[date_col].nunique()
    po_riders = sorted(delivered_by_rider_po.index.tolist())
    threshold_days = min_tier_days(tiers)
    # INCENTIVE only ever applies to per-order-scheme riders
    incentive_riders = [n for n in po_riders if login_days_by_rider_po[n] >= threshold_days]

    all_riders = list(po_riders)  # insurance covers everyone; MG riders added below if present

    wb = Workbook()

    if mg_sheet_name:
        df_mg, mg_date_col, mg_rider_col, mg_delivered_col, _ = load_sheet(raw_path, mg_sheet_name)
        mg_trip_id_col = find_col(df_mg.columns, 'TRIP', 'ID')
        base, threshold, mg_rate = mg_config['base'], mg_config['threshold'], mg_config['rate']
        build_trip_sheet(
            wb, 'MG ORDERS', df_mg, mg_rider_col, mg_date_col, mg_delivered_col,
            formula_fn=lambda cell: mg_formula_str(cell, base, threshold, mg_rate),
            is_first=True, trip_id_col=mg_trip_id_col, shipment_lookup=shipment_lookup,
        )
        build_trip_sheet(
            wb, 'PER ORDER', df_po, rider_col, date_col, delivered_col,
            formula_fn=lambda cell: f'={cell}*{rate}',
            is_first=False, trip_id_col=trip_id_col, shipment_lookup=shipment_lookup,
        )
        mg_riders = sorted(df_mg[mg_rider_col].dropna().unique().tolist())
        all_riders = sorted(set(po_riders) | set(mg_riders))
    else:
        build_trip_sheet(
            wb, trip_tab_title or 'PAYOUT', df_po, rider_col, date_col, delivered_col,
            formula_fn=lambda cell: f'={cell}*{rate}',
            is_first=True, trip_id_col=trip_id_col, shipment_lookup=shipment_lookup,
        )

    # INCENTIVE tab: pivot-style summary, riders below the lowest tier excluded entirely
    ws2 = wb.create_sheet('INCENTIVE')
    headers2 = ['RIDER NAME', 'ATTENDANCE', 'TOTAL DELIVERED ORDERS', 'INCENTIVE']
    style_header_row(ws2, headers2)
    for i, name in enumerate(incentive_riders):
        r = i + 2
        d = int(delivered_by_rider_po[name])
        ld = int(login_days_by_rider_po[name])
        la = login_amount(ld, d, tiers)
        ws2.cell(row=r, column=1, value=name).font = BODY_FONT
        ws2.cell(row=r, column=2, value=ld).font = BODY_FONT
        ws2.cell(row=r, column=3, value=d).font = BODY_FONT
        ws2.cell(row=r, column=4, value=la).font = BODY_FONT
        for col in range(1, 5):
            cell = ws2.cell(row=r, column=col)
            cell.alignment = CENTER
            cell.border = BORDER
    total_row2 = len(incentive_riders) + 2
    ws2.cell(row=total_row2, column=1, value='TOTAL').font = TOTAL_FONT
    for col_letter in ['C', 'D']:
        ws2.cell(row=total_row2, column=ord(col_letter) - 64,
                 value=f'=SUM({col_letter}2:{col_letter}{total_row2-1})').font = TOTAL_FONT
    for col in range(1, 5):
        cell = ws2.cell(row=total_row2, column=col)
        cell.fill = TOTAL_FILL
        cell.alignment = CENTER
        cell.border = BORDER
    for j, w in enumerate([22, 13, 20, 12], start=1):
        ws2.column_dimensions[get_column_letter(j)].width = w

    # INSURANCE DEDUCTION tab: every rider across all schemes, flat amount
    ws3 = wb.create_sheet('INSURANCE DEDUCTION')
    headers3 = ['RIDER NAME', 'INSURANCE DEDUCTION']
    style_header_row(ws3, headers3)
    for i, name in enumerate(all_riders):
        r = i + 2
        ws3.cell(row=r, column=1, value=name).font = BODY_FONT
        ws3.cell(row=r, column=2, value=insurance).font = BODY_FONT
        for col in range(1, 3):
            cell = ws3.cell(row=r, column=col)
            cell.alignment = CENTER
            cell.border = BORDER
    total_row3 = len(all_riders) + 2
    ws3.cell(row=total_row3, column=1, value='TOTAL').font = TOTAL_FONT
    ws3.cell(row=total_row3, column=2,
             value=f'=SUM(B2:B{total_row3-1})').font = TOTAL_FONT
    for col in range(1, 3):
        cell = ws3.cell(row=total_row3, column=col)
        cell.fill = TOTAL_FILL
        cell.alignment = CENTER
        cell.border = BORDER
    for j, w in enumerate([22, 20], start=1):
        ws3.column_dimensions[get_column_letter(j)].width = w

    wb.save(out_path)
    return out_path


def restrict_riders_by_attendance_vendor(raw_path, df, rider_col, vendor_name):
    """Some raw exports mix in a worker from a different vendor/role at the
    same physical store (e.g. a picker paid per-order, not a rider) — not
    part of this company's payout at all. If the raw file carries an
    attendance sheet with 'Vendor' and 'Name' columns, restrict df to only
    riders whose attendance-sheet Vendor matches vendor_name. Matched by
    first-name token since trip-sheet names are sometimes shortened
    (e.g. 'Rahul' in the trip data vs 'Rahul Verma' in attendance)."""
    att_sheet = find_sheet_by_keyword(raw_path, 'ATTEN')
    if not att_sheet:
        return df
    att = pd.read_excel(raw_path, sheet_name=att_sheet)
    att.columns = [str(c).strip() for c in att.columns]
    vendor_col = find_col(att.columns, 'VENDOR')
    name_col = find_col(att.columns, 'NAME')
    if not vendor_col or not name_col:
        return df

    def first_token(s):
        s = str(s).strip()
        return s.split()[0].upper() if s else ''

    allowed_tokens = {
        first_token(n) for n, v in zip(att[name_col], att[vendor_col])
        if str(v).strip().upper() == vendor_name.strip().upper()
    }
    if not allowed_tokens:
        return df
    mask = df[rider_col].apply(lambda n: first_token(n) in allowed_tokens)
    return df[mask].reset_index(drop=True)


def build_distance_mg(raw_path, out_path, insurance, mg_flat, per_km_rate, shipment_lookup=None,
                    distance_col_override=None, restrict_vendor=None):
    """Gamma's payout model has no per-order rate at all: every trip pays a
    flat MG amount plus a per-km distance bonus. Output: PAYOUT + INSURANCE
    DEDUCTION tabs only (no INCENTIVE tab, since there's no order-based tier).

    distance_col_override: some stores put the per-km distance figure under
    a differently-named column (e.g. 'KM TRAVELLED' instead of
    'DISTANCE') — pass the exact header text to use that column instead.
    restrict_vendor: if given, drop any rider whose attendance-sheet Vendor
    doesn't match this (see restrict_riders_by_attendance_vendor)."""
    df = pd.read_excel(raw_path, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    df = filter_absent_rows(df)
    date_col = find_col(df.columns, 'DATE')
    rider_col = find_col(df.columns, 'RIDER', 'NAME')
    if distance_col_override:
        distance_col = next((c for c in df.columns
                              if str(c).strip().upper() == distance_col_override.strip().upper()), None)
    else:
        distance_col = find_col(df.columns, 'DISTANCE')
    delivered_col = find_col(df.columns, 'DELIVERED', 'ORDER')
    trip_id_col = find_col(df.columns, 'TRIP', 'ID')
    missing = [n for n, c in [('DATE', date_col), ('RIDER NAME', rider_col),
                               ('DISTANCE', distance_col)] if c is None]
    if missing:
        raise MissingColumnError(0, missing, list(df.columns))
    df = normalize_date_col(df, date_col)
    if restrict_vendor:
        df = restrict_riders_by_attendance_vendor(raw_path, df, rider_col, restrict_vendor)

    wb = Workbook()
    ws = wb.active
    ws.title = 'PAYOUT'
    orig_cols = list(df.columns)
    style_header_row(ws, orig_cols)

    extra_idx = len(orig_cols) + 1
    mg_idx = len(orig_cols) + 2
    total_idx = len(orig_cols) + 3
    check_idx = total_idx + 1 if shipment_lookup is not None else None
    for idx, label in [(extra_idx, 'Extra km Amount'), (mg_idx, 'MG Amount'),
                        (total_idx, 'Total Calculated Amount')]:
        c = ws.cell(row=1, column=idx, value=label)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = BORDER
        c.alignment = HEADER_ALIGN
    if check_idx:
        c = ws.cell(row=1, column=check_idx, value='SHIPMENT CHECK')
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = BORDER
        c.alignment = HEADER_ALIGN

    df_sorted = df.sort_values([rider_col, date_col]).reset_index(drop=True)
    distance_idx = orig_cols.index(distance_col) + 1
    date_idx = orig_cols.index(date_col) + 1

    for i, row in df_sorted.iterrows():
        r = i + 2
        for j, col in enumerate(orig_cols, start=1):
            val = row.get(col, None)
            if pd.isna(val):
                val = None
            cell = ws.cell(row=r, column=j, value=val)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = CENTER
            if j == date_idx:
                cell.number_format = 'dd-mm-yyyy'
        dist_cell = f'{get_column_letter(distance_idx)}{r}'
        extra_cell = ws.cell(row=r, column=extra_idx, value=f'={dist_cell}*{per_km_rate}')
        extra_cell.font = BODY_FONT
        extra_cell.border = BORDER
        extra_cell.alignment = CENTER
        mg_cell = ws.cell(row=r, column=mg_idx, value=mg_flat)
        mg_cell.font = BODY_FONT
        mg_cell.border = BORDER
        mg_cell.alignment = CENTER
        extra_letter = get_column_letter(extra_idx)
        mg_letter = get_column_letter(mg_idx)
        total_cell = ws.cell(row=r, column=total_idx, value=f'={extra_letter}{r}+{mg_letter}{r}')
        total_cell.font = BODY_FONT
        total_cell.border = BORDER
        total_cell.alignment = CENTER

        if check_idx:
            trip_val = row.get(trip_id_col) if trip_id_col else None
            delivered_val = row.get(delivered_col) if delivered_col else None
            check_cell = ws.cell(row=r, column=check_idx)
            check_cell.border = BORDER
            check_cell.alignment = CENTER
            apply_shipment_check_cell(check_cell, trip_val, delivered_val, shipment_lookup)

    for j in range(1, (check_idx or total_idx) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 20

    # INSURANCE DEDUCTION tab: every rider, flat amount
    all_riders = sorted(df[rider_col].dropna().unique().tolist())
    ws2 = wb.create_sheet('INSURANCE DEDUCTION')
    style_header_row(ws2, ['RIDER NAME', 'INSURANCE DEDUCTION'])
    for i, name in enumerate(all_riders):
        r = i + 2
        ws2.cell(row=r, column=1, value=name).font = BODY_FONT
        ws2.cell(row=r, column=2, value=insurance).font = BODY_FONT
        for col in range(1, 3):
            cell = ws2.cell(row=r, column=col)
            cell.alignment = CENTER
            cell.border = BORDER
    total_row = len(all_riders) + 2
    ws2.cell(row=total_row, column=1, value='TOTAL').font = TOTAL_FONT
    ws2.cell(row=total_row, column=2, value=f'=SUM(B2:B{total_row-1})').font = TOTAL_FONT
    for col in range(1, 3):
        cell = ws2.cell(row=total_row, column=col)
        cell.fill = TOTAL_FILL
        cell.alignment = CENTER
        cell.border = BORDER
    for j, w in enumerate([22, 20], start=1):
        ws2.column_dimensions[get_column_letter(j)].width = w

    wb.save(out_path)
    return out_path


def build_mg_and_tiers(raw_path, out_path, rate, mg_config=None, tiers=None, shipment_lookup=None):
    """Delta has no insurance-deduction tab on any store. Beyond that, its
    tab set varies by store:
    - North Hub-style (has an 'MG' raw sheet): MG ORDERS + PER ORDER tabs.
    - Riverside Garden-style (single raw sheet, has tiers): PAYOUT + INCENTIVE tabs.
    The INCENTIVE tab's TOTAL row is omitted when only one rider qualifies —
    mirrors the layout of the vendor's reference workbooks."""
    mg_sheet_name = find_mg_sheet_name(raw_path) if mg_config else None
    po_sheet_name = find_non_mg_sheet_name(raw_path, mg_sheet_name) if mg_sheet_name else 0

    df_po, date_col, rider_col, delivered_col, _ = load_sheet(raw_path, po_sheet_name)
    trip_id_col = find_col(df_po.columns, 'TRIP', 'ID')

    wb = Workbook()
    trip_title = 'PER ORDER' if mg_sheet_name else 'PAYOUT'

    if mg_sheet_name:
        df_mg, mg_date_col, mg_rider_col, mg_delivered_col, _ = load_sheet(raw_path, mg_sheet_name)
        mg_trip_id_col = find_col(df_mg.columns, 'TRIP', 'ID')
        base, threshold, mg_rate = mg_config['base'], mg_config['threshold'], mg_config['rate']
        build_trip_sheet(
            wb, 'MG ORDERS', df_mg, mg_rider_col, mg_date_col, mg_delivered_col,
            formula_fn=lambda cell: mg_formula_str(cell, base, threshold, mg_rate),
            is_first=True, payout_label='MG', trip_id_col=mg_trip_id_col, shipment_lookup=shipment_lookup,
        )
        build_trip_sheet(
            wb, trip_title, df_po, rider_col, date_col, delivered_col,
            formula_fn=lambda cell: f'={cell}*{rate}',
            is_first=False, trip_id_col=trip_id_col, shipment_lookup=shipment_lookup,
        )
    else:
        build_trip_sheet(
            wb, trip_title, df_po, rider_col, date_col, delivered_col,
            formula_fn=lambda cell: f'={cell}*{rate}',
            is_first=True, trip_id_col=trip_id_col, shipment_lookup=shipment_lookup,
        )

    if tiers:
        grp = df_po.groupby(rider_col)
        delivered_by_rider = grp[delivered_col].sum()
        days_by_rider = grp[date_col].nunique()
        threshold_days = min_tier_days(tiers)
        incentive_riders = sorted(
            n for n in delivered_by_rider.index
            if days_by_rider[n] >= threshold_days
            and login_amount(int(days_by_rider[n]), int(delivered_by_rider[n]), tiers) > 0
        )

        ws_inc = wb.create_sheet('INCENTIVE')
        style_header_row(ws_inc, ['RIDER NAME', 'ATTENDANCE', 'TOTAL DELIVERED ORDERS', 'INCENTIVE'])
        for i, name in enumerate(incentive_riders):
            r = i + 2
            d = int(delivered_by_rider[name])
            ld = int(days_by_rider[name])
            la = login_amount(ld, d, tiers)
            ws_inc.cell(row=r, column=1, value=name).font = BODY_FONT
            ws_inc.cell(row=r, column=2, value=ld).font = BODY_FONT
            ws_inc.cell(row=r, column=3, value=d).font = BODY_FONT
            ws_inc.cell(row=r, column=4, value=la).font = BODY_FONT
            for col in range(1, 5):
                cell = ws_inc.cell(row=r, column=col)
                cell.alignment = CENTER
                cell.border = BORDER
        if len(incentive_riders) > 1:
            total_row = len(incentive_riders) + 2
            ws_inc.cell(row=total_row, column=1, value='TOTAL').font = TOTAL_FONT
            for col_letter in ['C', 'D']:
                ws_inc.cell(row=total_row, column=ord(col_letter) - 64,
                            value=f'=SUM({col_letter}2:{col_letter}{total_row-1})').font = TOTAL_FONT
            for col in range(1, 5):
                cell = ws_inc.cell(row=total_row, column=col)
                cell.fill = TOTAL_FILL
                cell.alignment = CENTER
                cell.border = BORDER
        for j, w in enumerate([22, 13, 20, 12], start=1):
            ws_inc.column_dimensions[get_column_letter(j)].width = w

    wb.save(out_path)
    return out_path


def build_mg_per_trip(raw_path, out_path, insurance, base, threshold, rate, shipment_lookup=None):
    """Epsilon Group's scheme: every trip (rider-day row) is paid its own MG
    formula directly off that row's delivered orders — no per-order rate, no
    incentive tiers, no aggregating first. Output: PAYOUT + INSURANCE
    DEDUCTION tabs only. (Any equipment-related deductions are handled
    outside this tool and intentionally not modeled here.)"""
    df, date_col, rider_col, delivered_col, store_col = load_sheet1(raw_path)
    df = df[df[rider_col].notna()].reset_index(drop=True)
    trip_id_col = find_col(df.columns, 'TRIP', 'ID')

    wb = Workbook()
    build_trip_sheet(
        wb, 'PAYOUT', df, rider_col, date_col, delivered_col,
        formula_fn=lambda cell: mg_formula_str(cell, base, threshold, rate),
        is_first=True, trip_id_col=trip_id_col, shipment_lookup=shipment_lookup,
    )

    riders = sorted(df[rider_col].dropna().unique().tolist())
    ws2 = wb.create_sheet('INSURANCE DEDUCTION')
    style_header_row(ws2, ['RIDER NAME', 'INSURANCE DEDUCTION'])
    for i, name in enumerate(riders):
        r = i + 2
        ws2.cell(row=r, column=1, value=name).font = BODY_FONT
        ws2.cell(row=r, column=2, value=insurance).font = BODY_FONT
        for col in range(1, 3):
            cell = ws2.cell(row=r, column=col)
            cell.alignment = CENTER
            cell.border = BORDER
    total_row2 = len(riders) + 2
    ws2.cell(row=total_row2, column=1, value='TOTAL').font = TOTAL_FONT
    ws2.cell(row=total_row2, column=2, value=f'=SUM(B2:B{total_row2-1})').font = TOTAL_FONT
    for col in range(1, 3):
        cell = ws2.cell(row=total_row2, column=col)
        cell.fill = TOTAL_FILL
        cell.alignment = CENTER
        cell.border = BORDER
    for j, w in enumerate([22, 20], start=1):
        ws2.column_dimensions[get_column_letter(j)].width = w

    wb.save(out_path)
    return out_path


def build_flat_per_trip_insured(raw_path, out_path, insurance, flat_rate, shipment_lookup=None):
    """Omega's scheme: every trip (rider-day row) pays the same flat amount
    regardless of delivered orders — no threshold, no per-order rate, no
    incentive tiers. Output: PAYOUT + INSURANCE DEDUCTION tabs only.
    (Any other deductions are handled outside this tool and intentionally
    not modeled here.)"""
    df, date_col, rider_col, delivered_col, store_col = load_sheet1(raw_path)
    df = df[df[rider_col].notna()].reset_index(drop=True)
    trip_id_col = find_col(df.columns, 'TRIP', 'ID')

    wb = Workbook()
    build_trip_sheet(
        wb, 'PAYOUT', df, rider_col, date_col, delivered_col,
        formula_fn=lambda cell: flat_rate,
        is_first=True, trip_id_col=trip_id_col, shipment_lookup=shipment_lookup,
    )

    riders = sorted(df[rider_col].dropna().unique().tolist())
    ws2 = wb.create_sheet('INSURANCE DEDUCTION')
    style_header_row(ws2, ['RIDER NAME', 'INSURANCE DEDUCTION'])
    for i, name in enumerate(riders):
        r = i + 2
        ws2.cell(row=r, column=1, value=name).font = BODY_FONT
        ws2.cell(row=r, column=2, value=insurance).font = BODY_FONT
        for col in range(1, 3):
            cell = ws2.cell(row=r, column=col)
            cell.alignment = CENTER
            cell.border = BORDER
    total_row2 = len(riders) + 2
    ws2.cell(row=total_row2, column=1, value='TOTAL').font = TOTAL_FONT
    ws2.cell(row=total_row2, column=2, value=f'=SUM(B2:B{total_row2-1})').font = TOTAL_FONT
    for col in range(1, 3):
        cell = ws2.cell(row=total_row2, column=col)
        cell.fill = TOTAL_FILL
        cell.alignment = CENTER
        cell.border = BORDER
    for j, w in enumerate([22, 20], start=1):
        ws2.column_dimensions[get_column_letter(j)].width = w

    wb.save(out_path)
    return out_path


def build_delivered_plus_cancelled(raw_path, out_path, rate, cancelled_rate, shipment_lookup=None):
    """Zeta's payout formula uses two raw columns directly:
    PAYOUT = (delivered * rate) + (cancelled * cancelled_rate).
    No insurance, no incentive tiers for this vendor. Single PAYOUT tab only."""
    df = pd.read_excel(raw_path, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    df = filter_absent_rows(df)
    date_col = find_col(df.columns, 'DATE')
    rider_col = find_col(df.columns, 'RIDER', 'NAME')
    delivered_col = find_col(df.columns, 'DELIVERED', 'ORDER')
    cancelled_col = find_col(df.columns, 'CANCELLED', 'ORDER')
    trip_id_col = find_col(df.columns, 'TRIP', 'ID')

    missing = [n for n, c in [('DATE', date_col), ('RIDER NAME', rider_col),
                               ('TOTAL DELIVERED ORDERS', delivered_col),
                               ('TOTAL CANCELLED ORDERS', cancelled_col)] if c is None]
    if missing:
        raise MissingColumnError(0, missing, list(df.columns))
    df = normalize_date_col(df, date_col)
    df = df[df[rider_col].notna()].reset_index(drop=True)
    # some Zeta raw files already ship with their own 'PAYOUT' column (with a
    # formula baked in) — drop it so we don't end up with two PAYOUT columns;
    # we rebuild it ourselves below with the verified formula.
    df = df[[c for c in df.columns if c.strip().upper() != 'PAYOUT']]

    wb = Workbook()
    ws = wb.active
    ws.title = 'PAYOUT'
    orig_cols = list(df.columns)
    style_header_row(ws, orig_cols)
    payout_idx = len(orig_cols) + 1
    check_idx = payout_idx + 1 if shipment_lookup is not None else None

    c = ws.cell(row=1, column=payout_idx, value='PAYOUT')
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.border = BORDER
    c.alignment = HEADER_ALIGN
    if check_idx:
        c = ws.cell(row=1, column=check_idx, value='SHIPMENT CHECK')
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = BORDER
        c.alignment = HEADER_ALIGN

    df_sorted = df.sort_values([rider_col, date_col]).reset_index(drop=True)
    delivered_idx = orig_cols.index(delivered_col) + 1
    cancelled_idx = orig_cols.index(cancelled_col) + 1
    date_idx = orig_cols.index(date_col) + 1

    for i, row in df_sorted.iterrows():
        r = i + 2
        for j, col in enumerate(orig_cols, start=1):
            val = row.get(col, None)
            if pd.isna(val):
                val = None
            cell = ws.cell(row=r, column=j, value=val)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = CENTER
            if j == date_idx:
                cell.number_format = 'dd-mm-yyyy'
        delivered_cell = f'{get_column_letter(delivered_idx)}{r}'
        cancelled_cell = f'{get_column_letter(cancelled_idx)}{r}'
        pay_cell = ws.cell(row=r, column=payout_idx,
                            value=f'=({delivered_cell}*{rate})+({cancelled_cell}*{cancelled_rate})')
        pay_cell.font = BODY_FONT
        pay_cell.border = BORDER
        pay_cell.alignment = CENTER

        if check_idx:
            trip_val = row.get(trip_id_col) if trip_id_col else None
            delivered_val = row.get(delivered_col)
            check_cell = ws.cell(row=r, column=check_idx)
            check_cell.border = BORDER
            check_cell.alignment = CENTER
            apply_shipment_check_cell(check_cell, trip_val, delivered_val, shipment_lookup)

    for j in range(1, (check_idx or payout_idx) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 22

    wb.save(out_path)
    return out_path


def build_delivered_plus_cancelled_tiered(raw_path, out_path, rate, cancelled_rate, tiers, shipment_lookup=None):
    """Theta's payout formula (like Zeta's) uses two raw columns directly:
    PAYOUT = (delivered * rate) + (cancelled * cancelled_rate).
    On top of that, Theta also pays an attendance/delivery incentive, so this
    adds an INCENTIVE tab (Delta-style: tiers keyed on login days + delivered
    orders, zero-incentive riders omitted from the tab entirely) after the
    PAYOUT tab. No insurance deduction tab for this vendor.

    NOTE: per-order remark-based adjustments (e.g. penalties for flagged
    orders) are not modeled here — a possible future extension."""
    df = pd.read_excel(raw_path, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    df = filter_absent_rows(df)
    date_col = find_col(df.columns, 'DATE')
    rider_col = find_col(df.columns, 'RIDER', 'NAME')
    delivered_col = find_col(df.columns, 'DELIVERED', 'ORDER')
    cancelled_col = find_col(df.columns, 'CANCELLED', 'ORDER')
    trip_id_col = find_col(df.columns, 'TRIP', 'ID')

    missing = [n for n, c in [('DATE', date_col), ('RIDER NAME', rider_col),
                               ('TOTAL DELIVERED ORDERS', delivered_col),
                               ('TOTAL CANCELLED ORDERS', cancelled_col)] if c is None]
    if missing:
        raise MissingColumnError(0, missing, list(df.columns))
    df = normalize_date_col(df, date_col)
    df = df[df[rider_col].notna()].reset_index(drop=True)
    df = df[[c for c in df.columns if c.strip().upper() != 'PAYOUT']]

    wb = Workbook()
    ws = wb.active
    ws.title = 'PAYOUT'
    orig_cols = list(df.columns)
    style_header_row(ws, orig_cols)
    payout_idx = len(orig_cols) + 1
    check_idx = payout_idx + 1 if shipment_lookup is not None else None

    c = ws.cell(row=1, column=payout_idx, value='PAYOUT')
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.border = BORDER
    c.alignment = HEADER_ALIGN
    if check_idx:
        c = ws.cell(row=1, column=check_idx, value='SHIPMENT CHECK')
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = BORDER
        c.alignment = HEADER_ALIGN

    df_sorted = df.sort_values([rider_col, date_col]).reset_index(drop=True)
    delivered_idx = orig_cols.index(delivered_col) + 1
    cancelled_idx = orig_cols.index(cancelled_col) + 1
    date_idx = orig_cols.index(date_col) + 1

    for i, row in df_sorted.iterrows():
        r = i + 2
        for j, col in enumerate(orig_cols, start=1):
            val = row.get(col, None)
            if pd.isna(val):
                val = None
            cell = ws.cell(row=r, column=j, value=val)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = CENTER
            if j == date_idx:
                cell.number_format = 'dd-mm-yyyy'
        delivered_cell = f'{get_column_letter(delivered_idx)}{r}'
        cancelled_cell = f'{get_column_letter(cancelled_idx)}{r}'
        pay_cell = ws.cell(row=r, column=payout_idx,
                            value=f'=({delivered_cell}*{rate})+({cancelled_cell}*{cancelled_rate})')
        pay_cell.font = BODY_FONT
        pay_cell.border = BORDER
        pay_cell.alignment = CENTER

        if check_idx:
            trip_val = row.get(trip_id_col) if trip_id_col else None
            delivered_val = row.get(delivered_col)
            check_cell = ws.cell(row=r, column=check_idx)
            check_cell.border = BORDER
            check_cell.alignment = CENTER
            apply_shipment_check_cell(check_cell, trip_val, delivered_val, shipment_lookup)

    for j in range(1, (check_idx or payout_idx) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 22

    if tiers:
        grp = df.groupby(rider_col)
        delivered_by_rider = grp[delivered_col].sum()
        days_by_rider = grp[date_col].nunique()
        threshold_days = min_tier_days(tiers)
        incentive_riders = sorted(
            n for n in delivered_by_rider.index
            if days_by_rider[n] >= threshold_days
            and login_amount(int(days_by_rider[n]), int(delivered_by_rider[n]), tiers) > 0
        )

        ws_inc = wb.create_sheet('INCENTIVE')
        style_header_row(ws_inc, ['RIDER NAME', 'ATTENDANCE', 'TOTAL DELIVERED ORDERS', 'INCENTIVE'])
        for i, name in enumerate(incentive_riders):
            r = i + 2
            d = int(delivered_by_rider[name])
            ld = int(days_by_rider[name])
            la = login_amount(ld, d, tiers)
            ws_inc.cell(row=r, column=1, value=name).font = BODY_FONT
            ws_inc.cell(row=r, column=2, value=ld).font = BODY_FONT
            ws_inc.cell(row=r, column=3, value=d).font = BODY_FONT
            ws_inc.cell(row=r, column=4, value=la).font = BODY_FONT
            for col in range(1, 5):
                cell = ws_inc.cell(row=r, column=col)
                cell.alignment = CENTER
                cell.border = BORDER
        if len(incentive_riders) > 1:
            total_row = len(incentive_riders) + 2
            ws_inc.cell(row=total_row, column=1, value='TOTAL').font = TOTAL_FONT
            for col_letter in ['C', 'D']:
                ws_inc.cell(row=total_row, column=ord(col_letter) - 64,
                            value=f'=SUM({col_letter}2:{col_letter}{total_row-1})').font = TOTAL_FONT
            for col in range(1, 5):
                cell = ws_inc.cell(row=total_row, column=col)
                cell.fill = TOTAL_FILL
                cell.alignment = CENTER
                cell.border = BORDER
        for j, w in enumerate([22, 13, 20, 12], start=1):
            ws_inc.column_dimensions[get_column_letter(j)].width = w

    wb.save(out_path)
    return out_path


def build_flat_per_day(raw_path, out_path, insurance, per_day_rate, shipment_lookup=None):
    """Kappa's scheme: every day a rider is present pays the same flat
    amount (e.g. ₹660), regardless of delivered-order count — so no rate,
    no incentive tiers, and shipment cross-checking doesn't apply (payout
    isn't order-driven). Output tabs: PAYOUT DATA (raw trip data + flat
    PAYOUT column) / FINAL PAYOUT (per-rider rollup). The raw file also
    carries an attendance sheet and a bank-details sheet, copied through to
    the final workbook exactly as-is; any other extra sheet in the raw file
    (e.g. a junk 'CAL' tab) is simply not carried over."""
    df, date_col, rider_col, delivered_col, store_col = load_sheet1(raw_path)
    trip_id_col = find_col(df.columns, 'TRIP', 'ID')

    wb = Workbook()
    build_trip_sheet(
        wb, 'PAYOUT DATA', df, rider_col, date_col, delivered_col,
        formula_fn=lambda cell: per_day_rate,
        is_first=True, trip_id_col=trip_id_col, shipment_lookup=shipment_lookup,
    )

    grp = df.groupby(rider_col)
    delivered_by_rider = grp[delivered_col].sum()
    login_days_by_rider = grp[date_col].nunique()
    riders = sorted(delivered_by_rider.index.tolist())

    # Store label per rider: first non-blank STORE NAME value for that rider,
    # with a 'kappa_' prefix stripped and underscores turned into spaces.
    store_label_by_rider = {}
    for name in riders:
        label = ''
        if store_col:
            vals = df.loc[df[rider_col] == name, store_col].dropna()
            if len(vals):
                s = str(vals.iloc[0]).strip()
                if s.upper().startswith('KAPPA_'):
                    s = s[len('KAPPA_'):]
                label = s.replace('_', ' ').upper()
        store_label_by_rider[name] = label

    ws2 = wb.create_sheet('FINAL PAYOUT')
    headers2 = ['STORE NAME', 'RIDERS NAME', 'ATTENDANCE', 'TOTAL DELIVERED ORDERS',
                'PAYOUT', 'Insurance Deduction', 'FINAL PAYOUT', 'DCPO']
    style_header_row(ws2, headers2)

    for i, name in enumerate(riders):
        r = i + 2
        d = int(delivered_by_rider[name])
        ld = int(login_days_by_rider[name])
        ws2.cell(row=r, column=1, value=store_label_by_rider.get(name, '')).font = BODY_FONT
        ws2.cell(row=r, column=2, value=name).font = BODY_FONT
        ws2.cell(row=r, column=3, value=ld).font = BODY_FONT
        ws2.cell(row=r, column=4, value=d).font = BODY_FONT
        ws2.cell(row=r, column=5, value=f'=C{r}*{per_day_rate}').font = BODY_FONT
        ws2.cell(row=r, column=6, value=insurance).font = BODY_FONT
        ws2.cell(row=r, column=7, value=f'=E{r}-F{r}').font = BODY_FONT
        ws2.cell(row=r, column=8, value=f'=E{r}/D{r}').font = BODY_FONT
        for col in range(1, 9):
            cell = ws2.cell(row=r, column=col)
            cell.alignment = CENTER
            cell.border = BORDER

    total_row = len(riders) + 2
    ws2.cell(row=total_row, column=1, value='TOTAL').font = TOTAL_FONT
    for col_letter in ['D', 'E', 'F', 'G']:
        ws2.cell(row=total_row, column=ord(col_letter) - 64,
                 value=f'=SUM({col_letter}2:{col_letter}{total_row-1})').font = TOTAL_FONT
    ws2.cell(row=total_row, column=8, value=f'=E{total_row}/D{total_row}').font = TOTAL_FONT
    for col in range(1, 9):
        cell = ws2.cell(row=total_row, column=col)
        cell.fill = TOTAL_FILL
        cell.alignment = CENTER
        cell.border = BORDER

    widths2 = [18, 22, 12, 20, 10, 18, 14, 10]
    for j, w in enumerate(widths2, start=1):
        ws2.column_dimensions[get_column_letter(j)].width = w

    attendance_sheet = find_sheet_by_keyword(raw_path, 'ATTEN')
    if attendance_sheet:
        copy_sheet_untouched(raw_path, attendance_sheet, wb, 'ATTENDANCE SHEET')
    bank_sheet = (find_sheet_by_keyword(raw_path, 'AC', 'DETAIL')
                  or find_sheet_by_keyword(raw_path, 'BANK'))
    if bank_sheet:
        copy_sheet_untouched(raw_path, bank_sheet, wb, 'BANK AC DETAILS')

    wb.save(out_path)
    return out_path


def build_flat_per_trip(raw_path, out_path, per_day_rate, shipment_lookup=None):
    """Lambda's scheme: every trip (rider-day row) pays the same flat amount
    (e.g. ₹720), no threshold/escalation, no per-order rate, no incentive
    tiers, and no insurance deduction tab at all. Output: a single PAYOUT
    tab. (Any other adjustments are tracked outside this tool and
    intentionally not modeled here.)"""
    df, date_col, rider_col, delivered_col, store_col = load_sheet1(raw_path)
    trip_id_col = find_col(df.columns, 'TRIP', 'ID')

    wb = Workbook()
    build_trip_sheet(
        wb, 'PAYOUT', df, rider_col, date_col, delivered_col,
        formula_fn=lambda cell: per_day_rate,
        is_first=True, trip_id_col=trip_id_col, shipment_lookup=shipment_lookup,
    )

    wb.save(out_path)
    return out_path


def build_per_order_insured(raw_path, out_path, rate, insurance, shipment_lookup=None):
    """Tau's scheme: plain per-order rate + flat insurance, no incentive
    tiers, no FINAL PAYOUT rollup — just PAYOUT and INSURANCE DEDUCTION tabs.
    Rows where the rider was logged in/out but no trip actually ran that day
    (no TRIP ID) are dropped from the PAYOUT tab entirely; every rider that
    appears anywhere in the raw file still gets an INSURANCE DEDUCTION row.

    NOTE: this store's raw export sometimes comes as *two* files, because a
    different last-mile vendor's riders share the same physical store (e.g.
    an 'Other Vendor' file mixed in with Tau's own). That other vendor's
    data is NOT part of Tau's payout — only the file containing Tau's
    own riders should be uploaded/processed here."""
    df, date_col, rider_col, delivered_col, store_col = load_sheet1(raw_path)
    trip_id_col = find_col(df.columns, 'TRIP', 'ID')

    # Insurance applies to every rider in the raw file, even on days with no
    # trip — compute before dropping no-trip rows.
    all_riders = sorted(df[rider_col].dropna().unique().tolist())

    trip_df = df
    if trip_id_col:
        trip_df = df[df[trip_id_col].notna()].reset_index(drop=True)

    wb = Workbook()
    build_trip_sheet(
        wb, 'PAYOUT', trip_df, rider_col, date_col, delivered_col,
        formula_fn=lambda cell: f'={cell}*{rate}',
        is_first=True, trip_id_col=trip_id_col, shipment_lookup=shipment_lookup,
    )

    ws2 = wb.create_sheet('INSURANCE DEDUCTION')
    headers2 = ['RIDER NAME', 'INSURANCE DEDUCTION']
    style_header_row(ws2, headers2)
    for i, name in enumerate(all_riders):
        r = i + 2
        ws2.cell(row=r, column=1, value=name).font = BODY_FONT
        ws2.cell(row=r, column=2, value=insurance).font = BODY_FONT
        for col in range(1, 3):
            cell = ws2.cell(row=r, column=col)
            cell.alignment = CENTER
            cell.border = BORDER
    total_row2 = len(all_riders) + 2
    ws2.cell(row=total_row2, column=1, value='TOTAL').font = TOTAL_FONT
    ws2.cell(row=total_row2, column=2, value=f'=SUM(B2:B{total_row2-1})').font = TOTAL_FONT
    for col in range(1, 3):
        cell = ws2.cell(row=total_row2, column=col)
        cell.fill = TOTAL_FILL
        cell.alignment = CENTER
        cell.border = BORDER
    for j, w in enumerate([22, 20], start=1):
        ws2.column_dimensions[get_column_letter(j)].width = w

    wb.save(out_path)
    return out_path


def resolve_config(company, store, rate=None, insurance=None, tiers=None, mg_config=None):
    """Look up a store's rate/insurance/tiers/mg-scheme/dcpo-range from
    COMPANY_CONFIG, allowing any of them to be explicitly overridden."""
    company_key = (company or 'ALPHA').strip().upper()
    store_key = (store or '').strip().upper()
    cfg = COMPANY_CONFIG.get(company_key, {}).get(store_key)

    if rate is None:
        if cfg is None:
            raise ValueError(f"Rate not given and '{store}' is not configured under "
                              f"company '{company_key}'; pass --rate explicitly.")
        rate = cfg['rate']
    if insurance is None:
        insurance = cfg['insurance'] if cfg else 0
    if tiers is None:
        tiers = cfg['tiers'] if cfg else [(7, 45, 640), (6, 45, 420)]
    if mg_config is None:
        mg_config = cfg.get('mg') if cfg else None
    dcpo_range = cfg.get('dcpo_range') if cfg else None
    return rate, insurance, tiers, mg_config, dcpo_range


def preview_summary(raw_path, company='ALPHA', store=None, rate=None, insurance=None, tiers=None):
    """Compute the same per-rider stats the final build would produce, without
    writing any file. Used for a validation preview before download."""
    company_key = (company or 'ALPHA').strip().upper()

    if company_key == 'EPSILON GROUP':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('EPSILON GROUP', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'EPSILON GROUP'.")
        ins = insurance if insurance is not None else cfg['insurance']
        m = cfg['mg_per_trip']
        base, threshold, mg_rate = m['base'], m['threshold'], m['rate']

        df, date_col, rider_col, delivered_col, store_col = load_sheet1(raw_path)
        df = df[df[rider_col].notna()].reset_index(drop=True)
        riders = sorted(df[rider_col].dropna().unique().tolist())

        rider_rows = []
        total_delivered = total_payout = total_final = 0
        for name in riders:
            rows = df[df[rider_col] == name]
            d = int(rows[delivered_col].sum())
            ld = rows[date_col].nunique()
            payout = float(sum(mg_payout_value(int(v), base, threshold, mg_rate)
                                for v in rows[delivered_col]))
            final = payout - ins
            rider_rows.append({
                'name': str(name), 'days': int(ld), 'delivered': d,
                'incentive': 0, 'payout': payout, 'total_payout': payout,
                'insurance': ins, 'final_payout': final, 'scheme': 'MG (per trip)',
            })
            total_delivered += d
            total_payout += payout
            total_final += final

        try:
            date_min = str(pd.to_datetime(df[date_col], errors='coerce').min().date())
            date_max = str(pd.to_datetime(df[date_col], errors='coerce').max().date())
        except Exception:
            date_min = date_max = None

        return {
            'rate': None, 'insurance': ins,
            'rider_count': len(rider_rows), 'row_count': len(df),
            'date_min': date_min, 'date_max': date_max,
            'riders': rider_rows,
            'dcpo_range': None,
            'totals': {
                'delivered': total_delivered, 'incentive': 0,
                'payout': total_payout, 'total_payout': total_payout,
                'insurance': ins * len(rider_rows), 'final_payout': total_final,
            },
        }

    if company_key == 'ZETA':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('ZETA', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'ZETA'.")
        r = rate if rate is not None else cfg['rate']
        cr = cfg['cancelled_rate']

        df = pd.read_excel(raw_path, sheet_name=0)
        df.columns = [str(c).strip() for c in df.columns]
        df = filter_absent_rows(df)
        date_col = find_col(df.columns, 'DATE')
        rider_col = find_col(df.columns, 'RIDER', 'NAME')
        delivered_col = find_col(df.columns, 'DELIVERED', 'ORDER')
        cancelled_col = find_col(df.columns, 'CANCELLED', 'ORDER')
        missing = [n for n, c in [('DATE', date_col), ('RIDER NAME', rider_col),
                                   ('TOTAL DELIVERED ORDERS', delivered_col),
                                   ('TOTAL CANCELLED ORDERS', cancelled_col)] if c is None]
        if missing:
            raise MissingColumnError(0, missing, list(df.columns))
        df = normalize_date_col(df, date_col)
        df = df[df[rider_col].notna()].reset_index(drop=True)

        riders = sorted(df[rider_col].dropna().unique().tolist())
        rider_rows = []
        total_delivered = total_payout = 0
        for name in riders:
            rows = df[df[rider_col] == name]
            d = int(rows[delivered_col].sum())
            cxl = int(rows[cancelled_col].sum())
            ld = int(rows[date_col].nunique())
            payout = float(d * r + cxl * cr)
            rider_rows.append({
                'name': str(name), 'days': ld, 'delivered': d,
                'incentive': 0, 'payout': payout, 'total_payout': payout,
                'insurance': 0, 'final_payout': payout, 'scheme': 'Delivered + Cancelled',
            })
            total_delivered += d
            total_payout += payout

        try:
            date_min = str(pd.to_datetime(df[date_col], errors='coerce').min().date())
            date_max = str(pd.to_datetime(df[date_col], errors='coerce').max().date())
        except Exception:
            date_min = date_max = None

        return {
            'rate': r, 'insurance': 0,
            'rider_count': len(rider_rows), 'row_count': len(df),
            'date_min': date_min, 'date_max': date_max,
            'riders': rider_rows,
            'dcpo_range': None,
            'totals': {
                'delivered': total_delivered, 'incentive': 0,
                'payout': total_payout, 'total_payout': total_payout,
                'insurance': 0, 'final_payout': total_payout,
            },
        }

    if company_key == 'KAPPA':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('KAPPA', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'KAPPA'.")
        ins = insurance if insurance is not None else cfg['insurance']
        per_day = cfg['flat_per_day']

        df, date_col, rider_col, delivered_col, store_col = load_sheet1(raw_path)
        grp = df.groupby(rider_col)
        delivered_by_rider = grp[delivered_col].sum()
        login_days_by_rider = grp[date_col].nunique()
        riders = sorted(delivered_by_rider.index.tolist())

        rider_rows = []
        total_delivered = total_payout = total_final = 0
        for name in riders:
            d = int(delivered_by_rider[name])
            ld = int(login_days_by_rider[name])
            payout = float(ld * per_day)
            final = payout - ins
            rider_rows.append({
                'name': str(name), 'days': ld, 'delivered': d,
                'incentive': 0, 'payout': payout, 'total_payout': payout,
                'insurance': ins, 'final_payout': final, 'scheme': 'FLAT PER DAY',
            })
            total_delivered += d
            total_payout += payout
            total_final += final

        try:
            date_min = str(pd.to_datetime(df[date_col], errors='coerce').min().date())
            date_max = str(pd.to_datetime(df[date_col], errors='coerce').max().date())
        except Exception:
            date_min = date_max = None

        return {
            'rate': None, 'insurance': ins,
            'rider_count': len(rider_rows), 'row_count': len(df),
            'date_min': date_min, 'date_max': date_max,
            'riders': rider_rows,
            'dcpo_range': None,
            'totals': {
                'delivered': total_delivered, 'incentive': 0,
                'payout': total_payout, 'total_payout': total_payout,
                'insurance': ins * len(rider_rows), 'final_payout': total_final,
            },
        }

    if company_key == 'LAMBDA':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('LAMBDA', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'LAMBDA'.")
        per_day = cfg['flat_per_day']

        df, date_col, rider_col, delivered_col, store_col = load_sheet1(raw_path)
        grp = df.groupby(rider_col)
        delivered_by_rider = grp[delivered_col].sum()
        login_days_by_rider = grp[date_col].nunique()
        riders = sorted(delivered_by_rider.index.tolist())

        rider_rows = []
        total_delivered = total_payout = 0
        for name in riders:
            d = int(delivered_by_rider[name])
            ld = int(login_days_by_rider[name])
            payout = float(ld * per_day)
            rider_rows.append({
                'name': str(name), 'days': ld, 'delivered': d,
                'incentive': 0, 'payout': payout, 'total_payout': payout,
                'insurance': 0, 'final_payout': payout, 'scheme': 'FLAT PER DAY (no insurance)',
            })
            total_delivered += d
            total_payout += payout

        try:
            date_min = str(pd.to_datetime(df[date_col], errors='coerce').min().date())
            date_max = str(pd.to_datetime(df[date_col], errors='coerce').max().date())
        except Exception:
            date_min = date_max = None

        return {
            'rate': None, 'insurance': 0,
            'rider_count': len(rider_rows), 'row_count': len(df),
            'date_min': date_min, 'date_max': date_max,
            'riders': rider_rows,
            'dcpo_range': None,
            'totals': {
                'delivered': total_delivered, 'incentive': 0,
                'payout': total_payout, 'total_payout': total_payout,
                'insurance': 0, 'final_payout': total_payout,
            },
        }

    if company_key == 'OMEGA':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('OMEGA', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'OMEGA'.")
        ins = insurance if insurance is not None else cfg['insurance']
        flat_rate = cfg['flat_per_day']

        df, date_col, rider_col, delivered_col, store_col = load_sheet1(raw_path)
        df = df[df[rider_col].notna()].reset_index(drop=True)
        grp = df.groupby(rider_col)
        delivered_by_rider = grp[delivered_col].sum()
        login_days_by_rider = grp[date_col].nunique()
        trip_count_by_rider = grp.size()
        riders = sorted(delivered_by_rider.index.tolist())

        rider_rows = []
        total_delivered = total_payout = total_final = 0
        for name in riders:
            d = int(delivered_by_rider[name])
            ld = int(login_days_by_rider[name])
            trips = int(trip_count_by_rider[name])
            payout = float(trips * flat_rate)
            final = payout - ins
            rider_rows.append({
                'name': str(name), 'days': ld, 'delivered': d,
                'incentive': 0, 'payout': payout, 'total_payout': payout,
                'insurance': ins, 'final_payout': final, 'scheme': 'FLAT PER TRIP',
            })
            total_delivered += d
            total_payout += payout
            total_final += final

        try:
            date_min = str(pd.to_datetime(df[date_col], errors='coerce').min().date())
            date_max = str(pd.to_datetime(df[date_col], errors='coerce').max().date())
        except Exception:
            date_min = date_max = None

        return {
            'rate': None, 'insurance': ins,
            'rider_count': len(rider_rows), 'row_count': len(df),
            'date_min': date_min, 'date_max': date_max,
            'riders': rider_rows,
            'dcpo_range': None,
            'totals': {
                'delivered': total_delivered, 'incentive': 0,
                'payout': total_payout, 'total_payout': total_payout,
                'insurance': ins * len(rider_rows), 'final_payout': total_final,
            },
        }

    if company_key == 'TAU':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('TAU', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'TAU'.")
        r = rate if rate is not None else cfg['rate']
        ins = insurance if insurance is not None else cfg['insurance']

        df, date_col, rider_col, delivered_col, store_col = load_sheet1(raw_path)
        trip_id_col = find_col(df.columns, 'TRIP', 'ID')
        all_riders = sorted(df[rider_col].dropna().unique().tolist())
        trip_df = df
        if trip_id_col:
            trip_df = df[df[trip_id_col].notna()].reset_index(drop=True)

        grp = trip_df.groupby(rider_col)
        delivered_by_rider = grp[delivered_col].sum()
        login_days_by_rider = grp[date_col].nunique()

        rider_rows = []
        total_delivered = total_payout = total_final = 0
        for name in all_riders:
            d = int(delivered_by_rider.get(name, 0))
            ld = int(login_days_by_rider.get(name, 0))
            payout = float(d * r)
            final = payout - ins
            rider_rows.append({
                'name': str(name), 'days': ld, 'delivered': d,
                'incentive': 0, 'payout': payout, 'total_payout': payout,
                'insurance': ins, 'final_payout': final, 'scheme': 'Rate + Insurance (no tiers)',
            })
            total_delivered += d
            total_payout += payout
            total_final += final

        try:
            date_min = str(pd.to_datetime(df[date_col], errors='coerce').min().date())
            date_max = str(pd.to_datetime(df[date_col], errors='coerce').max().date())
        except Exception:
            date_min = date_max = None

        return {
            'rate': r, 'insurance': ins,
            'rider_count': len(rider_rows), 'row_count': len(trip_df),
            'date_min': date_min, 'date_max': date_max,
            'riders': rider_rows,
            'dcpo_range': None,
            'totals': {
                'delivered': total_delivered, 'incentive': 0,
                'payout': total_payout, 'total_payout': total_payout,
                'insurance': ins * len(rider_rows), 'final_payout': total_final,
            },
        }

    if company_key == 'THETA':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('THETA', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'THETA'.")
        r = rate if rate is not None else cfg['rate']
        cr = cfg['cancelled_rate']
        t = tiers if tiers is not None else cfg.get('tiers')

        df = pd.read_excel(raw_path, sheet_name=0)
        df.columns = [str(c).strip() for c in df.columns]
        df = filter_absent_rows(df)
        date_col = find_col(df.columns, 'DATE')
        rider_col = find_col(df.columns, 'RIDER', 'NAME')
        delivered_col = find_col(df.columns, 'DELIVERED', 'ORDER')
        cancelled_col = find_col(df.columns, 'CANCELLED', 'ORDER')
        missing = [n for n, c in [('DATE', date_col), ('RIDER NAME', rider_col),
                                   ('TOTAL DELIVERED ORDERS', delivered_col),
                                   ('TOTAL CANCELLED ORDERS', cancelled_col)] if c is None]
        if missing:
            raise MissingColumnError(0, missing, list(df.columns))
        df = normalize_date_col(df, date_col)
        df = df[df[rider_col].notna()].reset_index(drop=True)

        threshold_days = min_tier_days(t) if t else None
        riders = sorted(df[rider_col].dropna().unique().tolist())
        rider_rows = []
        total_delivered = total_payout = total_incentive = 0
        for name in riders:
            rows = df[df[rider_col] == name]
            d = int(rows[delivered_col].sum())
            cxl = int(rows[cancelled_col].sum())
            ld = int(rows[date_col].nunique())
            la = login_amount(ld, d, t) if (t and ld >= threshold_days) else 0
            payout = float(d * r + cxl * cr)
            tp = payout + la
            rider_rows.append({
                'name': str(name), 'days': ld, 'delivered': d,
                'incentive': la, 'payout': payout, 'total_payout': tp,
                'insurance': 0, 'final_payout': tp, 'scheme': 'Delivered + Cancelled + Incentive',
            })
            total_delivered += d
            total_payout += tp
            total_incentive += la

        try:
            date_min = str(pd.to_datetime(df[date_col], errors='coerce').min().date())
            date_max = str(pd.to_datetime(df[date_col], errors='coerce').max().date())
        except Exception:
            date_min = date_max = None

        return {
            'rate': r, 'insurance': 0,
            'rider_count': len(rider_rows), 'row_count': len(df),
            'date_min': date_min, 'date_max': date_max,
            'riders': rider_rows,
            'dcpo_range': None,
            'totals': {
                'delivered': total_delivered, 'incentive': total_incentive,
                'payout': total_payout, 'total_payout': total_payout,
                'insurance': 0, 'final_payout': total_payout,
            },
        }

    if company_key == 'GAMMA':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('GAMMA', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'GAMMA'.")
        if cfg.get('distance_mg'):
            ins = insurance if insurance is not None else cfg['insurance']
            dmg = cfg['distance_mg']
            return preview_distance_mg(raw_path, ins, dmg['mg_flat'], dmg['per_km_rate'],
                                        distance_col_override=cfg.get('distance_col_override'),
                                        restrict_vendor=cfg.get('restrict_vendor'))
        # else: BETA-style model (per-order rate + incentive tiers + insurance) —
        # falls through to the shared ALPHA/BETA preview logic below.

    if company_key == 'SIGMA':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('SIGMA', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'SIGMA'.")
        ins = insurance if insurance is not None else cfg['insurance']
        dmg = cfg['distance_mg']
        return preview_distance_mg(raw_path, ins, dmg['mg_flat'], dmg['per_km_rate'])

    if company_key == 'DELTA':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('DELTA', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'DELTA'.")
        r = rate if rate is not None else cfg['rate']
        mg_cfg = cfg.get('mg')
        t = tiers if tiers is not None else cfg.get('tiers')

        mg_sheet_name = find_mg_sheet_name(raw_path) if mg_cfg else None
        po_sheet_name = find_non_mg_sheet_name(raw_path, mg_sheet_name) if mg_sheet_name else 0

        df, date_col, rider_col, delivered_col, _ = load_sheet(raw_path, po_sheet_name)
        grp = df.groupby(rider_col)
        delivered_by_rider = grp[delivered_col].sum()
        days_by_rider = grp[date_col].nunique()
        po_riders = sorted(delivered_by_rider.index.tolist())
        threshold_days = min_tier_days(t) if t else None

        rider_rows = []
        total_delivered = total_payout = total_incentive = 0
        for name in po_riders:
            d = int(delivered_by_rider[name])
            ld = int(days_by_rider[name])
            la = login_amount(ld, d, t) if (t and ld >= threshold_days) else 0
            payout = float(r * d)
            tp = payout + la
            rider_rows.append({
                'name': str(name), 'days': ld, 'delivered': d,
                'incentive': la, 'payout': payout, 'total_payout': tp,
                'insurance': 0, 'final_payout': tp, 'scheme': 'PER ORDER',
            })
            total_delivered += d
            total_payout += tp
            total_incentive += la

        date_sources = [df[date_col]]

        if mg_sheet_name:
            df_mg, mg_date_col, mg_rider_col, mg_delivered_col, _ = load_sheet(raw_path, mg_sheet_name)
            grp_mg = df_mg.groupby(mg_rider_col)
            mg_days = grp_mg[mg_date_col].nunique()
            mg_riders = sorted(df_mg[mg_rider_col].dropna().unique().tolist())
            base, threshold, mg_rate = mg_cfg['base'], mg_cfg['threshold'], mg_cfg['rate']
            for name in mg_riders:
                rows = df_mg[df_mg[mg_rider_col] == name]
                d = int(rows[mg_delivered_col].sum())
                ld = int(mg_days[name])
                payout = float(sum(mg_payout_value(int(v), base, threshold, mg_rate)
                                    for v in rows[mg_delivered_col]))
                rider_rows.append({
                    'name': str(name), 'days': ld, 'delivered': d,
                    'incentive': 0, 'payout': payout, 'total_payout': payout,
                    'insurance': 0, 'final_payout': payout, 'scheme': 'MG',
                })
                total_delivered += d
                total_payout += payout
            date_sources.append(df_mg[mg_date_col])

        rider_rows.sort(key=lambda x: x['name'])
        try:
            all_dates = pd.concat(date_sources)
            date_min = str(pd.to_datetime(all_dates, errors='coerce').min().date())
            date_max = str(pd.to_datetime(all_dates, errors='coerce').max().date())
        except Exception:
            date_min = date_max = None

        return {
            'rate': r, 'insurance': 0,
            'rider_count': len(rider_rows), 'row_count': len(df) + (len(df_mg) if mg_sheet_name else 0),
            'date_min': date_min, 'date_max': date_max,
            'riders': rider_rows,
            'dcpo_range': None,
            'totals': {
                'delivered': total_delivered, 'incentive': total_incentive,
                'payout': total_payout, 'total_payout': total_payout,
                'insurance': 0, 'final_payout': total_payout,
            },
        }

    rate, insurance, tiers, mg_config, dcpo_range = resolve_config(company, store, rate, insurance, tiers)

    mg_sheet_name = find_mg_sheet_name(raw_path) if mg_config else None
    po_sheet_name = find_non_mg_sheet_name(raw_path, mg_sheet_name) if mg_sheet_name else 0

    df, date_col, rider_col, delivered_col, store_col = load_sheet(raw_path, po_sheet_name)
    grp = df.groupby(rider_col)
    delivered_by_rider = grp[delivered_col].sum()
    login_days_by_rider = grp[date_col].nunique()
    po_riders = sorted(delivered_by_rider.index.tolist())
    threshold_days = min_tier_days(tiers)

    def dcpo_flag(dcpo_value):
        if not dcpo_range or dcpo_value is None:
            return None
        lo, hi = dcpo_range
        if dcpo_value < lo:
            return 'low'
        if dcpo_value > hi:
            return 'high'
        return None

    rider_rows = []
    total_delivered = total_incentive = total_payout = total_final = 0

    for name in po_riders:
        d = int(delivered_by_rider[name])
        ld = int(login_days_by_rider[name])
        la = login_amount(ld, d, tiers) if ld >= threshold_days else 0
        payout = rate * d
        tp = payout + la
        final = tp - insurance
        dcpo = round(tp / d, 2) if d else None
        rider_rows.append({
            'name': str(name), 'days': ld, 'delivered': d,
            'incentive': la, 'payout': payout, 'total_payout': tp,
            'insurance': insurance, 'final_payout': final, 'scheme': 'PER ORDER',
            'dcpo': dcpo, 'dcpo_flag': dcpo_flag(dcpo),
        })
        total_delivered += d
        total_incentive += la
        total_payout += tp
        total_final += final

    date_sources = [df[date_col]]

    if mg_sheet_name:
        df_mg, mg_date_col, mg_rider_col, mg_delivered_col, _ = load_sheet(raw_path, mg_sheet_name)
        grp_mg = df_mg.groupby(mg_rider_col)
        mg_days = grp_mg[mg_date_col].nunique()
        mg_riders = sorted(df_mg[mg_rider_col].dropna().unique().tolist())
        base, threshold, mg_rate = mg_config['base'], mg_config['threshold'], mg_config['rate']

        for name in mg_riders:
            rows = df_mg[df_mg[mg_rider_col] == name]
            d = int(rows[mg_delivered_col].sum())
            ld = int(mg_days[name])
            payout = sum(mg_payout_value(int(v), base, threshold, mg_rate)
                         for v in rows[mg_delivered_col])
            tp = payout  # MG riders are not eligible for the per-order INCENTIVE tab
            final = tp - insurance
            dcpo = round(tp / d, 2) if d else None
            rider_rows.append({
                'name': str(name), 'days': ld, 'delivered': d,
                'incentive': 0, 'payout': payout, 'total_payout': tp,
                'insurance': insurance, 'final_payout': final, 'scheme': 'MG',
                'dcpo': dcpo, 'dcpo_flag': dcpo_flag(dcpo),
            })
            total_delivered += d
            total_payout += tp
            total_final += final
        date_sources.append(df_mg[mg_date_col])

    rider_rows.sort(key=lambda r: r['name'])

    try:
        all_dates = pd.concat(date_sources)
        date_min = str(pd.to_datetime(all_dates, errors='coerce', dayfirst=False).min().date())
        date_max = str(pd.to_datetime(all_dates, errors='coerce', dayfirst=False).max().date())
    except Exception:
        date_min = date_max = None

    return {
        'rate': rate, 'insurance': insurance,
        'rider_count': len(rider_rows),
        'row_count': len(df) + (len(df_mg) if mg_sheet_name else 0),
        'date_min': date_min, 'date_max': date_max,
        'riders': rider_rows,
        'dcpo_range': list(dcpo_range) if dcpo_range else None,
        'totals': {
            'delivered': total_delivered, 'incentive': total_incentive,
            'payout': total_payout, 'total_payout': total_payout,
            'insurance': insurance * len(rider_rows), 'final_payout': total_final,
        },
    }


def build(raw_path, out_path, company='ALPHA', rate=None, insurance=None, store=None, tiers=None, shipment_path=None):
    company_key = (company or 'ALPHA').strip().upper()
    shipment_lookup = load_shipment_lookup(shipment_path) if shipment_path else None

    if company_key == 'EPSILON GROUP':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('EPSILON GROUP', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'EPSILON GROUP'.")
        ins = insurance if insurance is not None else cfg['insurance']
        m = cfg['mg_per_trip']
        return build_mg_per_trip(raw_path, out_path, ins, m['base'], m['threshold'], m['rate'],
                            shipment_lookup=shipment_lookup)

    if company_key == 'ZETA':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('ZETA', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'ZETA'.")
        r = rate if rate is not None else cfg['rate']
        cr = cfg['cancelled_rate']
        return build_delivered_plus_cancelled(raw_path, out_path, r, cr, shipment_lookup=shipment_lookup)

    if company_key == 'THETA':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('THETA', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'THETA'.")
        r = rate if rate is not None else cfg['rate']
        cr = cfg['cancelled_rate']
        t = tiers if tiers is not None else cfg.get('tiers')
        return build_delivered_plus_cancelled_tiered(raw_path, out_path, r, cr, t, shipment_lookup=shipment_lookup)

    if company_key == 'KAPPA':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('KAPPA', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'KAPPA'.")
        ins = insurance if insurance is not None else cfg['insurance']
        return build_flat_per_day(raw_path, out_path, ins, cfg['flat_per_day'],
                             shipment_lookup=shipment_lookup)

    if company_key == 'LAMBDA':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('LAMBDA', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'LAMBDA'.")
        return build_flat_per_trip(raw_path, out_path, cfg['flat_per_day'], shipment_lookup=shipment_lookup)

    if company_key == 'OMEGA':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('OMEGA', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'OMEGA'.")
        ins = insurance if insurance is not None else cfg['insurance']
        return build_flat_per_trip_insured(raw_path, out_path, ins, cfg['flat_per_day'], shipment_lookup=shipment_lookup)

    if company_key == 'TAU':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('TAU', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'TAU'.")
        r = rate if rate is not None else cfg['rate']
        ins = insurance if insurance is not None else cfg['insurance']
        return build_per_order_insured(raw_path, out_path, r, ins, shipment_lookup=shipment_lookup)

    if company_key == 'GAMMA':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('GAMMA', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'GAMMA'.")
        if cfg.get('distance_mg'):
            ins = insurance if insurance is not None else cfg['insurance']
            dmg = cfg['distance_mg']
            return build_distance_mg(raw_path, out_path, ins, dmg['mg_flat'], dmg['per_km_rate'],
                                   shipment_lookup=shipment_lookup,
                                   distance_col_override=cfg.get('distance_col_override'),
                                   restrict_vendor=cfg.get('restrict_vendor'))
        # BETA-style model: per-order rate + incentive tiers + insurance
        r = rate if rate is not None else cfg['rate']
        ins = insurance if insurance is not None else cfg['insurance']
        t = tiers if tiers is not None else cfg.get('tiers')
        return build_tiered_split(raw_path, out_path, r, ins, t, mg_config=cfg.get('mg'),
                           shipment_lookup=shipment_lookup, trip_tab_title=cfg.get('trip_tab_title'))

    if company_key == 'SIGMA':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('SIGMA', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'SIGMA'.")
        ins = insurance if insurance is not None else cfg['insurance']
        dmg = cfg['distance_mg']
        return build_distance_mg(raw_path, out_path, ins, dmg['mg_flat'], dmg['per_km_rate'],
                               shipment_lookup=shipment_lookup)

    if company_key == 'DELTA':
        store_key = (store or '').strip().upper()
        cfg = COMPANY_CONFIG.get('DELTA', {}).get(store_key)
        if cfg is None:
            raise ValueError(f"'{store}' is not configured under company 'DELTA'.")
        r = rate if rate is not None else cfg['rate']
        t = tiers if tiers is not None else cfg.get('tiers')
        return build_mg_and_tiers(raw_path, out_path, r, mg_config=cfg.get('mg'), tiers=t,
                              shipment_lookup=shipment_lookup)

    rate, insurance, tiers, mg_config, dcpo_range = resolve_config(company, store, rate, insurance, tiers)
    if company_key == 'BETA':
        return build_tiered_split(raw_path, out_path, rate, insurance, tiers, mg_config=mg_config,
                           shipment_lookup=shipment_lookup)
    else:
        return build_tiered_summary(raw_path, out_path, rate, insurance, tiers, shipment_lookup=shipment_lookup)


if __name__ == '__main__':
    p = argparse.ArgumentParser()
    p.add_argument('raw_file')
    p.add_argument('out_file')
    p.add_argument('--company', default='ALPHA', help='Vendor name, e.g. ALPHA, BETA')
    p.add_argument('--rate', type=float, default=None)
    p.add_argument('--insurance', type=float, default=None)
    p.add_argument('--store', default=None, help='Store name, e.g. "NORTH HUB", "EAST HUB", RIVERSIDE')
    p.add_argument('--tiers', default=None,
                    help='Override tiers, format: "days:min_orders:amount,..."')
    p.add_argument('--shipment_file', default=None,
                    help='Optional shipment/tracking file to cross-check delivered orders per trip ID')
    args = p.parse_args()
    t = parse_tiers(args.tiers) if args.tiers else None
    build(args.raw_file, args.out_file, company=args.company, rate=args.rate,
          insurance=args.insurance, store=args.store, tiers=t, shipment_path=args.shipment_file)
    print('Saved to', args.out_file)
