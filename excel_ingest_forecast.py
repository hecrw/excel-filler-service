"""
excel_ingest_forecast.py
Reads the Procapita n8n working Excel, extracts all indicator data,
forecasts empty future-year (e.g. 2027f) columns using Holt's Damped Trend,
upserts all rows to Supabase indicator_values, and writes forecasted values
back into the Excel file.

Usage:
    python3 excel_ingest_forecast.py <path_to_xlsx> [output_path]
"""

import sys, re, math, json, requests, openpyxl

# ── Config ────────────────────────────────────────────────────────────────────
SB_URL = 'https://girovxkebvmaiwficcbe.supabase.co'
SB_KEY = ('eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imdpcm'
          '92eGtlYnZtYWl3ZmljY2JlIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3NzcxODk2Myw'
          'iZXhwIjoyMDkzMjk0OTYzfQ.oIdP-ZJ_db4nXF7rKc6JlNIR7VOi02-Icwv6STxgQ3E')
SB_HDR = {
    'Authorization': f'Bearer {SB_KEY}',
    'apikey': SB_KEY,
    'Content-Type': 'application/json',
    'Prefer': 'resolution=merge-duplicates,return=minimal'
}
BATCH_SIZE = 200

SHEET_CC = {
    'GCC': 'GCC', 'Kuwait': 'KW', 'KSA': 'SA',
    'UAE': 'AE', 'Qatar': 'QA', 'Oman': 'OM', 'Bahrain': 'BH'
}

# Manual overrides for names that don't match catalog exactly
MANUAL_KEY_MAP = {
    'retention strategies implementation': 'talent_mobility__retention_strategy_implementation',
    'retention strategy implementation': 'talent_mobility__retention_strategy_implementation',
    'succession plan rating': 'talent_mobility__succession_planning_readiness',
    'succession planning readiness': 'talent_mobility__succession_planning_readiness',
    'organizations planning to provide ltips': 'rewards__organizations_planning_to_provide_annual_bonus',
    'organizations planning to  provide annual bonus': 'rewards__organizations_planning_to_provide_annual_bonus',
    'expected business challenges': 'business__expected_business_chalenges',
    "job seekers' nationality": 'talent_mobility__job_seekers_nationality',
    "job seekers' gender": 'talent_mobility__job_seekers_gender',
    "job seekers' generation": 'talent_mobility__job_seekers_generation',
    "job seekers' qualification": 'talent_mobility__job_seekers_qualification',
    "job seekers' job level": 'talent_mobility__job_seekers_job_level',
    "job seekers job level": 'talent_mobility__job_seekers_job_level',
    'average employee growth rate': 'talent_mobility__average_employee_growth_rate',
    'average expected employee growth rate': 'talent_mobility__average_expected_employee_growth_rate',
    'employees passing probation period': 'talent_mobility__employees_passing_probation_period',
    'organizations implementing outsourcing': 'talent_mobility__organizations_implementing_outsourcing',
    'functions outsourced': 'talent_mobility__functions_outsourced',
    'industries in high or lower demand': 'talent_mobility__industries_in_high_or_lower_demand',
    'employee exit reasons': 'talent_mobility__employee_exit_reasons',
    'employee engagement score': 'talent_mobility__employee_engagement_score',
    'preferred destinations to work': 'talent_mobility__preferred_destinations_to_work',
    'percentage of increments provided to employees': 'rewards__percentage_of_increments_provided_to_employees',
    'average total compensation per employee': 'rewards__average_total_compensation_per_employee',
    'compensation types motivating talent': 'rewards__compensation_types_motivating_talent',
    'talent salary negotiation attempts': 'rewards__talent_salary_negotiation_attempts',
    'fixed salary vs revenue': 'hr_financials__fixed_salary_vs_revenue',
    'fixed salary vs profit': 'hr_financials__fixed_salary_vs_profit',
    'bonus vs revenue': 'hr_financials__bonus_vs_revenue',
    'bonus vs profit': 'hr_financials__bonus_vs_profit',
    'total compensation vs opex': 'hr_financials__total_compensation_vs_opex',
    'total compensation vs ebitda': 'hr_financials__total_compensation_vs_ebitda',
    'total compensation vs asset': 'hr_financials__total_compensation_vs_asset',
    'average number of board members': 'bod__average_number_of_board_members',
    'average bod remuneration per board member in usd thousands': 'bod__average_bod_remuneration_per_board_member_in_usd_thousands',
    'average net profit in usd thousands': 'bod__average_net_profit_in_usd_thousands',
    'average total remuneration in usd thousands': 'bod__average_total_remuneration_in_usd_thousands',
    'sustainability reporting': 'esg__sustainability_reporting',
    'female representation in the workforce': 'esg__female_representation_in_the_workforce',
    'female representation in the bod': 'esg__female_representation_in_the_bod',
    'average hours of training provided to employees': 'esg__average_hours_of_training_provided_to_employees',
    # Nationality sub-rows that appear as indicator titles in some sheets
    'jordan': 'talent_mobility__job_seekers_nationality',
}


def slugify(s):
    return re.sub(r'[^a-z0-9]+', '_', s.lower().strip()).strip('_')


def get_catalog_key(label, catalog_by_slug):
    """Try to resolve an indicator label to a catalog key."""
    clean = label.strip().lower()
    # Manual override first
    if clean in MANUAL_KEY_MAP:
        return MANUAL_KEY_MAP[clean]
    # Slug match
    slug = slugify(label)
    if slug in catalog_by_slug:
        return catalog_by_slug[slug]
    # Try without trailing spaces / punctuation
    slug2 = slugify(label.rstrip(' .,;'))
    if slug2 in catalog_by_slug:
        return catalog_by_slug[slug2]
    return None


# ── Holt's Damped Trend ───────────────────────────────────────────────────────
def holts_damped_forecast(values, steps=1, alpha=0.3, beta=0.1, phi=0.98):
    if len(values) == 0:
        return [None] * steps
    if len(values) == 1:
        return [values[0]] * steps
    l = values[0]
    b = values[1] - values[0]
    for v in values[1:]:
        l_prev, b_prev = l, b
        l = alpha * v + (1 - alpha) * (l_prev + phi * b_prev)
        b = beta * (l - l_prev) + (1 - beta) * phi * b_prev
    forecasts = []
    for h in range(1, steps + 1):
        damped_trend = sum(phi**j for j in range(1, h + 1))
        forecasts.append(l + damped_trend * b)
    return forecasts


# ── Parse a single country sheet ─────────────────────────────────────────────
def parse_sheet(ws, country_code, catalog_by_slug):
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 3:
        return []

    header = rows[1]

    # Identify year columns (general, cols 4–12ish)
    general_year_cols = {}  # col_idx → (year_int, is_forecast)
    for i, h in enumerate(header):
        if h is None:
            continue
        h_str = str(h).strip()
        m = re.match(r'^(\d{4})([a-zA-Z]*)$', h_str)
        if m:
            yr = int(m.group(1))
            is_fc = bool(m.group(2))
            if 2010 <= yr <= 2035:
                general_year_cols[i] = (yr, is_fc)

    # Identify industry columns (start at col 12+, spaced every 4)
    industry_cols = {}  # col_idx → industry_name
    for i, h in enumerate(header):
        if h is not None and isinstance(h, str) and i >= 12:
            if not re.match(r'^\d{4}', str(h).strip()):
                industry_cols[i] = h.strip()

    # Last 3 measured years for industry sub-col mapping
    measured_year_cols = sorted(
        [(col, yr) for col, (yr, is_fc) in general_year_cols.items() if not is_fc],
        key=lambda x: x[1]
    )
    last3 = measured_year_cols[-3:] if len(measured_year_cols) >= 3 else measured_year_cols

    industry_sub_col_map = {}  # col_idx → (industry_name, year_int)
    for ind_start in sorted(industry_cols.keys()):
        ind_name = industry_cols[ind_start]
        for offset, (_, yr) in enumerate(last3):
            col = ind_start + offset
            industry_sub_col_map[col] = (ind_name, yr)

    # Walk data rows
    records = []
    current_section = None
    current_indicator_num = None
    current_indicator_label = None
    current_catalog_key = None

    for row_idx in range(2, len(rows)):
        row = rows[row_idx]

        # Section label is in col 0
        if row[0] is not None:
            current_section = str(row[0]).strip()

        ind_num = row[2]
        if ind_num is not None:
            try:
                current_indicator_num = int(ind_num)
                raw_label = str(row[3]).strip() if row[3] else f'Indicator_{ind_num}'
                current_indicator_label = raw_label
                current_catalog_key = get_catalog_key(raw_label, catalog_by_slug)
            except (ValueError, TypeError):
                pass

        sub_row_raw = row[3]
        if sub_row_raw is None:
            continue
        sub_row = str(sub_row_raw).strip()
        if not sub_row:
            continue

        # Skip the indicator title row itself (same as label)
        if sub_row == current_indicator_label:
            continue

        # Resolve catalog key — use current indicator's key
        cat_key = current_catalog_key
        if cat_key is None:
            # Try matching the sub_row itself as an indicator
            cat_key = get_catalog_key(sub_row, catalog_by_slug)
        if cat_key is None:
            continue  # Can't map — skip

        # General year values
        year_data = {}
        for col_idx, (yr, is_fc) in general_year_cols.items():
            val = row[col_idx] if col_idx < len(row) else None
            if val is not None:
                try:
                    year_data[yr] = float(val)
                except (ValueError, TypeError):
                    year_data[yr] = None
            else:
                year_data[yr] = None

        has_value = any(v is not None for v in year_data.values())
        if has_value:
            forecast_years = [yr for col_idx, (yr, is_fc) in general_year_cols.items() if is_fc]
            records.append({
                'indicator_key': cat_key,
                'indicator_label': current_indicator_label,
                'sub_row': sub_row,
                'industry': 'All',
                'country_code': country_code,
                'year_data': year_data,
                'forecast_years': forecast_years,
                'row_idx': row_idx,
                'col_map': {yr: col_idx for col_idx, (yr, is_fc) in general_year_cols.items()},
            })

        # Per-industry values
        ind_data = {}
        for col_idx, (ind_name, yr) in industry_sub_col_map.items():
            val = row[col_idx] if col_idx < len(row) else None
            if ind_name not in ind_data:
                ind_data[ind_name] = {}
            if val is not None:
                try:
                    ind_data[ind_name][yr] = float(val)
                except (ValueError, TypeError):
                    pass

        for ind_name, yr_vals in ind_data.items():
            if yr_vals:
                records.append({
                    'indicator_key': cat_key,
                    'indicator_label': current_indicator_label,
                    'sub_row': sub_row,
                    'industry': ind_name,
                    'country_code': country_code,
                    'year_data': yr_vals,
                    'forecast_years': [],
                    'row_idx': row_idx,
                    'col_map': {},
                })

    return records


# Sub-rows that are metadata counts, not business metrics — skip forecasting
SKIP_FORECAST_SUB_ROWS = {
    'number of years', 'sample size', 'n', 'count', 'total responses'
}

# ── Apply Holt's Damped Trend forecasts ──────────────────────────────────────
def apply_forecasts(records):
    for rec in records:
        if not rec['forecast_years']:
            continue
        if rec['sub_row'].strip().lower() in SKIP_FORECAST_SUB_ROWS:
            continue
        measured = sorted(
            [(yr, v) for yr, v in rec['year_data'].items()
             if yr not in rec['forecast_years'] and v is not None],
            key=lambda x: x[0]
        )
        if not measured:
            continue
        hist_values = [v for _, v in measured]
        last_measured_year = measured[-1][0]

        for fc_yr in sorted(rec['forecast_years']):
            if rec['year_data'].get(fc_yr) is not None:
                continue
            steps = fc_yr - last_measured_year
            if steps <= 0:
                continue
            forecasted = holts_damped_forecast(hist_values, steps=steps)
            fc_val = forecasted[-1] if forecasted else None
            if fc_val is not None:
                if all(0 <= v <= 1 for v in hist_values):
                    # Percentage indicator: cap at [2%, 98%]
                    fc_val = max(0.02, min(0.98, fc_val))
                elif all(v >= 0 for v in hist_values):
                    # Count/duration indicator: must stay non-negative
                    fc_val = max(0.0, fc_val)
                rec['year_data'][fc_yr] = round(fc_val, 6)
                rec['is_forecast'] = True


# ── Build Supabase rows ───────────────────────────────────────────────────────
def build_sb_rows(records):
    sb_rows = []
    for rec in records:
        for yr, val in rec['year_data'].items():
            if val is None:
                continue
            is_fc = yr in rec.get('forecast_years', [])
            sb_rows.append({
                'indicator_key': rec['indicator_key'],
                'country_code': rec['country_code'],
                'industry': rec['industry'],
                'sub_row': rec['sub_row'],
                'year': yr,
                'value': round(float(val), 6),
                'value_type': 'forecast' if is_fc else 'measured',
                'source': 'Procapita Insights Excel',
                'citation': None,
                'confidence': 0.85 if is_fc else None,
            })
    return sb_rows


# ── Upsert to Supabase ────────────────────────────────────────────────────────
def upsert_to_supabase(sb_rows):
    inserted = 0
    errors = []
    for i in range(0, len(sb_rows), BATCH_SIZE):
        batch = sb_rows[i:i + BATCH_SIZE]
        try:
            r = requests.post(
                f'{SB_URL}/rest/v1/indicator_values'
                f'?on_conflict=indicator_key,country_code,industry,sub_row,year',
                headers=SB_HDR,
                data=json.dumps(batch)
            )
            if r.status_code not in (200, 201):
                errors.append({'batch_start': i, 'status': r.status_code, 'body': r.text[:300]})
            else:
                inserted += len(batch)
        except Exception as e:
            errors.append({'batch_start': i, 'error': str(e)})
    return inserted, errors


# ── Write forecasts back to Excel ─────────────────────────────────────────────
def write_forecasts_to_excel(wb, sheet_records):
    for sheet_name, records in sheet_records.items():
        ws = wb[sheet_name]
        for rec in records:
            if not rec.get('forecast_years') or not rec.get('col_map'):
                continue
            for fc_yr in rec['forecast_years']:
                val = rec['year_data'].get(fc_yr)
                if val is None:
                    continue
                col_idx = rec['col_map'].get(fc_yr)
                if col_idx is None:
                    continue
                cell = ws.cell(row=rec['row_idx'] + 1, column=col_idx + 1)
                cell.value = val


# ── Main ──────────────────────────────────────────────────────────────────────
def main(xlsx_path, output_path=None):
    if output_path is None:
        output_path = xlsx_path.replace('.xlsx', '_forecasted.xlsx')

    print(f'Loading catalog from Supabase...')
    r = requests.get(f'{SB_URL}/rest/v1/indicator_catalog',
                     headers={'Authorization': f'Bearer {SB_KEY}', 'apikey': SB_KEY},
                     params={'select': 'indicator_key,name', 'limit': '500'})
    catalog = r.json()
    catalog_by_slug = {slugify(row['name']): row['indicator_key'] for row in catalog}
    print(f'  {len(catalog)} catalog entries loaded')

    print(f'Loading: {xlsx_path}')
    wb = openpyxl.load_workbook(xlsx_path)

    all_sb_rows = []
    sheet_records = {}
    stats = {}
    skipped_keys = set()

    for sheet_name, country_code in SHEET_CC.items():
        if sheet_name not in wb.sheetnames:
            print(f'  Skipping {sheet_name} (not in workbook)')
            continue

        ws = wb[sheet_name]
        records = parse_sheet(ws, country_code, catalog_by_slug)
        apply_forecasts(records)
        sb_rows = build_sb_rows(records)
        all_sb_rows.extend(sb_rows)
        sheet_records[sheet_name] = records

        fc_count = sum(1 for r in records if r.get('forecast_years') and
                       any(r['year_data'].get(y) is not None for y in r['forecast_years']))
        stats[sheet_name] = {
            'records': len(records),
            'sb_rows': len(sb_rows),
            'forecasted_rows': fc_count,
        }
        print(f'  {sheet_name} ({country_code}): {len(records)} records, '
              f'{len(sb_rows)} SB rows, {fc_count} forecasted')

    print(f'\nTotal Supabase rows to upsert: {len(all_sb_rows)}')

    inserted, errors = upsert_to_supabase(all_sb_rows)
    print(f'Upserted: {inserted}, Errors: {len(errors)}')
    if errors:
        for e in errors[:5]:
            print(f'  Error: {e}')

    write_forecasts_to_excel(wb, sheet_records)
    wb.save(output_path)
    print(f'\nSaved forecasted Excel to: {output_path}')

    return {
        'inserted': inserted,
        'errors': errors,
        'stats': stats,
        'output_path': output_path,
    }


if __name__ == '__main__':
    path = sys.argv[1] if len(sys.argv) > 1 else \
        'filled-2026-06-28T08-47-01-n8nworking.xlsx'
    out = sys.argv[2] if len(sys.argv) > 2 else None
    result = main(path, out)
    print(json.dumps({'inserted': result['inserted'],
                      'errors_count': len(result['errors']),
                      'stats': result['stats']}, indent=2))
