#!/usr/bin/env python3
import mysql.connector
from xlsx_utils import safe_save, load_template
from openpyxl.styles import Font, PatternFill, Border, Alignment
from openpyxl.styles.numbers import is_date_format as _is_date_format
from pathlib import Path
from datetime import datetime
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config


def sanitize_field_name(field_name):
    if not field_name:
        return None
    field = field_name.strip().lower()
    field = field.replace(' ', '_').replace('-', '_')
    import re
    field = re.sub(r'[^\w]', '', field)
    return field


def strip_numeric_suffix(name):
    import re
    return re.sub(r'_\d+$', '', name)


CALCULATED_FIELDS = {
    'month(entry_datesubmitted)': '__month__',
    'year(entry_datesubmitted)': '__year__',
}


def read_match_mapping(wb, existing_columns):
    if 'Match' not in wb.sheetnames:
        raise ValueError("El template no tiene un sheet llamado 'Match'")

    ws = wb['Match']
    mapping = {}
    defaults = {}
    col_lower = {k: v for k, v in existing_columns.items()}

    for row in range(2, ws.max_row + 1):
        report_col = ws.cell(row, 1).value
        db_field = ws.cell(row, 3).value
        default_val = ws.cell(row, 4).value

        if not report_col:
            break
        report_col = str(report_col).strip()

        if isinstance(default_val, str) and default_val.strip():
            defaults[report_col] = default_val.strip()

        if db_field:
            db_field_raw = str(db_field).strip()
            db_field_lower = db_field_raw.lower().replace(' ', '')

            if db_field_lower in CALCULATED_FIELDS:
                mapping[report_col] = CALCULATED_FIELDS[db_field_lower]
                continue

            sanitized_db = sanitize_field_name(db_field_raw)

            found = col_lower.get(sanitized_db)
            if not found:
                for lk, lv in col_lower.items():
                    if sanitized_db == lk.replace(' ', '_').replace('-', '_'):
                        found = lv
                        break
            if not found:
                base = strip_numeric_suffix(sanitized_db)
                for lk, lv in col_lower.items():
                    if base == strip_numeric_suffix(lk.replace(' ', '_').replace('-', '_')):
                        found = lv
                        break
            if not found:
                clean_db = sanitized_db.replace('_', '')
                for lk, lv in col_lower.items():
                    clean_col = lk.replace('_', '').replace(' ', '')
                    if clean_db == clean_col or (len(clean_db) > 3 and clean_db in clean_col):
                        found = lv
                        break

            if found:
                mapping[report_col] = found

    return mapping, defaults


def get_table_columns(cursor, table_name):
    cursor.execute(f"SHOW COLUMNS FROM `{table_name}`")
    return {col[0].lower(): col[0] for col in cursor.fetchall()}


def get_data_from_db(db_fields, table_name):
    conn = mysql.connector.connect(**config.MYSQL_CONFIG)
    cursor = conn.cursor()

    real_fields = [f for f in db_fields if f and not f.startswith('__')]
    if not real_fields:
        fields_str = '*'
    else:
        fields_str = ', '.join([f'`{f}`' for f in real_fields])

    query = f"""
        SELECT {fields_str},
               MONTH(entry_datesubmitted) as report_month,
               YEAR(entry_datesubmitted) as report_year
        FROM `{table_name}`
        ORDER BY id DESC
    """
    cursor.execute(query)
    rows = cursor.fetchall()
    cursor.close()
    conn.close()

    # Append calculated field values to each row
    extended_fields = real_fields + ['__month__', '__year__']
    return rows, extended_fields


def generate_excel():
    table_name = 'calls_consolidate_report'
    template_path = config.TEMPLATES['calls_consolidate']
    output_path = config.OUTPUT_FILES['calls_consolidate']

    print(f"Template: {template_path}")
    wb = load_template(template_path)

    conn = mysql.connector.connect(**config.MYSQL_CONFIG)
    cursor = conn.cursor()
    existing_columns = get_table_columns(cursor, table_name)
    cursor.close()
    conn.close()
    print(f"Columnas en tabla: {len(existing_columns)}")

    match_mapping, defaults = read_match_mapping(wb, existing_columns)
    print(f"Mapeo Match: {len(match_mapping)} columnas, defaults: {len(defaults)}")

    db_fields = list(match_mapping.values())
    rows, valid_fields = get_data_from_db(db_fields, table_name)
    print(f"Filas obtenidas: {len(rows)}")

    if 'Template' not in wb.sheetnames:
        raise ValueError("El template no tiene un sheet llamado 'Template'")

    ws = wb['Template']

    # Build mapping: template column index -> db_field
    col_to_db = {}
    col_has_default = {}
    max_col = ws.max_column if ws.max_column > 50 else 100
    for col in range(1, max_col + 1):
        header_value = ws.cell(1, col).value
        if header_value:
            header_str = str(header_value).strip()
            if header_str in match_mapping:
                col_to_db[col] = match_mapping[header_str]
            if header_str in defaults:
                col_has_default[col] = defaults[header_str]

    print(f"Columnas en template con mapping: {len(col_to_db)}")

    data_start_row = 2
    # Clear all existing data rows (need enough rows for all data)
    for row in range(data_start_row, data_start_row + len(rows) + 10):
        for col in range(1, max_col + 1):
            cell = ws.cell(row, col)
            cell.value = None

    for row_idx, row_data in enumerate(rows, start=data_start_row):
        for excel_col, db_field in col_to_db.items():
            if db_field in valid_fields:
                field_pos = valid_fields.index(db_field)
                value = row_data[field_pos]
                cell = ws.cell(row_idx, excel_col)
                if value is None or value == '':
                    if excel_col in col_has_default:
                        cell.value = col_has_default[excel_col]
                    else:
                        cell.value = None
                elif isinstance(value, datetime):
                    cell.value = value
                elif isinstance(value, (int, float)):
                    if _is_date_format(cell.number_format):
                        cell.value = str(value)
                    else:
                        cell.value = value
                else:
                    cell.value = str(value) if value else None

    clean_font = Font()
    clean_fill = PatternFill(fill_type=None)
    clean_border = Border()
    clean_alignment = Alignment(horizontal='left', vertical='center')

    for row in range(data_start_row, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row, col)
            if cell.value is not None:
                cell.font = clean_font
                cell.fill = clean_fill
                cell.border = clean_border
                cell.alignment = clean_alignment

    for sheet_name in wb.sheetnames:
        if sheet_name != 'Template':
            wb.remove(wb[sheet_name])

    safe_save(wb, output_path)
    print(f"Excel generado: {output_path}")
    return output_path


if __name__ == '__main__':
    generate_excel()
