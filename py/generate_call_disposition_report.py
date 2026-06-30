#!/usr/bin/env python3
import mysql.connector
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import sys, os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config

TEMPLATE_PATH = config.TEMPLATES['call_disposition']
OUTPUT_PATH = config.OUTPUT_FILES['call_disposition']
TARGET_SHEET = 'Call_Disposition_Report'


def get_table_columns(cursor, table_name):
    cursor.execute(f"SHOW COLUMNS FROM `{table_name}`")
    return {col[0].lower(): col[0] for col in cursor.fetchall()}


def read_template_info(template_path):
    wb = load_workbook(template_path, read_only=True)

    if 'Match' not in wb.sheetnames:
        raise ValueError("El template no tiene sheet 'Match'")

    ws_match = wb['Match']
    match = {}
    for row in range(2, ws_match.max_row + 1):
        report_col = ws_match.cell(row, 1).value
        db_field = ws_match.cell(row, 2).value
        if report_col and db_field:
            match[str(report_col).strip()] = str(db_field).strip()

    if TARGET_SHEET not in wb.sheetnames:
        raise ValueError(f"El template no tiene sheet '{TARGET_SHEET}'")

    ws = wb[TARGET_SHEET]
    headers = []
    for col in range(1, 200):
        val = ws.cell(1, col).value
        if val:
            headers.append(str(val).strip())
        else:
            break

    wb.close()
    return match, headers


def get_db_data(cursor, db_fields, table_name):
    valid_fields = [f for f in db_fields if f]
    fields_str = ', '.join([f'`{f}`' for f in valid_fields]) if valid_fields else '*'
    has_entry_date = any('entry_datesubmitted' in f.lower() for f in valid_fields)

    if has_entry_date:
        query = f"""
            SELECT {fields_str},
                   MONTH(STR_TO_DATE(entry_datesubmitted, '%%Y-%%m-%%d %%H:%%i:%%s.%%f')) as report_month,
                   YEAR(STR_TO_DATE(entry_datesubmitted, '%%Y-%%m-%%d %%H:%%i:%%s.%%f')) as report_year
            FROM `{table_name}`
            WHERE entry_datesubmitted IS NOT NULL AND entry_datesubmitted != ''
            ORDER BY id DESC
        """
    else:
        query = f"SELECT {fields_str} FROM `{table_name}` ORDER BY id DESC"

    cursor.execute(query)
    rows = cursor.fetchall()
    field_names = [d[0] for d in cursor.description]
    return rows, field_names


def generate_excel():
    table_name = 'call_disposition_report'
    print(f"Template: {TEMPLATE_PATH}")

    match, template_headers = read_template_info(TEMPLATE_PATH)
    print(f"Match entries: {len(match)}")
    print(f"Columnas en reporte: {len(template_headers)}")

    col_to_db = {}
    needs_month = False
    needs_year = False
    entry_date_field = None

    for report_col_name, db_field in match.items():
        if db_field.lower().startswith('month('):
            needs_month = True
            entry_date_field = db_field[db_field.find('(')+1:db_field.find(')')]
            col_to_db[report_col_name] = ('month', entry_date_field)
        elif db_field.lower().startswith('year('):
            needs_year = True
            entry_date_field = db_field[db_field.find('(')+1:db_field.find(')')]
            col_to_db[report_col_name] = ('year', entry_date_field)
        else:
            col_to_db[report_col_name] = ('field', db_field)

    db_fields_set = set()
    for entry_type, field_name in col_to_db.values():
        if entry_type == 'field' and field_name:
            db_fields_set.add(field_name)

    conn = mysql.connector.connect(**config.MYSQL_CONFIG)
    cursor = conn.cursor()

    if (needs_month or needs_year) and entry_date_field:
        db_fields_set.add(entry_date_field)

    existing_columns = get_table_columns(cursor, table_name)
    valid_fields = []
    for f in db_fields_set:
        if f.lower() in existing_columns:
            valid_fields.append(existing_columns[f.lower()])

    rows, field_names = get_db_data(cursor, valid_fields, table_name)
    cursor.close()
    conn.close()
    print(f"Total registros: {len(rows)}")

    field_index = {name: idx for idx, name in enumerate(field_names)}

    output_dir = os.path.dirname(OUTPUT_PATH)
    os.makedirs(output_dir, exist_ok=True)

    with pd.ExcelWriter(OUTPUT_PATH, engine='xlsxwriter') as writer:
        df = pd.DataFrame(columns=template_headers)
        df.to_excel(writer, sheet_name=TARGET_SHEET, index=False)

        workbook = writer.book
        worksheet = writer.sheets[TARGET_SHEET]

        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': '#FFFFFF',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter'
        })
        data_fmt = workbook.add_format({'valign': 'vcenter'})

        for col_num, hdr in enumerate(template_headers):
            worksheet.write(0, col_num, hdr, header_fmt)

        data_row = 1
        for row_data in rows:
            for col_num, hdr in enumerate(template_headers):
                if hdr in col_to_db:
                    entry_type, field_name = col_to_db[hdr]
                    if entry_type == 'month':
                        idx = field_index.get('report_month')
                        value = row_data[idx] if idx is not None else ''
                    elif entry_type == 'year':
                        idx = field_index.get('report_year')
                        value = row_data[idx] if idx is not None else ''
                    elif entry_type == 'field':
                        idx = field_index.get(field_name)
                        value = row_data[idx] if idx is not None else ''
                    else:
                        value = ''

                if value is None:
                    value = ''
                elif isinstance(value, datetime):
                    worksheet.write_datetime(data_row, col_num, value, data_fmt)
                    continue
                elif isinstance(value, (int, float)):
                    pass
                else:
                    value = str(value) if value else ''

                worksheet.write(data_row, col_num, value, data_fmt)
            data_row += 1

    print(f"Excel generado: {OUTPUT_PATH}")
    return OUTPUT_PATH, len(rows)


if __name__ == '__main__':
    try:
        output_path, count = generate_excel()
        print(f"\nExcel generado exitosamente!")
        print(f"Archivo: {output_path}")
        print(f"Total registros: {count}")
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
