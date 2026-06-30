#!/usr/bin/env python3
import mysql.connector
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from pathlib import Path
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config

TEMPLATE_PATH = config.TEMPLATES['pr_escalation']
OUTPUT_PATH = config.OUTPUT_FILES['pr_escalation']

REPORT_MAPPING = {
    'Date': 'date',
    'Agent Name': 'agent_name',
    'Program': 'program',
    'Contact Name': 'contact_name',
    'Phone Number': 'phone_number',
    'Business Name': 'business_name',
    'SE Number': 'se_number',
    'Complaint Reason': 'complaint_reason',
    'Issue Description': 'issue_description',
    'Callback Requested?': 'callback_requested',
    'Call Resolution Notes': 'call_resolution_notes',
    'Resolved?': 'resolved',
}


def read_template_info(template_path):
    wb = load_workbook(template_path, read_only=True)

    sheet_name = None
    for sn in wb.sheetnames:
        if sn.startswith('PR Calls and Visits Escalation'):
            sheet_name = sn
            break
    if not sheet_name:
        sheet_name = wb.sheetnames[0]

    ws = wb[sheet_name]
    headers = []
    for col in range(1, 200):
        val = ws.cell(1, col).value
        if val:
            headers.append(str(val).strip())
        else:
            break

    wb.close()
    return headers, sheet_name


def get_data_from_db(from_date=None, to_date=None):
    conn = mysql.connector.connect(**config.MYSQL_CONFIG)
    cursor = conn.cursor()
    db_fields = list(REPORT_MAPPING.values())
    fields_str = ', '.join([f'`{f}`' for f in db_fields])
    query = f"SELECT {fields_str} FROM pr_calls_visits_escalation"
    params = []
    if from_date and to_date:
        query += " WHERE date >= %s AND date <= %s"
        params.extend([from_date, to_date])
    query += " ORDER BY date DESC"
    cursor.execute(query, params)
    rows = cursor.fetchall()
    field_names = [d[0] for d in cursor.description]
    cursor.close()
    conn.close()
    return rows, field_names


def generate_excel(from_date=None, to_date=None):
    print(f"Template: {TEMPLATE_PATH}")

    template_headers, sheet_name = read_template_info(TEMPLATE_PATH)
    print(f"Columnas en template: {len(template_headers)}")

    rows, field_names = get_data_from_db(from_date, to_date)
    print(f"Filas obtenidas: {len(rows)}")

    field_index = {name: idx for idx, name in enumerate(field_names)}

    output_dir = os.path.dirname(OUTPUT_PATH)
    os.makedirs(output_dir, exist_ok=True)

    with pd.ExcelWriter(OUTPUT_PATH, engine='xlsxwriter') as writer:
        df = pd.DataFrame(columns=template_headers)
        df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name[:31]]

        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': '#FFFFFF',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter'
        })
        data_fmt = workbook.add_format({'valign': 'vcenter'})

        for col_num, hdr in enumerate(template_headers):
            worksheet.write(0, col_num, hdr, header_fmt)

        col_map = {}
        for idx, hdr in enumerate(template_headers):
            if hdr in REPORT_MAPPING:
                db_field = REPORT_MAPPING[hdr]
                if db_field in field_index:
                    col_map[idx] = field_index[db_field]

        data_row = 1
        for row_data in rows:
            for col_num in range(len(template_headers)):
                if col_num in col_map:
                    value = row_data[col_map[col_num]]
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
    return OUTPUT_PATH


if __name__ == '__main__':
    from_date = None
    to_date = None
    if '--from' in sys.argv and '--to' in sys.argv:
        from_idx = sys.argv.index('--from')
        to_idx = sys.argv.index('--to')
        from_date = sys.argv[from_idx + 1] if from_idx + 1 < len(sys.argv) else None
        to_date = sys.argv[to_idx + 1] if to_idx + 1 < len(sys.argv) else None
    generate_excel(from_date, to_date)
