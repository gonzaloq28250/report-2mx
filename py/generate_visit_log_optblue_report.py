#!/usr/bin/env python3
import mysql.connector
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Alignment
from pathlib import Path
from datetime import datetime
import sys
import os
import shutil

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config


def sanitize_field_name(field_name):
    if not field_name:
        return None
    field = field_name.strip()
    field = field.replace(' ', '_').replace('-', '_')
    field = field.lower()
    return field


def read_match_mapping(wb, existing_columns):
    if 'Match' not in wb.sheetnames:
        raise ValueError("El template no tiene un sheet llamado 'Match'")
    ws = wb['Match']
    mapping = {}
    col_lower = {k: v for k, v in existing_columns.items()}

    for row in range(2, ws.max_row + 1):
        report_col = ws.cell(row, 1).value
        db_field = ws.cell(row, 3).value
        if report_col and db_field:
            report_col = str(report_col).strip()
            db_field_raw = str(db_field).strip()
            sanitized_db = sanitize_field_name(db_field_raw)

            found = col_lower.get(sanitized_db)
            if not found:
                for lk, lv in col_lower.items():
                    if sanitized_db == lk.replace(' ', '_').replace('-', '_'):
                        found = lv
                        break
            if not found:
                import re
                base = re.sub(r'_\d+$', '', sanitized_db)
                for lk, lv in col_lower.items():
                    if base == re.sub(r'_\d+$', '', lk.replace(' ', '_').replace('-', '_')):
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
    return mapping


def get_table_columns(cursor, table_name):
    cursor.execute(f"SHOW COLUMNS FROM `{table_name}`")
    return {col[0].lower(): col[0] for col in cursor.fetchall()}


def get_data_from_db(db_fields, table_name):
    conn = mysql.connector.connect(**config.MYSQL_CONFIG)
    cursor = conn.cursor()
    valid_fields = [f for f in db_fields if f]
    if valid_fields:
        fields_str = ', '.join([f'`{f}`' for f in valid_fields])
    else:
        fields_str = '*'
    query = f"SELECT {fields_str} FROM `{table_name}` ORDER BY id DESC"
    cursor.execute(query)
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return rows, valid_fields


def generate_excel():
    table_name = 'visit_log_optblue_report'
    template_path = config.TEMPLATES['visit_log_optblue']
    output_path = config.OUTPUT_FILES['visit_log_optblue']

    print(f"Template: {template_path}")
    wb = load_workbook(template_path)

    conn = mysql.connector.connect(**config.MYSQL_CONFIG)
    cursor = conn.cursor()
    existing_columns = get_table_columns(cursor, table_name)
    cursor.close()
    conn.close()
    print(f"Columnas en tabla: {len(existing_columns)}")

    match_mapping = read_match_mapping(wb, existing_columns)
    print(f"Mapeo Match: {len(match_mapping)} columnas")
    db_fields = list(match_mapping.values())
    rows, valid_fields = get_data_from_db(db_fields, table_name)
    print(f"Filas obtenidas: {len(rows)}")

    if 'Template' not in wb.sheetnames:
        raise ValueError("El template no tiene un sheet llamado 'Template'")
    ws = wb['Template']

    header_row = None
    for row in range(1, ws.max_row + 1):
        if ws.cell(row, 1).value and 'Visitas' in str(ws.cell(row, 1).value):
            header_row = row
            break
    if not header_row:
        header_row = 1

    col_mapping = {}
    max_col = ws.max_column if ws.max_column > 50 else 100
    for col in range(1, max_col + 1):
        header_value = ws.cell(header_row, col).value
        if header_value:
            header_str = str(header_value).strip()
            if header_str in match_mapping:
                col_mapping[match_mapping[header_str]] = col

    data_start_row = header_row + 1
    for row in range(data_start_row, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            ws.cell(row, col).value = None

    for row_idx, row_data in enumerate(rows, start=data_start_row):
        for col_idx, db_field in enumerate(valid_fields):
            if db_field in col_mapping:
                excel_col = col_mapping[db_field]
                cell = ws.cell(row_idx, excel_col)
                value = row_data[col_idx]
                if isinstance(value, datetime):
                    cell.value = value
                elif isinstance(value, (int, float)):
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

    wb.save(output_path)
    wb.close()
    print(f"Excel generado: {output_path}")
    return output_path


if __name__ == '__main__':
    generate_excel()
