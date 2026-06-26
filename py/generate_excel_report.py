"""
Generar Service-Level-Report.xlsx desde MySQL para ICC Amex
Genera Excel desde cero (sin template) para evitar corrupcion
"""

import mysql.connector
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config

OUTPUT_PATH = config.OUTPUT_FILES['service_level']


def get_connection():
    return mysql.connector.connect(**config.MYSQL_CONFIG)


def get_service_level_data(conn):
    cursor = conn.cursor()
    query = """
    SELECT
        DATE(o.date) as fecha,
        COUNT(*) as total_dials,
        AVG(CASE WHEN o.talk_time IS NOT NULL AND o.talk_time != '00:00:00'
            THEN TIME_TO_SEC(o.talk_time) ELSE NULL END) as avg_talk_time_seconds,
        AVG(CASE WHEN o.after_call_work_time IS NOT NULL AND o.after_call_work_time != '00:00:00'
            THEN TIME_TO_SEC(o.after_call_work_time) ELSE NULL END) as avg_followup_seconds,
        AVG(CASE WHEN o.handle_time IS NOT NULL AND o.handle_time != '00:00:00'
            THEN TIME_TO_SEC(o.handle_time) ELSE NULL END) as avg_handle_time_seconds,
        COALESCE(MAX(l.production_hours), 0) as production_hours
    FROM outbound_call_log o
    LEFT JOIN (
        SELECT DATE(date) as ldate, SUM(TIME_TO_SEC(login_time)) / 3600.0 as production_hours
        FROM login_logout
        WHERE date IS NOT NULL AND login_time IS NOT NULL
        GROUP BY DATE(date)
    ) l ON DATE(o.date) = l.ldate
    WHERE o.date IS NOT NULL
      AND YEAR(o.date) = 2026
    GROUP BY DATE(o.date)
    ORDER BY DATE(o.date)
    """
    cursor.execute(query)
    results = cursor.fetchall()
    cursor.close()

    def sec_to_time(seconds):
        if seconds is None or seconds == 0:
            return '00:00:00'
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = int(seconds % 60)
        return f'{hours:02d}:{minutes:02d}:{secs:02d}'

    data = []
    for row in results:
        data.append({
            'Date': row[0],
            'Total Dials': row[1] or 0,
            'OB Average Talk Time': sec_to_time(row[2]),
            'OB Average Followup': sec_to_time(row[3]),
            'OB Total Average Handle Time': sec_to_time(row[4]),
            'Production Hours': round(float(row[5] or 0), 2)
        })
    return data


def generate_excel(output_path=None):
    print("=== Generando Service-Level-Report.xlsx ===\n")

    if output_path is None:
        output_path = OUTPUT_PATH

    try:
        print("Conectando a MySQL...")
        conn = get_connection()
        print("Conexion exitosa!\n")

        sl_data = get_service_level_data(conn)
        print(f"Registros obtenidos: {len(sl_data)}")

        wb = Workbook()
        ws = wb.active
        ws.title = 'Service Level'

        headers = ['Date', 'Total Dials', 'OB Average Talk Time', 'OB Average Followup',
                    'OB Total Average Handle Time', 'Production Hours']
        header_font = Font(bold=True)
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_font_white = Font(bold=True, color='FFFFFF')

        for col, header in enumerate(headers, 1):
            cell = ws.cell(1, col, header)
            cell.font = header_font_white
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')

        for row_idx, record in enumerate(sl_data, 2):
            ws.cell(row_idx, 1, record['Date'])
            ws.cell(row_idx, 2, record['Total Dials'])
            ws.cell(row_idx, 3, record['OB Average Talk Time'])
            ws.cell(row_idx, 4, record['OB Average Followup'])
            ws.cell(row_idx, 5, record['OB Total Average Handle Time'])
            ph_cell = ws.cell(row_idx, 6, record['Production Hours'])
            ph_cell.number_format = '0.00'

        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 20
        ws.column_dimensions['E'].width = 28
        ws.column_dimensions['F'].width = 16

        wb.save(str(output_path))
        wb.close()
        print(f"\nExcel generado: {output_path}")
        return str(output_path)

    except Exception as e:
        print(f"Error: {e}")
        raise
    finally:
        if 'conn' in locals():
            conn.close()
            print("Conexion cerrada")


if __name__ == "__main__":
    generate_excel()
