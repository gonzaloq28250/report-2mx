"""
Generar Service-Level-Report.xlsx desde MySQL para ICC Amex
Solo hoja Service Level con Production Hours
"""

import mysql.connector
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config

TEMPLATE_PATH = config.TEMPLATES['service_level']
OUTPUT_PATH = config.OUTPUT_FILES['service_level']


def get_connection():
    return mysql.connector.connect(**config.MYSQL_CONFIG)


def get_service_level_data(conn):
    """Obtener datos para la hoja Service Level"""
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
    print(f"Registros obtenidos: {len(results)}")
    if results:
        print(f"Primera fila: {results[0]}")
    cursor.close()

    data = []
    for row in results:
        date_val = row[0]
        total_dials = row[1] or 0
        avg_talk = row[2] or 0
        avg_followup = row[3] or 0
        avg_handle = row[4] or 0
        production_hours = row[5] or 0

        def sec_to_time(seconds):
            if seconds is None or seconds == 0:
                return '00:00:00'
            hours = int(seconds // 3600)
            minutes = int((seconds % 3600) // 60)
            secs = int(seconds % 60)
            return f'{hours:02d}:{minutes:02d}:{secs:02d}'

        data.append({
            'Date': date_val,
            'Total Dials': total_dials,
            'OB Average Talk Time': sec_to_time(avg_talk),
            'OB Average Followup': sec_to_time(avg_followup),
            'OB Total Average Handle Time': sec_to_time(avg_handle),
            'Production Hours': round(production_hours, 2)
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

        print(f"Cargando template: {TEMPLATE_PATH}")
        wb = load_workbook(TEMPLATE_PATH)

        print("Procesando hoja: Service Level")
        ws = wb['Service Level']

        sl_data = get_service_level_data(conn)

        header_row = None
        for row in range(1, 20):
            if ws.cell(row, 1).value == 'Date':
                header_row = row
                break

        if header_row:
            for row in range(header_row + 1, 100):
                for col in range(1, 20):
                    cell = ws.cell(row, col)
                    if cell.value:
                        cell.value = None

            data_row = header_row + 1
            for record in sl_data:
                ws.cell(data_row, 1).value = record['Date']
                ws.cell(data_row, 2).value = record['Total Dials']
                ws.cell(data_row, 3).value = record['OB Average Talk Time']
                ws.cell(data_row, 4).value = record['OB Average Followup']
                ws.cell(data_row, 5).value = record['OB Total Average Handle Time']
                ph_cell = ws.cell(data_row, 6)
                ph_cell.value = float(record['Production Hours'])
                ph_cell.number_format = '0.00'
                data_row += 1

            print(f"  Escritos {len(sl_data)} registros")

        print("Eliminando hojas no necesarias...")
        sheets_to_remove = []
        for sheet_name in wb.sheetnames:
            if sheet_name != 'Service Level':
                sheets_to_remove.append(sheet_name)

        for sheet_name in sheets_to_remove:
            wb.remove(wb[sheet_name])
            print(f"  Eliminada hoja: {sheet_name}")

        wb.save(output_path)
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
