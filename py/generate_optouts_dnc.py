#!/usr/bin/env python3
"""
Generar archivo TXT de OPTOUTS-DNC desde MySQL
Formato: "sell_bus_phone_no","sell_se_no"
"""
import mysql.connector
from pathlib import Path
from datetime import datetime
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config

OUTPUT_PATH = config.OUTPUT_FILES['optouts_dnc']


def get_data_from_db():
    conn = mysql.connector.connect(**config.MYSQL_CONFIG)
    cursor = conn.cursor()

    query = """
        SELECT dnc_number, seller_id
        FROM optouts_dnc
        ORDER BY id DESC
    """

    cursor.execute(query)
    rows = cursor.fetchall()

    cursor.close()
    conn.close()

    return rows


def generate_txt():
    print("Obteniendo datos de la base de datos...")
    rows = get_data_from_db()
    print(f"  Total registros: {len(rows)}")

    print(f"Generando TXT: {OUTPUT_PATH}")

    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        f.write('"sell_bus_phone_no","sell_se_no"\n')
        for row in rows:
            dnc_number = row[0] if row[0] else ''
            seller_id = row[1] if row[1] else ''
            if dnc_number and str(dnc_number).startswith('+'):
                dnc_number = str(dnc_number)[1:]
            f.write(f'"{dnc_number}","{seller_id}"\n')

    print(f"Archivo generado: {OUTPUT_PATH}")
    return OUTPUT_PATH, len(rows)


if __name__ == '__main__':
    try:
        output_path, count = generate_txt()
        print(f"\nTXT generado exitosamente!")
        print(f"Archivo: {output_path}")
        print(f"Total registros: {count}")
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
