"""
Generar reportes XLSX desde tablas MySQL usando templates
- Lee templates del directorio template
- Usa sheet Match para mapear columnas de MySQL a columnas del reporte
- El Match tiene indices basados en el CSV original (no en la tabla MySQL)
- Genera reporte en sheet Template
"""
import mysql.connector
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pathlib import Path
from datetime import datetime
from copy import copy
import csv
import re
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config

TEMPLATE_DIR = config.TEMPLATE_DIR
OUTPUT_DIR = config.REPORTS_DIR
OUTPUT_DIR.mkdir(exist_ok=True)


def sanitize_column_name(col_name):
    """Sanitizar nombre de columna (misma logica que create_ftp_csv_tables.py)"""
    if not col_name or col_name.strip() == '':
        return 'column'

    name = col_name.strip()
    name = name.replace(' – ', '_')
    name = name.replace('-', '_')
    name = name.replace(' ', '_')
    name = name.replace('/', '_')
    name = name.replace('\\', '_')
    name = name.replace('(', '_')
    name = name.replace(')', '_')
    name = name.replace('"', '')
    name = name.replace("'", '')
    name = name.replace('?', '')
    name = name.replace('!', '')
    name = name.replace(',', '_')
    name = name.replace('.', '_')
    name = name.replace(':', '_')
    name = name.replace(';', '_')
    name = name.replace('|', '_')
    name = name.replace('__', '_')
    name = name.strip('_')
    name = re.sub(r'[^\w]', '', name)
    name = name.lower()

    if not name:
        return 'column'

    if len(name) > 64:
        last_underscore = name.rfind('_', 0, 60)
        if last_underscore > 0:
            name = name[:last_underscore]
        else:
            name = name[:60]

    return name


def sanitize_column_names_unique(headers):
    """Sanitizar nombres de columnas sin duplicados (misma logica que al importar)"""
    sanitized = {}
    seen = {}
    suffix_counter = {}

    for header in headers:
        sanitized_name = sanitize_column_name(header)

        if sanitized_name in seen:
            if sanitized_name not in suffix_counter:
                suffix_counter[sanitized_name] = 1
            else:
                suffix_counter[sanitized_name] += 1

            max_base_len = 60
            if len(sanitized_name) > max_base_len:
                sanitized_name = sanitized_name[:max_base_len]

            new_name = f"{sanitized_name}_{suffix_counter[sanitized_name]}"
        else:
            new_name = sanitized_name

        seen[new_name] = True
        sanitized[header] = new_name

    return sanitized


def get_csv_headers(csv_filename):
    """
    Leer headers del CSV original
    Retorna lista de nombres de columnas en orden (indice base 1)
    """
    # Buscar archivo CSV correspondiente
    csv_files = list(FTP_CSV_DIR.glob(f'{csv_filename}*.csv'))

    if not csv_files:
        return None

    csv_path = csv_files[0]

    with open(csv_path, 'r', encoding='utf-8', errors='ignore') as f:
        reader = csv.reader(f)
        headers = next(reader)

    return headers


def get_csv_data(csv_filename):
    """
    Leer todos los datos del CSV original
    Retorna lista de diccionarios {column_name: value} y headers
    """
    # Buscar archivo CSV correspondiente
    csv_files = list(FTP_CSV_DIR.glob(f'{csv_filename}*.csv'))

    if not csv_files:
        return None, None

    csv_path = csv_files[0]

    with open(csv_path, 'r', encoding='utf-8', errors='ignore') as f:
        reader = csv.reader(f)
        headers = next(reader)
        rows = list(reader)

    # Convertir a lista de diccionarios
    data = []
    for row in rows:
        row_dict = {}
        for i, header in enumerate(headers):
            if i < len(row):
                row_dict[header] = row[i]
            else:
                row_dict[header] = ''
        data.append(row_dict)

    return data, headers


def get_connection():
    """Crear conexion a MySQL"""
    return mysql.connector.connect(**config.MYSQL_CONFIG)


def get_table_name_from_template(template_name):
    """
    Obtener nombre de tabla desde nombre de template
    """
    name = template_name.replace('_2026.xlsx', '').replace('_2026.XLSX', '')

    mapping = {
        'Visit_Log_Report(Optblue)': 'visit_log_optblue',
        'Visits_Disposition_Report': 'visits_disposition_report',
        'Call_Disposition_Report': 'call_disposition_report',
        'Calls_Consolidate_Report': 'calls_consolidate_report'
    }

    if name in mapping:
        return mapping[name]

    return name.lower()


def read_match_sheet(template_path):
    """
    Leer sheet Match y retornar diccionario de mapeo
    Retorna: {report_column_name: csv_column_index}
    El indice base es 1 (primera columna del CSV)
    """
    wb = load_workbook(template_path, data_only=False)

    if 'Match' not in wb.sheetnames:
        wb.close()
        return None

    ws = wb['Match']
    mapping = {}

    row = 2
    while True:
        report_col = ws.cell(row, 1).value
        table_idx = ws.cell(row, 2).value

        if not report_col:
            break

        try:
            if isinstance(table_idx, str):
                table_idx = int(table_idx.strip())
            elif table_idx is None:
                table_idx = 0
            else:
                table_idx = int(table_idx)
        except (ValueError, TypeError):
            table_idx = 0

        mapping[report_col] = table_idx
        row += 1

    wb.close()
    return mapping


def read_match_sheet_with_mysql_columns(template_path):
    """
    Leer sheet Match y retornar diccionario de mapeo usando nombres de columnas MySQL
    Retorna: {report_column_name: mysql_column_name}
    Lee la columna C del sheet Match que contiene el nombre de la columna en MySQL
    """
    wb = load_workbook(template_path, data_only=False)

    if 'Match' not in wb.sheetnames:
        wb.close()
        return None

    ws = wb['Match']
    mapping = {}

    row = 2
    while True:
        report_col = ws.cell(row, 1).value
        mysql_col_name = ws.cell(row, 3).value  # Columna C: nombre de columna MySQL

        if not report_col:
            break

        if mysql_col_name and str(mysql_col_name).strip():
            # Sanitizar el nombre de columna MySQL
            sanitized_name = str(mysql_col_name).strip()
            mapping[report_col] = sanitized_name

        row += 1

    # Aplicar correcciones manuales para columnas que no coinciden exactamente
    corrections = {
        'type_of_relationship_with_processor_bank': 'type_of_relationship_with_processor__bank',
        'source_create_date_1': 'source_create_date_1',
        'Phone': 'phone',
        'Sales_Channel_Name': 'sales_channel_name',
    }

    for report_col, mysql_col in list(mapping.items()):
        if mysql_col in corrections:
            mapping[report_col] = corrections[mysql_col]

    wb.close()
    return mapping


def get_template_headers(template_path):
    """
    Leer headers del sheet Template (fila 1)
    Retorna lista de nombres de columnas en orden
    """
    wb = load_workbook(template_path, data_only=False)
    ws = wb['Template']

    headers = []
    col = 1
    while True:
        val = ws.cell(1, col).value
        if not val:
            break
        headers.append(val)
        col += 1

    wb.close()
    return headers


def build_column_mapping(match_mapping, csv_headers, table_columns):
    """
    Construir mapeo final: {report_column_name: mysql_column_index}
    1. match_mapping: {report_col: csv_index}
    2. csv_headers: lista de nombres en CSV (orden, indice base 1)
    3. table_columns: lista de nombres en tabla MySQL (orden, indice base 0 incluye id)

    Retorna indice base 0 para usar en db_row

    NOTA: db_row incluye id y excluye created_at/updated_at (que estan al final)
    Por lo tanto, db_row[i] = table_columns[i] para columnas de datos
    """
    final_mapping = {}

    # Sanitizar nombres del CSV para comparar con MySQL
    csv_sanitized = sanitize_column_names_unique(csv_headers)

    for report_col, csv_idx in match_mapping.items():
        if csv_idx < 1 or csv_idx > len(csv_headers):
            continue

        # Obtener nombre original del CSV (indice base 1)
        csv_col_name = csv_headers[csv_idx - 1]

        # Obtener nombre sanitizado (como esta en MySQL)
        mysql_col_name = csv_sanitized.get(csv_col_name)

        if mysql_col_name:
            # Buscar indice en tabla MySQL
            try:
                mysql_idx = table_columns.index(mysql_col_name)
                # Como created_at y updated_at estan al final, y db_row las excluye,
                # pero incluye id, el indice es el mismo para columnas de datos
                final_mapping[report_col] = mysql_idx
            except ValueError:
                pass  # Columna no encontrada en tabla

    return final_mapping


def copy_template_formatting(src_ws, dst_ws, num_rows):
    """Copiar formatos del template a las nuevas filas de datos"""
    template_row = 2

    for row in range(2, 2 + num_rows):
        for col in range(1, src_ws.max_column + 1):
            src_cell = src_cell = src_ws.cell(template_row, col)
            dst_cell = dst_ws.cell(row, col)

            if src_cell.has_style:
                dst_cell.font = copy(src_cell.font)
                dst_cell.fill = copy(src_cell.fill)
                dst_cell.border = copy(src_cell.border)
                dst_cell.alignment = copy(src_cell.alignment)
                dst_cell.number_format = src_cell.number_format


def generate_report(template_name, csv_filename, conn):
    """
    Generar reporte desde template
    csv_filename: nombre base del archivo CSV (sin extension ni fecha)
    """
    print(f"\n--- Generando: {template_name} ---")

    template_path = TEMPLATE_DIR / template_name

    if not template_path.exists():
        print(f"  [ERROR] Template no encontrado: {template_path}")
        return None

    # Obtener nombre de tabla
    table_name = get_table_name_from_template(template_name)
    print(f"  Tabla: {table_name}")

    # Leer headers del CSV original
    csv_headers = get_csv_headers(csv_filename)
    if not csv_headers:
        print(f"  [WARNING] No se encontro CSV: {csv_filename}")
        csv_data = None
    else:
        print(f"  Columnas CSV: {len(csv_headers)}")
        # Leer datos del CSV para calcular Month/Year
        csv_data, _ = get_csv_data(csv_filename)
        if csv_data:
            print(f"  Filas CSV: {len(csv_data)}")

    # Verificar si es Visits Disposition Report o Visit Log Optblue para usar mapeo por nombre de columna MySQL
    use_mysql_names = (template_name == 'Visits_Disposition_Report_2026.xlsx' or
                       template_name == 'Visit_Log_Report(Optblue)_2026.xlsx')

    # Leer mapeo del Match Sheet
    if use_mysql_names:
        match_mapping = read_match_sheet_with_mysql_columns(template_path)
    else:
        match_mapping = read_match_sheet(template_path)

    if not match_mapping:
        print(f"  [ERROR] No se encontro sheet Match")
        return None

    print(f"  Mapeo Match: {len(match_mapping)} columnas")

    # Obtener columnas de la tabla MySQL
    cursor = conn.cursor()
    cursor.execute(f"DESCRIBE `{table_name}`")
    table_columns = [row[0] for row in cursor.fetchall()]
    print(f"  Columnas tabla MySQL: {len(table_columns)}")

    # Construir mapeo final: report_col -> mysql_column_index o mysql_column_name
    if use_mysql_names:
        # Para Visits Disposition: mapeo directo usando nombres de columnas MySQL
        final_mapping = {}
        for report_col, mysql_col_name in match_mapping.items():
            # Buscar la columna MySQL en la tabla
            sanitized_name = sanitize_column_name(mysql_col_name)

            # Buscar coincidencia exacta primero - PRIORIZAR mysql_col_name sin sanitizar
            found_col = None
            if mysql_col_name in table_columns:
                found_col = mysql_col_name
            elif sanitized_name in table_columns:
                found_col = sanitized_name
            else:
                # Buscar variantes: sin guiones bajos, con guiones diferentes, etc.
                clean_compare = sanitized_name.replace('_', '')
                for table_col in table_columns:
                    clean_table = table_col.replace('_', '')
                    if clean_compare == clean_table.lower():
                        found_col = table_col
                        break

            if found_col:
                final_mapping[report_col] = found_col
    elif csv_headers:
        final_mapping = build_column_mapping(match_mapping, csv_headers, table_columns)
    else:
        # Si no hay CSV, usar mapeo directo (asumiendo indices correctos)
        final_mapping = {}
        for report_col, csv_idx in match_mapping.items():
            if csv_idx > 0 and csv_idx < len(table_columns):
                final_mapping[report_col] = csv_idx

    print(f"  Mapeo final: {len(final_mapping)} columnas")

    # Leer headers del template
    template_headers = get_template_headers(template_path)
    print(f"  Columnas template: {len(template_headers)}")

    # Obtener datos de MySQL
    data_columns = [col for col in table_columns if col not in ['created_at', 'updated_at']]
    data_columns_sql = ', '.join([f'`{col}`' for col in data_columns])

    cursor.execute(f"SELECT {data_columns_sql} FROM `{table_name}` ORDER BY id")
    rows = cursor.fetchall()

    print(f"  Filas obtenidas: {len(rows)}")

    # Cargar template
    wb = load_workbook(template_path)
    template_ws = wb['Template']

    # ELIMINAR todos los sheets excepto Template
    for sheet_name in wb.sheetnames:
        if sheet_name != 'Template':
            del wb[sheet_name]

    # Limpiar datos existentes del Template (mantener solo encabezados)
    # Eliminar filas completas desde la fila 2 hasta el final
    # Esto es mas eficiente que limpiar celda por celda
    max_row_to_check = max(template_ws.max_row, 1000)  # Limpiar hasta 1000 filas
    for row in range(max_row_to_check, 1, -1):  # Iterar hacia atras para evitar problemas
        if row > 1:  # Mantener encabezados
            template_ws.delete_rows(row, 1)

    # Escribir datos (sin formato)
    from openpyxl.styles import Font, PatternFill, Border, Alignment

    # Buscar indices de Month y Year en template
    month_col_idx = None
    year_col_idx = None

    for idx, header in enumerate(template_headers, start=1):
        if header == 'Month':
            month_col_idx = idx
        elif header == 'Year':
            year_col_idx = idx

    data_row = 2

    for db_row in rows:
        for col_idx, header_name in enumerate(template_headers, start=1):
            mysql_mapping = final_mapping.get(header_name)

            if mysql_mapping is None:
                continue

            # Determinar valor según el tipo de mapeo
            if use_mysql_names:
                # Mapeo por nombre de columna MySQL
                mysql_col_name = mysql_mapping
                # IMPORTANTE: Usar data_columns no table_columns porque db_row excluye created_at/updated_at
                if mysql_col_name in data_columns:
                    col_index = data_columns.index(mysql_col_name)
                    if col_index < len(db_row):
                        value = db_row[col_index]
                    else:
                        value = ''
                else:
                    value = ''
            else:
                # Mapeo por indice (metodo original)
                mysql_col_index = mysql_mapping
                if mysql_col_index > 0 and mysql_col_index <= len(db_row):
                    value = db_row[mysql_col_index]
                else:
                    value = ''

            if value is None:
                value = ''

            cell = template_ws.cell(data_row, col_idx)
            cell.value = value

        # Escribir Month y Year desde columnas de MySQL (calculadas del CSV)
        # Buscar indices de 'month' y 'year' en data_columns (no table_columns)
        month_idx_in_db = None
        year_idx_in_db = None
        if 'month' in data_columns:
            month_idx_in_db = data_columns.index('month')
        if 'year' in data_columns:
            year_idx_in_db = data_columns.index('year')

        # Escribir Month
        if month_col_idx is not None and month_idx_in_db is not None and month_idx_in_db < len(db_row):
            template_ws.cell(data_row, month_col_idx).value = db_row[month_idx_in_db]

        # Escribir Year
        if year_col_idx is not None and year_idx_in_db is not None and year_idx_in_db < len(db_row):
            template_ws.cell(data_row, year_col_idx).value = db_row[year_idx_in_db]

        # Limpiar formato de TODA la fila de datos (despues de escribir)
        for col_idx in range(1, len(template_headers) + 1):
            cell = template_ws.cell(data_row, col_idx)
            # Crear Font completamente nuevo sin heredar tema
            new_font = Font(
                name='Calibri',
                size=11,
                color='FF000000',
                bold=False,
                italic=False,
                underline='none',
                strike=False
            )
            cell.font = new_font
            cell.fill = PatternFill(fill_type=None)
            cell.border = Border()
            cell.alignment = Alignment(horizontal='general', vertical='bottom', wrap_text=False)

        # Caso especial: Visits Disposition Report - Columna I (9) = MerchantState = "PR"
        if template_name == 'Visits_Disposition_Report_2026.xlsx':
            template_ws.cell(data_row, 9).value = 'PR'

        data_row += 1

    # Guardar archivo (sin copiar formatos)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    output_filename = f"{template_name.replace('.xlsx', '')}-{timestamp}.xlsx"
    output_path = OUTPUT_DIR / output_filename

    wb.save(output_path)
    wb.close()

    print(f"  [OK] Reporte generado: {output_path}")

    return str(output_path)


def main():
    """Funcion principal"""
    print("=== Generar Reportes FTP desde Templates ===")

    print("\nConectando a MySQL...")
    conn = get_connection()
    print("Conexion exitosa!")

    # Templates con sus CSV correspondientes
    templates = [
        ('Calls_Consolidate_Report_2026.xlsx', 'Calls_Consolidate_Report'),
        ('Call_Disposition_Report_2026.xlsx', 'Call_Disposition_Report'),
        ('Visit_Log_Report(Optblue)_2026.xlsx', 'Visit_Log_(Optblue)'),
        ('Visits_Disposition_Report_2026.xlsx', 'Visits_Disposition_Report')
    ]

    generated_reports = []

    for template, csv_base in templates:
        try:
            result = generate_report(template, csv_base, conn)
            if result:
                generated_reports.append(result)
        except Exception as e:
            print(f"  [ERROR] {e}")
            import traceback
            traceback.print_exc()

    conn.close()

    print("\n=== Proceso completado ===")
    print(f"Reportes generados: {len(generated_reports)}")
    for report in generated_reports:
        print(f"  - {report}")


if __name__ == '__main__':
    main()
