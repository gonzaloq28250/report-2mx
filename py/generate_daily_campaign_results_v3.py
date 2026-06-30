"""
Generar Daily_Campaign_Results.xlsx desde MySQL
VERSION 3 - Usa xlsxwriter para evitar corrupcion Excel
- Crea archivo desde cero con formulas Excel
- Mantiene todas las formulas (SUM, divisiones, cross-sheet refs)
"""
import mysql.connector
from pathlib import Path
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config

TEMPLATE_PATH = config.TEMPLATES['daily_campaign']
OUTPUT_DIR = config.REPORTS_DIR
OUTPUT_DIR.mkdir(exist_ok=True)


def get_connection():
    return mysql.connector.connect(**config.MYSQL_CONFIG)


def get_list_name_records_map(conn):
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS list_name_records (id INT AUTO_INCREMENT PRIMARY KEY, list_name VARCHAR(255) NOT NULL, list_records INT NOT NULL DEFAULT 0, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP, UNIQUE KEY (list_name))")
    cursor.execute("SELECT COUNT(*) FROM list_name_records")
    count = cursor.fetchone()[0]
    if count == 0:
        cursor.execute("SELECT COUNT(*) FROM information_schema.tables WHERE table_schema = DATABASE() AND table_name = 'list_status_summary_report'")
        if cursor.fetchone()[0] > 0:
            cursor.execute("SELECT COUNT(*) FROM list_status_summary_report")
            src_count = cursor.fetchone()[0]
            if src_count > 0:
                print(f"  list_name_records vacia, migrando {src_count} filas...")
                cursor.execute("""
                    SELECT lead_list_name, SUM(CAST(list_records AS UNSIGNED)) as total_records
                    FROM list_status_summary_report
                    WHERE lead_list_name IS NOT NULL AND lead_list_name != ''
                    GROUP BY lead_list_name
                """)
                for row in cursor.fetchall():
                    cursor.execute(
                        "INSERT INTO list_name_records (list_name, list_records) VALUES (%s, %s) "
                        "ON DUPLICATE KEY UPDATE list_records = VALUES(list_records)",
                        (row[0].strip(), row[1])
                    )
                conn.commit()
    cursor.close()
    query = "SELECT list_name, list_records FROM list_name_records"
    df = pd.read_sql(query, conn)
    return {row['list_name'].strip(): row['list_records'] for _, row in df.iterrows()}


def get_support_tables_map(conn):
    query = """
    SELECT disposition, final_disposition, redial, consultation,
           agent_completes, presentations, unworkable, no_contact, net_contacts
    FROM support_tables ORDER BY disposition
    """
    df = pd.read_sql(query, conn)
    result = {k: [] for k in ['final_disposition', 'redial', 'consultation', 'agent_completes', 'presentations', 'unworkable', 'no_contact', 'net_contacts']}
    for _, row in df.iterrows():
        disp = row['disposition'].strip()
        for key in result:
            if row[key]:
                result[key].append(disp)
    for key, values in result.items():
        if values:
            print(f"    {key:20s}: {len(values)} dispositions")
    return result


def get_summary_data(conn, st_map):
    query_hours = """
    SELECT DATE(date) as fecha,
        SUM(CASE WHEN skill_availability LIKE '%%PR%%' THEN TIME_TO_SEC(login_time) ELSE 0 END) / 3600.0 as spanish_hours,
        SUM(CASE WHEN skill_availability LIKE '%%VI%%' THEN TIME_TO_SEC(login_time) ELSE 0 END) / 3600.0 as english_hours
    FROM login_logout WHERE date IS NOT NULL AND YEAR(date) = 2026 GROUP BY DATE(date) ORDER BY fecha DESC
    """
    hours_df = pd.read_sql(query_hours, conn)

    query_dials = "SELECT DATE(date) as fecha, COUNT(*) as total_dials FROM outbound_call_log WHERE date IS NOT NULL AND YEAR(date) = 2026 GROUP BY DATE(date)"
    dials_df = pd.read_sql(query_dials, conn)

    query_contacts = """
    SELECT DATE(date) as fecha, COUNT(*) as total_contacts FROM outbound_call_log o
    LEFT JOIN support_tables s ON o.disposition COLLATE utf8mb4_unicode_ci = s.disposition COLLATE utf8mb4_unicode_ci
    WHERE o.date IS NOT NULL AND YEAR(o.date) = 2026 AND (s.no_contact IS NULL OR s.no_contact = 0)
    GROUP BY DATE(date)
    """
    contacts_df = pd.read_sql(query_contacts, conn)

    query_net = """
    SELECT DATE(o.date) as fecha, COUNT(*) as net_contacts FROM outbound_call_log o
    JOIN support_tables s ON o.disposition COLLATE utf8mb4_unicode_ci = s.disposition COLLATE utf8mb4_unicode_ci
    WHERE o.date IS NOT NULL AND YEAR(o.date) = 2026 AND s.net_contacts = 1
    GROUP BY DATE(o.date)
    """
    net_df = pd.read_sql(query_net, conn)

    query_pres = """
    SELECT DATE(date) as fecha, COUNT(*) as total_presentations FROM outbound_call_log o
    JOIN support_tables s ON o.disposition COLLATE utf8mb4_unicode_ci = s.disposition COLLATE utf8mb4_unicode_ci
    WHERE o.date IS NOT NULL AND YEAR(o.date) = 2026 AND s.presentations = 1
    GROUP BY DATE(date)
    """
    pres_df = pd.read_sql(query_pres, conn)

    query_cons = """
    SELECT DATE(date) as fecha, COUNT(*) as total_consultations FROM outbound_call_log o
    JOIN support_tables s ON o.disposition COLLATE utf8mb4_unicode_ci = s.disposition COLLATE utf8mb4_unicode_ci
    WHERE o.date IS NOT NULL AND YEAR(o.date) = 2026 AND s.consultation = 1
    GROUP BY DATE(date)
    """
    cons_df = pd.read_sql(query_cons, conn)

    query_comp = """
    SELECT DATE(date) as fecha, COUNT(*) as agent_completes FROM outbound_call_log o
    JOIN support_tables s ON o.disposition COLLATE utf8mb4_unicode_ci = s.disposition COLLATE utf8mb4_unicode_ci
    WHERE o.date IS NOT NULL AND YEAR(o.date) = 2026 AND s.agent_completes = 1
    GROUP BY DATE(date)
    """
    comp_df = pd.read_sql(query_comp, conn)

    query_unw = """
    SELECT DATE(date) as fecha, COUNT(*) as unworkables FROM outbound_call_log o
    JOIN support_tables s ON o.disposition COLLATE utf8mb4_unicode_ci = s.disposition COLLATE utf8mb4_unicode_ci
    WHERE o.date IS NOT NULL AND YEAR(o.date) = 2026 AND s.unworkable = 1
    GROUP BY DATE(date)
    """
    unw_df = pd.read_sql(query_unw, conn)

    result = []
    for _, rh in hours_df.iterrows():
        fecha = rh['fecha']
        sh = float(rh['spanish_hours'] or 0)
        eh = float(rh['english_hours'] or 0)
        th = sh + eh
        td = int(dials_df[dials_df['fecha'] == fecha].iloc[0]['total_dials']) if not dials_df[dials_df['fecha'] == fecha].empty else 0
        tc = int(contacts_df[contacts_df['fecha'] == fecha].iloc[0]['total_contacts']) if not contacts_df[contacts_df['fecha'] == fecha].empty else 0
        nc = int(net_df[net_df['fecha'] == fecha].iloc[0]['net_contacts']) if not net_df[net_df['fecha'] == fecha].empty else 0
        tp = int(pres_df[pres_df['fecha'] == fecha].iloc[0]['total_presentations']) if not pres_df[pres_df['fecha'] == fecha].empty else 0
        tcs = int(cons_df[cons_df['fecha'] == fecha].iloc[0]['total_consultations']) if not cons_df[cons_df['fecha'] == fecha].empty else 0
        ac = int(comp_df[comp_df['fecha'] == fecha].iloc[0]['agent_completes']) if not comp_df[comp_df['fecha'] == fecha].empty else 0
        uw = int(unw_df[unw_df['fecha'] == fecha].iloc[0]['unworkables']) if not unw_df[unw_df['fecha'] == fecha].empty else 0
        result.append({
            'fecha': fecha, 'spanish_hours': sh, 'english_hours': eh, 'total_hours': th,
            'total_dials': td, 'total_contacts': tc, 'net_contacts': nc,
            'total_presentations': tp, 'total_consultations': tcs,
            'agent_completes': ac, 'total_completes': ac, 'completes_wo_eoc': ac,
            'unworkables': uw
        })
    return result


def get_campaign_data(conn, st_map):
    query = """
    SELECT DATE(o.date) as fecha, COALESCE(o.lead_list_name, 'Unknown') as list_name,
        COUNT(*) as total_dials,
        COUNT(CASE WHEN s.no_contact IS NULL OR s.no_contact = 0 THEN 1 END) as total_contacts,
        COUNT(CASE WHEN s.net_contacts = 1 THEN 1 END) as net_contacts,
        COUNT(CASE WHEN s.presentations = 1 THEN 1 END) as presentations,
        COUNT(CASE WHEN s.consultation = 1 THEN 1 END) as consultations,
        COUNT(CASE WHEN s.agent_completes = 1 THEN 1 END) as agent_completes,
        COUNT(CASE WHEN s.unworkable = 1 THEN 1 END) as unworkables
    FROM outbound_call_log o
    LEFT JOIN support_tables s ON o.disposition COLLATE utf8mb4_unicode_ci = s.disposition COLLATE utf8mb4_unicode_ci
    WHERE o.date IS NOT NULL AND YEAR(o.date) = 2026
    GROUP BY DATE(o.date), o.lead_list_name ORDER BY fecha DESC, list_name
    """
    df = pd.read_sql(query, conn)
    query_hours = """
    SELECT DATE(date) as fecha,
        SUM(CASE WHEN skill_availability LIKE '%%PR%%' THEN TIME_TO_SEC(login_time) ELSE 0 END) / 3600.0 as spanish_hours,
        SUM(CASE WHEN skill_availability LIKE '%%VI%%' THEN TIME_TO_SEC(login_time) ELSE 0 END) / 3600.0 as english_hours
    FROM login_logout WHERE date IS NOT NULL GROUP BY DATE(date)
    """
    hours_df = pd.read_sql(query_hours, conn)
    hours_map = {}
    for _, hr in hours_df.iterrows():
        hours_map[hr['fecha']] = {'spanish': float(hr['spanish_hours'] or 0), 'english': float(hr['english_hours'] or 0)}
    for idx, row in df.iterrows():
        fecha = row['fecha']
        total_sp = hours_map.get(fecha, {}).get('spanish', 0)
        total_en = hours_map.get(fecha, {}).get('english', 0)
        total_d = df[df['fecha'] == fecha]['total_dials'].sum()
        if total_d > 0:
            ratio = row['total_dials'] / total_d
            df.loc[idx, 'spanish_hours'] = total_sp * ratio
            df.loc[idx, 'english_hours'] = total_en * ratio
            df.loc[idx, 'total_hours'] = (total_sp + total_en) * ratio
        else:
            df.loc[idx, 'spanish_hours'] = 0
            df.loc[idx, 'english_hours'] = 0
            df.loc[idx, 'total_hours'] = 0
    return df


def write_summary_sheet(workbook, summary_data):
    ws = workbook.add_worksheet('Summary')
    date_fmt = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    num_fmt = workbook.add_format({'num_format': '0.00'})
    pct_fmt = workbook.add_format({'num_format': '0.00%'})
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': '#FFFFFF', 'border': 1, 'text_wrap': True})

    headers = ['Date Range', 'Total Hours', 'Spanish Hours', 'English Hours', 'Total Dials',
               'DPH', 'Total Contacts', 'Net Contacts', 'CPH', 'Net CPH', 'Contact Rate',
               'Total Presentations', 'PPH', 'Pres to Total', 'Pres to Net',
               'Total Consultations', 'Consult/Hour', 'Consult to Total', 'Consult to Net',
               'Agent Completes', 'Agent Comp/Hour', 'Total Completes', 'Total Comp/Hour',
               'Completes W/O EOC', 'Comp WO EOC/Hour', 'Unworkables', 'Unworkable Rate']
    for col, h in enumerate(headers):
        ws.write(0, col, h, header_fmt)

    summary_sorted = sorted(summary_data, key=lambda x: x['fecha'])
    for idx, rec in enumerate(summary_sorted):
        r = idx + 1
        ws.write_datetime(r, 0, datetime(rec['fecha'].year, rec['fecha'].month, rec['fecha'].day), date_fmt)
        ws.write(r, 2, rec['spanish_hours'], num_fmt)
        ws.write(r, 3, rec['english_hours'], num_fmt)
        ws.write(r, 1, f'=C{r+1}+D{r+1}', num_fmt)
        ws.write(r, 4, rec['total_dials'])
        ws.write(r, 5, f'=E{r+1}/B{r+1}', num_fmt)
        ws.write(r, 6, rec['total_contacts'])
        ws.write(r, 7, rec['net_contacts'])
        ws.write(r, 8, f'=G{r+1}/B{r+1}', num_fmt)
        ws.write(r, 9, f'=H{r+1}/B{r+1}', num_fmt)
        ws.write(r, 10, f'=H{r+1}/E{r+1}', pct_fmt)
        ws.write(r, 11, rec['total_presentations'])
        ws.write(r, 12, f'=L{r+1}/B{r+1}', num_fmt)
        ws.write(r, 13, f'=L{r+1}/G{r+1}', pct_fmt)
        ws.write(r, 14, f'=L{r+1}/H{r+1}', pct_fmt)
        ws.write(r, 15, rec['total_consultations'])
        ws.write(r, 16, f'=P{r+1}/B{r+1}', num_fmt)
        ws.write(r, 17, f'=P{r+1}/G{r+1}', pct_fmt)
        ws.write(r, 18, f'=P{r+1}/H{r+1}', pct_fmt)
        ws.write(r, 19, rec['agent_completes'])
        ws.write(r, 20, f'=T{r+1}/B{r+1}', num_fmt)
        ws.write(r, 21, rec['total_completes'])
        ws.write(r, 22, f'=V{r+1}/B{r+1}', num_fmt)
        ws.write(r, 23, rec['completes_wo_eoc'])
        ws.write(r, 24, f'=X{r+1}/B{r+1}', num_fmt)
        ws.write(r, 25, rec['unworkables'])
        ws.write(r, 26, f'=Z{r+1}/X{r+1}', pct_fmt)

    last_data = len(summary_sorted)
    from collections import defaultdict
    month_data = defaultdict(list)
    for idx, rec in enumerate(summary_sorted):
        fecha = rec['fecha']
        if isinstance(fecha, str):
            fecha = datetime.strptime(fecha, '%Y-%m-%d')
        month_data[fecha.strftime('%B')].append(idx + 1)

    month_order = ['January', 'February', 'March', 'April', 'May', 'June',
                   'July', 'August', 'September', 'October', 'November', 'December']
    main_year = summary_sorted[0]['fecha'].year if summary_sorted else 2026
    cr = last_data + 2
    monthly_rows = {}

    for month in month_order:
        if month in month_data and month_data[month]:
            rows = month_data[month]
            s, e = min(rows), max(rows)
            ws.write(cr, 0, month)
            ws.write(cr, 1, f'=SUM(C{cr+1}:D{cr+1})', num_fmt)
            ws.write(cr, 2, f'=SUM(C{s+1}:C{e+1})', num_fmt)
            ws.write(cr, 3, f'=SUM(D{s+1}:D{e+1})', num_fmt)
            ws.write(cr, 4, f'=SUM(E{s+1}:E{e+1})')
            ws.write(cr, 5, f'=E{cr+1}/B{cr+1}', num_fmt)
            ws.write(cr, 6, f'=SUM(G{s+1}:G{e+1})')
            ws.write(cr, 7, f'=SUM(H{s+1}:H{e+1})')
            ws.write(cr, 8, f'=G{cr+1}/B{cr+1}', num_fmt)
            ws.write(cr, 9, f'=H{cr+1}/B{cr+1}', num_fmt)
            ws.write(cr, 10, f'=H{cr+1}/E{cr+1}', pct_fmt)
            ws.write(cr, 11, f'=SUM(L{s+1}:L{e+1})')
            ws.write(cr, 12, f'=L{cr+1}/B{cr+1}', num_fmt)
            ws.write(cr, 13, f'=L{cr+1}/G{cr+1}', pct_fmt)
            ws.write(cr, 14, f'=L{cr+1}/H{cr+1}', pct_fmt)
            ws.write(cr, 15, f'=SUM(P{s+1}:P{e+1})')
            ws.write(cr, 16, f'=P{cr+1}/B{cr+1}', num_fmt)
            ws.write(cr, 17, f'=P{cr+1}/G{cr+1}', pct_fmt)
            ws.write(cr, 18, f'=P{cr+1}/H{cr+1}', pct_fmt)
            ws.write(cr, 19, f'=SUM(T{s+1}:T{e+1})')
            ws.write(cr, 20, f'=T{cr+1}/B{cr+1}', num_fmt)
            ws.write(cr, 21, f'=SUM(V{s+1}:V{e+1})')
            ws.write(cr, 22, f'=V{cr+1}/B{cr+1}', num_fmt)
            ws.write(cr, 23, f'=SUM(X{s+1}:X{e+1})')
            ws.write(cr, 24, f'=X{cr+1}/B{cr+1}', num_fmt)
            ws.write(cr, 25, f'=SUM(Z{s+1}:Z{e+1})')
            ws.write(cr, 26, f'=Z{cr+1}/X{cr+1}', pct_fmt)
            monthly_rows[month] = cr
            cr += 1

    quarters = {'Q1': ['January', 'February', 'March'], 'Q2': ['April', 'May', 'June'],
                'Q3': ['July', 'August', 'September'], 'Q4': ['October', 'November', 'December']}
    for q, months in quarters.items():
        if all(m in monthly_rows for m in months):
            fs, fe = monthly_rows[months[0]], monthly_rows[months[-1]]
            ws.write(cr, 0, q)
            ws.write(cr, 1, f'=SUM(C{cr+1}:D{cr+1})', num_fmt)
            ws.write(cr, 2, f'=SUM(C{fs+1}:C{fe+1})', num_fmt)
            ws.write(cr, 3, f'=SUM(D{fs+1}:D{fe+1})', num_fmt)
            ws.write(cr, 4, f'=SUM(E{fs+1}:E{fe+1})')
            ws.write(cr, 5, f'=E{cr+1}/B{cr+1}', num_fmt)
            ws.write(cr, 6, f'=SUM(G{fs+1}:G{fe+1})')
            ws.write(cr, 7, f'=SUM(H{fs+1}:H{fe+1})')
            ws.write(cr, 8, f'=G{cr+1}/B{cr+1}', num_fmt)
            ws.write(cr, 9, f'=H{cr+1}/B{cr+1}', num_fmt)
            ws.write(cr, 10, f'=H{cr+1}/E{cr+1}', pct_fmt)
            ws.write(cr, 11, f'=SUM(L{fs+1}:L{fe+1})')
            ws.write(cr, 12, f'=L{cr+1}/B{cr+1}', num_fmt)
            ws.write(cr, 13, f'=L{cr+1}/G{cr+1}', pct_fmt)
            ws.write(cr, 14, f'=L{cr+1}/H{cr+1}', pct_fmt)
            ws.write(cr, 15, f'=SUM(P{fs+1}:P{fe+1})')
            ws.write(cr, 16, f'=P{cr+1}/B{cr+1}', num_fmt)
            ws.write(cr, 17, f'=P{cr+1}/G{cr+1}', pct_fmt)
            ws.write(cr, 18, f'=P{cr+1}/H{cr+1}', pct_fmt)
            ws.write(cr, 19, f'=SUM(T{fs+1}:T{fe+1})')
            ws.write(cr, 20, f'=T{cr+1}/B{cr+1}', num_fmt)
            ws.write(cr, 21, f'=SUM(V{fs+1}:V{fe+1})')
            ws.write(cr, 22, f'=V{cr+1}/B{cr+1}', num_fmt)
            ws.write(cr, 23, f'=SUM(X{fs+1}:X{fe+1})')
            ws.write(cr, 24, f'=X{cr+1}/B{cr+1}', num_fmt)
            ws.write(cr, 25, f'=SUM(Z{fs+1}:Z{fe+1})')
            ws.write(cr, 26, f'=Z{cr+1}/X{cr+1}', pct_fmt)
            cr += 1

    cr += 1
    if monthly_rows:
        fml, lml = list(monthly_rows.values())[0], list(monthly_rows.values())[-1]
        ws.write(cr, 0, f'{main_year} YTD')
        ws.write(cr, 1, f'=SUM(C{cr+1},D{cr+1})', num_fmt)
        ws.write(cr, 2, f'=SUM(C{fml+1}:C{lml+1})', num_fmt)
        ws.write(cr, 3, f'=SUM(D{fml+1}:D{lml+1})', num_fmt)
        ws.write(cr, 4, f'=SUM(E{fml+1}:E{lml+1})')
        ws.write(cr, 5, f'=E{cr+1}/B{cr+1}', num_fmt)
        ws.write(cr, 6, f'=SUM(G{fml+1}:G{lml+1})')
        ws.write(cr, 7, f'=SUM(H{fml+1}:H{lml+1})')
        ws.write(cr, 8, f'=G{cr+1}/B{cr+1}', num_fmt)
        ws.write(cr, 9, f'=H{cr+1}/B{cr+1}', num_fmt)
        ws.write(cr, 10, f'=H{cr+1}/E{cr+1}', pct_fmt)
        ws.write(cr, 11, f'=SUM(L{fml+1}:L{lml+1})')
        ws.write(cr, 12, f'=L{cr+1}/B{cr+1}', num_fmt)
        ws.write(cr, 13, f'=L{cr+1}/G{cr+1}', pct_fmt)
        ws.write(cr, 14, f'=L{cr+1}/H{cr+1}', pct_fmt)
        ws.write(cr, 15, f'=SUM(P{fml+1}:P{lml+1})')
        ws.write(cr, 16, f'=P{cr+1}/B{cr+1}', num_fmt)
        ws.write(cr, 17, f'=P{cr+1}/G{cr+1}', pct_fmt)
        ws.write(cr, 18, f'=P{cr+1}/H{cr+1}', pct_fmt)
        ws.write(cr, 19, f'=SUM(T{fml+1}:T{lml+1})')
        ws.write(cr, 20, f'=T{cr+1}/B{cr+1}', num_fmt)
        ws.write(cr, 21, f'=SUM(V{fml+1}:V{lml+1})')
        ws.write(cr, 22, f'=V{cr+1}/B{cr+1}', num_fmt)
        ws.write(cr, 23, f'=SUM(X{fml+1}:X{lml+1})')
        ws.write(cr, 24, f'=X{cr+1}/B{cr+1}', num_fmt)
        ws.write(cr, 25, f'=SUM(Z{fml+1}:Z{lml+1})')
        ws.write(cr, 26, f'=Z{cr+1}/X{cr+1}', pct_fmt)
        ytd_row = cr

        cr += 1
        ws.write(cr, 0, 'Penetration Rate - YTD')
        ws.write(cr, 1, f'=V{ytd_row+1}/(B{cr+4}-Z{ytd_row+1})', pct_fmt)
        cr += 1
        ws.write(cr, 0, 'Net Contact to Callable Leads - YTD')
        ws.write(cr, 1, f'=H{ytd_row+1}/B{cr+3}', pct_fmt)
        cr += 1
        ws.write(cr, 0, 'Total Callable Leads Received - YTD')
        ws.write(cr, 1, 21781)
        cr += 1
        ws.write(cr, 0, 'Total Workable Leads - YTD')
        ws.write(cr, 1, f'=B{cr}-Z{ytd_row+1}')

    return last_data


def write_campaign_sheet(workbook, campaign_name, campaign_data, summary_data, data_start_row, callable_leads=0):
    safe_name = campaign_name[:31]
    ws = workbook.add_worksheet(safe_name)
    date_fmt = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    num_fmt = workbook.add_format({'num_format': '0.00'})
    pct_fmt = workbook.add_format({'num_format': '0.00%'})
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': '#FFFFFF', 'border': 1})
    label_fmt = workbook.add_format({'bold': True})

    headers = ['Date Range', 'Total Hours', 'Spanish Hours', 'English Hours', 'Total Dials',
               'DPH', 'Total Contacts', 'Net Contacts', 'CPH', 'Net CPH', 'Contact Rate',
               'Total Presentations', 'PPH', 'Pres to Total', 'Pres to Net',
               'Total Consultations', 'Consult/Hour', 'Consult to Total', 'Consult to Net',
               'Agent Completes', 'Agent Comp/Hour', 'Total Completes', 'Total Comp/Hour',
               'Completes W/O EOC', 'Comp WO EOC/Hour', 'Unworkables', 'Unworkable Rate']
    ws.write(0, 0, campaign_name, label_fmt)
    for col, h in enumerate(headers):
        ws.write(1, col, h, header_fmt)

    cr = data_start_row
    if not campaign_data.empty:
        for fecha in sorted(campaign_data['fecha'].unique()):
            fd = campaign_data[campaign_data['fecha'] == fecha].iloc[0]
            th = float(fd.get('total_hours', 0) or 0)
            sh = float(fd.get('spanish_hours', 0) or 0)
            eh = float(fd.get('english_hours', 0) or 0)
            td = int(fd.get('total_dials', 0) or 0)
            tc = int(fd.get('total_contacts', 0) or 0)
            nc = int(fd.get('net_contacts', 0) or 0)
            tp = int(fd.get('presentations', 0) or 0)
            tcs = int(fd.get('consultations', 0) or 0)
            ac = int(fd.get('agent_completes', 0) or 0)
            uw = int(fd.get('unworkables', 0) or 0)

            r = cr
            ws.write_datetime(r, 0, datetime(fecha.year, fecha.month, fecha.day), date_fmt)
            ws.write(r, 2, sh, num_fmt)
            ws.write(r, 3, eh, num_fmt)
            ws.write(r, 1, f'=C{r+1}+D{r+1}', num_fmt)
            ws.write(r, 4, td)
            ws.write(r, 5, f'=E{r+1}/B{r+1}', num_fmt)
            ws.write(r, 6, tc)
            ws.write(r, 7, nc)
            ws.write(r, 8, f'=G{r+1}/B{r+1}', num_fmt)
            ws.write(r, 9, f'=H{r+1}/B{r+1}', num_fmt)
            ws.write(r, 10, f'=H{r+1}/E{r+1}', pct_fmt)
            ws.write(r, 11, tp)
            ws.write(r, 12, f'=L{r+1}/B{r+1}', num_fmt)
            ws.write(r, 13, f'=L{r+1}/G{r+1}', pct_fmt)
            ws.write(r, 14, f'=L{r+1}/H{r+1}', pct_fmt)
            ws.write(r, 15, tcs)
            ws.write(r, 16, f'=P{r+1}/B{r+1}', num_fmt)
            ws.write(r, 17, f'=P{r+1}/G{r+1}', pct_fmt)
            ws.write(r, 18, f'=P{r+1}/H{r+1}', pct_fmt)
            ws.write(r, 19, ac)
            ws.write(r, 20, f'=T{r+1}/B{r+1}', num_fmt)
            ws.write(r, 21, f'=T{r+1}')
            ws.write(r, 22, f'=V{r+1}/B{r+1}', num_fmt)
            ws.write(r, 23, f'=T{r+1}')
            ws.write(r, 24, f'=X{r+1}/B{r+1}', num_fmt)
            ws.write(r, 25, uw)
            ws.write(r, 26, f'=Z{r+1}/X{r+1}', pct_fmt)
            cr += 1

    sep = cr
    ws.write(cr, 0, 'Summary', label_fmt)
    ws.write(cr, 2, f'=SUM(C{data_start_row+1}:C{sep})', num_fmt)
    ws.write(cr, 3, f'=SUM(D{data_start_row+1}:D{sep})', num_fmt)
    ws.write(cr, 1, f'=C{cr+1}+D{cr+1}', num_fmt)
    ws.write(cr, 4, f'=SUM(E{data_start_row+1}:E{sep})')
    ws.write(cr, 5, f'=E{cr+1}/B{cr+1}', num_fmt)
    ws.write(cr, 6, f'=SUM(G{data_start_row+1}:G{sep})')
    ws.write(cr, 7, f'=SUM(H{data_start_row+1}:H{sep})')
    ws.write(cr, 8, f'=G{cr+1}/B{cr+1}', num_fmt)
    ws.write(cr, 9, f'=H{cr+1}/B{cr+1}', num_fmt)
    ws.write(cr, 10, f'=H{cr+1}/E{cr+1}', pct_fmt)
    ws.write(cr, 11, f'=SUM(L{data_start_row+1}:L{sep})')
    ws.write(cr, 12, f'=L{cr+1}/B{cr+1}', num_fmt)
    ws.write(cr, 13, f'=L{cr+1}/G{cr+1}', pct_fmt)
    ws.write(cr, 14, f'=L{cr+1}/H{cr+1}', pct_fmt)
    ws.write(cr, 15, f'=SUM(P{data_start_row+1}:P{sep})')
    ws.write(cr, 16, f'=P{cr+1}/B{cr+1}', num_fmt)
    ws.write(cr, 17, f'=P{cr+1}/G{cr+1}', pct_fmt)
    ws.write(cr, 18, f'=P{cr+1}/H{cr+1}', pct_fmt)
    ws.write(cr, 19, f'=SUM(T{data_start_row+1}:T{sep})')
    ws.write(cr, 20, f'=T{cr+1}/B{cr+1}', num_fmt)
    ws.write(cr, 21, f'=SUM(V{data_start_row+1}:V{sep})')
    ws.write(cr, 22, f'=V{cr+1}/B{cr+1}', num_fmt)
    ws.write(cr, 23, f'=SUM(X{data_start_row+1}:X{sep})')
    ws.write(cr, 24, f'=X{cr+1}/B{cr+1}', num_fmt)
    ws.write(cr, 25, f'=SUM(Z{data_start_row+1}:Z{sep})')
    ws.write(cr, 26, f'=Z{cr+1}/X{cr+1}', pct_fmt)
    summary_row = cr

    cr += 2
    ws.write(cr, 0, 'Penetration Rate - YTD', label_fmt)
    ws.write(cr, 1, f'=V{summary_row+1}/(B{cr+4}-Z{summary_row+1})', pct_fmt)
    cr += 1
    ws.write(cr, 0, 'Net Contact to Callable Leads - YTD', label_fmt)
    ws.write(cr, 1, f'=H{summary_row+1}/B{cr+3}', pct_fmt)
    cr += 1
    ws.write(cr, 0, 'Total Callable Leads Received - YTD', label_fmt)
    callable_row = cr
    ws.write(cr, 1, callable_leads)
    cr += 1
    ws.write(cr, 0, 'Total Workable Leads - YTD', label_fmt)
    workable_row = cr
    ws.write(cr, 1, f'=B{cr}-Z{summary_row+1}')

    return summary_row, callable_row, workable_row


def write_campaign_summary(workbook, campaigns_info):
    ws = workbook.add_worksheet('Campaign Summary')
    label_fmt = workbook.add_format({'bold': True})
    pct_fmt = workbook.add_format({'num_format': '0.00%'})
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': '#FFFFFF', 'border': 1})

    seg_headers = ['Target Segmentation Roll up', 'Penetration Rate', 'Net Contact Rate', 'Total Callable Leads', 'Total Workable Leads']
    for col, h in enumerate(seg_headers):
        ws.write(0, col + 7, h, header_fmt)

    ws.write(0, 0, 'Campaign', header_fmt)
    ws.write(0, 1, 'Penetration Rate', header_fmt)
    ws.write(0, 2, 'Net Contact Rate', header_fmt)
    ws.write(0, 3, 'Total Callable Leads', header_fmt)
    ws.write(0, 4, 'Total Workable Leads', header_fmt)
    ws.write(0, 5, 'Segmentation', header_fmt)

    seg_campaigns = {}
    data_row = 1
    for info in campaigns_info:
        campaign = info['campaign']
        seg = info['seg']

        ws.write(data_row, 0, campaign)
        ws.write(data_row, 1, f"='{campaign}'!B{info['summary_row']+3}", pct_fmt)
        ws.write(data_row, 2, f"='{campaign}'!H{info['summary_row']+4}", pct_fmt)
        ws.write(data_row, 3, f"='{campaign}'!B{info['callable_row']+1}")
        ws.write(data_row, 4, f"='{campaign}'!B{info['workable_row']+1}")
        ws.write(data_row, 5, seg)

        if seg not in seg_campaigns:
            seg_campaigns[seg] = []
        seg_campaigns[seg].append(data_row)
        data_row += 1

    seg_row = 2
    for seg_name, rows in seg_campaigns.items():
        ws.write(seg_row, 7, seg_name, label_fmt)

        completes_parts = []
        net_contacts_parts = []
        col_d_refs = []
        col_e_refs = []

        for r in rows:
            campaign_name = campaigns_info[r - 1]['campaign']
            summary_row = campaigns_info[r - 1]['summary_row']
            completes_parts.append(f"'{campaign_name}'!V{summary_row+1}")
            net_contacts_parts.append(f"'{campaign_name}'!H{summary_row+1}")
            col_d_refs.append(f"D{r+1}")
            col_e_refs.append(f"E{r+1}")

        sum_completes = ",".join(completes_parts)
        sum_net_contacts = ",".join(net_contacts_parts)
        col_d_range = ",".join(col_d_refs)
        col_e_range = ",".join(col_e_refs)

        ws.write(seg_row, 8, f"=SUM({sum_completes})/SUM({col_e_range})", pct_fmt)
        ws.write(seg_row, 9, f"=SUM({sum_net_contacts})/SUM({col_e_range})", pct_fmt)
        ws.write(seg_row, 10, f"=SUM({col_d_range})")
        ws.write(seg_row, 11, f"=SUM({col_e_range})")

        seg_row += 1


def generate_daily_campaign_results_v3():
    print("=== Generando Daily_Campaign_Results.xlsx (V3) ===\n")
    conn = None
    try:
        conn = get_connection()
        st_map = get_support_tables_map(conn)
        summary_data = get_summary_data(conn, st_map)
        print(f"  Summary: {len(summary_data)} dias")
        campaign_df = get_campaign_data(conn, st_map)
        campaign_names = sorted(campaign_df['list_name'].unique())
        print(f"  Campañas: {len(campaign_names)} campañas únicas\n")
        list_name_records_map = get_list_name_records_map(conn)

        output_path = config.OUTPUT_FILES['daily_campaign']
        output_dir = os.path.dirname(output_path)
        os.makedirs(output_dir, exist_ok=True)

        workbook = pd.ExcelWriter(output_path, engine='xlsxwriter')
        writerWorkbook = workbook.book

        last_data_row = write_summary_sheet(writerWorkbook, summary_data)

        campaigns_info = []
        for campaign in campaign_names:
            cd = campaign_df[campaign_df['list_name'] == campaign]
            callable_leads = int(list_name_records_map.get(campaign, 0))
            print(f"  {campaign}: callable_leads={callable_leads}")
            summary_row, callable_row, workable_row = write_campaign_sheet(writerWorkbook, campaign, cd, summary_data, 2, callable_leads)
            campaigns_info.append({
                'campaign': campaign,
                'seg': next((v for k, v in {'TRENDING': 'Trending Inactive', 'INACT': 'Inactive', 'NYA': 'Never Active', 'NS': 'New Signings', 'ACT': 'Active'}.items() if k in campaign), 'Unknown'),
                'summary_row': summary_row,
                'callable_row': callable_row,
                'workable_row': workable_row,
            })

        write_campaign_summary(writerWorkbook, campaigns_info)

        workbook.close()
        print(f"\n[OK] Excel generado: {output_path}")
        return str(output_path)

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        raise
    finally:
        if conn:
            conn.close()


if __name__ == '__main__':
    generate_daily_campaign_results_v3()
