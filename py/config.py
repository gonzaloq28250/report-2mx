from pathlib import Path
import os

BASE_DIR = Path(__file__).resolve().parent.parent

TEMPLATE_DIR = BASE_DIR / 'templates'
REPORTS_DIR = BASE_DIR / 'reports'

TEMPLATE_DIR.mkdir(exist_ok=True)
REPORTS_DIR.mkdir(exist_ok=True)

TEMPLATES = {
    'calls_consolidate': TEMPLATE_DIR / 'Calls_Consolidate_Report_2026.xlsx',
    'visit_log_optblue': TEMPLATE_DIR / 'Visit_Log_Report(Optblue)_2026.xlsx',
    'call_disposition': TEMPLATE_DIR / 'Call_Disposition_Report_2026.xlsx',
    'visits_disposition': TEMPLATE_DIR / 'Visits_Disposition_Report_2026.xlsx',
    'pr_calls_duration': TEMPLATE_DIR / 'PR Calls Duration Report_2026.xlsx',
    'pr_escalation': TEMPLATE_DIR / 'PR Calls and Visits Escalation Tracker_Template.xlsx',
    'daily_campaign': TEMPLATE_DIR / 'Daily-Campaign-Results-v2.xlsx',
    'service_level': TEMPLATE_DIR / 'Service-Level-Report.xlsx',
}

OUTPUT_FILES = {
    'calls_consolidate': REPORTS_DIR / 'Calls_Consolidate_Report_2026.xlsx',
    'visit_log_optblue': REPORTS_DIR / 'Visit_Log_Report_Optblue_2026.xlsx',
    'call_disposition': REPORTS_DIR / 'Call_Disposition_Report_2026.xlsx',
    'visits_disposition': REPORTS_DIR / 'Visits_Disposition_Report_2026.xlsx',
    'pr_calls_duration': REPORTS_DIR / 'PR_Calls_Duration_Report_2026.xlsx',
    'pr_escalation': REPORTS_DIR / 'PR_Calls_and_Visits_Escalation_2026.xlsx',
    'daily_campaign': REPORTS_DIR / 'Daily_Campaign_Results_2026.xlsx',
    'service_level': REPORTS_DIR / 'Service_Level_Report_2026.xlsx',
    'optouts_dnc': REPORTS_DIR / 'OPTBLUE_INSIGHT_OPTOUTS.txt',
}

MYSQL_CONFIG = {
    'host': os.environ.get('ICC_MYSQL_HOST', 'icqdbmysqlreports.mysql.database.azure.com'),
    'user': os.environ.get('ICC_MYSQL_USER', 'gonzaloq'),
    'password': os.environ.get('ICC_MYSQL_PASSWORD', '73ch$iCC'),
    'database': os.environ.get('ICC_MYSQL_DB', 'icc-amex'),
    'charset': 'utf8mb4',
}
