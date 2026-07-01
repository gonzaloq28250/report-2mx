import datetime
import logging
import sys
from decimal import Decimal
from pathlib import Path

import mysql.connector

from odoo import models, fields, api, _
from odoo.exceptions import UserError

_logger = logging.getLogger(__name__)

_PY_DIR = Path(__file__).resolve().parent.parent.parent / 'py'
if str(_PY_DIR) not in sys.path:
    sys.path.insert(0, str(_PY_DIR))

try:
    from config import MYSQL_CONFIG
except ImportError:
    MYSQL_CONFIG = {
        'host': 'icqdbmysqlreports.mysql.database.azure.com',
        'user': 'gonzaloq',
        'password': '73ch$iCC',
        'database': 'icc-amex',
        'charset': 'utf8mb4',
    }


def _get_mysql():
    conn = mysql.connector.connect(**MYSQL_CONFIG)
    cur = conn.cursor(dictionary=True)
    cur.execute("SET time_zone = '-04:00'")
    cur.close()
    return conn


def _row_to_vals(row, model_fields):
    vals = {}
    for fname, v in row.items():
        if fname not in model_fields:
            continue
        if isinstance(v, Decimal):
            v = float(v)
        if isinstance(v, str):
            field = model_fields[fname]
            if field.type == 'date':
                for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%m/%d/%Y', '%m-%d-%Y'):
                    try:
                        v = datetime.datetime.strptime(v, fmt).date()
                        break
                    except (ValueError, TypeError):
                        continue
        if v is None or isinstance(v, (int, float, str, datetime.date, datetime.datetime)):
            vals[fname] = v
    return vals


# ============================================================
# 1. Service Level Metrics
# ============================================================

class IccServiceLevelMetric(models.Model):
    _name = 'icc.service.level.metric'
    _description = 'Service Level Metrics by Date'
    _order = 'fecha desc'
    _rec_name = 'fecha'

    fecha = fields.Date(string='Date', required=True)
    total_llamadas = fields.Integer(string='Total Calls')
    agentes_unicos = fields.Integer(string='Unique Agents')
    llamadas_con_talk_time = fields.Integer(string='Calls with Talk Time')
    avg_talk_time_segundos = fields.Float(string='Avg Talk Time (s)')
    avg_acw_segundos = fields.Float(string='Avg ACW (s)')
    avg_handle_time_segundos = fields.Float(string='Avg Handle Time (s)')
    production_hours = fields.Float(string='Production Hours')

    def action_sync_from_mysql(self):
        conn = _get_mysql()
        try:
            cur = conn.cursor(dictionary=True)
            cur.execute("""
                SELECT
                    DATE(o.date) as fecha,
                    COUNT(*) as total_llamadas,
                    COUNT(DISTINCT o.agent_name) as agentes_unicos,
                    SUM(CASE WHEN o.talk_time IS NOT NULL THEN 1 ELSE 0 END) as llamadas_con_talk_time,
                    AVG(CASE WHEN o.talk_time IS NOT NULL AND o.talk_time != '00:00:00'
                        THEN TIME_TO_SEC(o.talk_time) ELSE NULL END) as avg_talk_time_segundos,
                    AVG(CASE WHEN o.after_call_work_time IS NOT NULL AND o.after_call_work_time != '00:00:00'
                        THEN TIME_TO_SEC(o.after_call_work_time) ELSE NULL END) as avg_acw_segundos,
                    AVG(CASE WHEN o.handle_time IS NOT NULL AND o.handle_time != '00:00:00'
                        THEN TIME_TO_SEC(o.handle_time) ELSE NULL END) as avg_handle_time_segundos,
                    COALESCE(MAX(l.production_hours), 0) as production_hours
                FROM outbound_call_log o
                LEFT JOIN (
                    SELECT DATE(date) as ldate, SUM(TIME_TO_SEC(login_time)) / 3600.0 as production_hours
                    FROM login_logout
                    WHERE date IS NOT NULL AND login_time IS NOT NULL
                    GROUP BY DATE(date)
                ) l ON DATE(o.date) = l.ldate
                WHERE o.date IS NOT NULL AND YEAR(o.date) = 2026
                GROUP BY DATE(o.date)
                ORDER BY fecha DESC
            """)
            rows = cur.fetchall()
            self.search([]).unlink()
            for r in rows:
                vals = _row_to_vals(r, self._fields)
                self.create(vals)
            cur.close()
            action = self.env.ref('odoo_reports.action_icc_service_level_metric')
            if action:
                return action.read()[0]
        finally:
            conn.close()


# ============================================================
# 2. Daily Campaign Summary
# ============================================================

class IccDailyCampaignSummary(models.Model):
    _name = 'icc.daily.campaign.summary'
    _description = 'Daily Campaign Summary'
    _order = 'fecha desc'
    _rec_name = 'fecha'

    fecha = fields.Date(string='Date', required=True)
    spanish_hours = fields.Float(string='Spanish Hours')
    english_hours = fields.Float(string='English Hours')
    total_hours = fields.Float(string='Total Hours')
    total_dials = fields.Integer(string='Total Dials')
    dph = fields.Float(string='Dials Per Hour')
    total_contacts = fields.Integer(string='Total Contacts')
    net_contacts = fields.Integer(string='Net Contacts')
    presentations = fields.Integer(string='Presentations')
    unworkables = fields.Integer(string='Unworkables')
    consultations = fields.Integer(string='Consultations')
    agent_completes = fields.Integer(string='Agent Completes')

    def action_sync_from_mysql(self):
        conn = _get_mysql()
        try:
            cur = conn.cursor(dictionary=True)

            cur.execute("""
                SELECT
                    DATE(date) as fecha,
                    skill_availability,
                    SUM(TIME_TO_SEC(login_time)) / 3600.0 as total_hours
                FROM login_logout
                WHERE date IS NOT NULL
                GROUP BY DATE(date), skill_availability
                ORDER BY fecha DESC
            """)
            login_rows = cur.fetchall()
            login_by_date = {}
            for r in login_rows:
                fd = str(r['fecha'])
                if fd not in login_by_date:
                    login_by_date[fd] = {'spanish': 0.0, 'english': 0.0}
                skill = (r['skill_availability'] or '').upper()
                if 'PR' in skill:
                    login_by_date[fd]['spanish'] += float(r['total_hours'] or 0)
                elif 'VI' in skill:
                    login_by_date[fd]['english'] += float(r['total_hours'] or 0)

            cur.execute("""
                SELECT
                    DATE(o.date) as fecha,
                    COUNT(*) as total_dials,
                    COUNT(CASE WHEN s.net_contacts = 1 THEN 1 END) as net_contacts,
                    COUNT(CASE WHEN s.presentations = 1 THEN 1 END) as presentations,
                    COUNT(CASE WHEN s.unworkable = 1 THEN 1 END) as unworkables,
                    COUNT(CASE WHEN s.consultation = 1 THEN 1 END) as consultations,
                    COUNT(CASE WHEN s.agent_completes = 1 THEN 1 END) as agent_completes
                FROM outbound_call_log o
                LEFT JOIN support_tables s
                    ON o.disposition COLLATE utf8mb4_unicode_ci = s.disposition COLLATE utf8mb4_unicode_ci
                WHERE o.date IS NOT NULL AND YEAR(o.date) = 2026
                GROUP BY DATE(o.date)
                ORDER BY fecha DESC
                LIMIT 120
            """)
            rows = cur.fetchall()

            self.search([]).unlink()
            for r in rows:
                fd = str(r['fecha'])
                l = login_by_date.get(fd, {'spanish': 0.0, 'english': 0.0})
                total_h = l['spanish'] + l['english']
                self.create({
                    'fecha': r['fecha'],
                    'spanish_hours': l['spanish'],
                    'english_hours': l['english'],
                    'total_hours': total_h,
                    'total_dials': r['total_dials'],
                    'dph': round(r['total_dials'] / total_h, 1) if total_h > 0 else 0,
                    'total_contacts': r['total_dials'],
                    'net_contacts': r['net_contacts'],
                    'presentations': r['presentations'],
                    'unworkables': r['unworkables'],
                    'consultations': r['consultations'],
                    'agent_completes': r['agent_completes'],
                })
            cur.close()
            action = self.env.ref('odoo_reports.action_icc_daily_campaign_summary')
            if action:
                return action.read()[0]
        finally:
            conn.close()


# ============================================================
# 3. Calls Consolidate
# ============================================================

class IccCallsConsolidate(models.Model):
    _name = 'icc.calls.consolidate'
    _description = 'Calls Consolidate Report'
    _order = 'entry_datesubmitted desc, id desc'

    mysql_id = fields.Integer(string='MySQL ID', readonly=True)
    llamadasoptbluenew_id = fields.Integer(string='Optblue New ID')
    name = fields.Char(string='Name')
    seller_id_dba_name = fields.Char(string='DBA Name')
    seller_id_city_name = fields.Char(string='City')
    mcc_code = fields.Char(string='MCC Code')
    mcc_description = fields.Char(string='MCC Description')
    broad_industry = fields.Char(string='Industry')
    lead_list_name = fields.Char(string='Campaign')
    target_segmentation_2 = fields.Char(string='Segmentation')
    calls_completed = fields.Char(string='Calls Completed')
    contacted_correct_person = fields.Char(string='Contacted Correct Person')
    disposition = fields.Char(string='Disposition')
    merchant_accept_amex = fields.Char(string='Accepts Amex')
    willing_accept_amex_customers = fields.Char(string='Willing Accept')
    calls_agent = fields.Char(string='Agent')
    entry_status = fields.Char(string='Status')
    entry_datesubmitted = fields.Datetime(string='Date Submitted')

    def action_sync_from_mysql(self):
        conn = _get_mysql()
        try:
            cur = conn.cursor(dictionary=True)
            cur.execute("""
                SELECT * FROM calls_consolidate_report ORDER BY id DESC LIMIT 500
            """)
            rows = cur.fetchall()
            self.search([]).unlink()
            for r in rows:
                vals = _row_to_vals(r, self._fields)
                if 'id' in r:
                    vals['mysql_id'] = r['id']
                self.create(vals)
            cur.close()
            action = self.env.ref('odoo_reports.action_icc_calls_consolidate')
            if action:
                return action.read()[0]
        finally:
            conn.close()


# ============================================================
# 4. Visit Log Optblue
# ============================================================

class IccVisitLogOptblue(models.Model):
    _name = 'icc.visit.log.optblue'
    _description = 'Visit Log Optblue Report'
    _order = 'entry_datesubmitted desc, id desc'

    mysql_id = fields.Integer(string='MySQL ID', readonly=True)
    name = fields.Char(string='Name')
    se_number = fields.Char(string='SE Number')
    country = fields.Char(string='Country')
    sales_channel_name = fields.Char(string='Sales Channel')
    visits_agent = fields.Char(string='Agent')
    phone = fields.Char(string='Phone')
    seller_id = fields.Char(string='Seller ID')
    completed_visit = fields.Char(string='Visit Completed')
    pop_on_entry = fields.Char(string='POP on Entry')
    was_pop_placed = fields.Char(string='POP Placed')
    which_kit_was_delivered = fields.Char(string='Kit Delivered')
    awareness_of_axp_acceptance = fields.Char(string='Awareness of AXP')
    welcome_acceptance = fields.Char(string='Welcome Acceptance')
    terminal_tested = fields.Char(string='Terminal Tested')
    terminal_successfully_working = fields.Char(string='Terminal Working')
    entry_status = fields.Char(string='Status')
    entry_datesubmitted = fields.Datetime(string='Date Submitted')

    def action_sync_from_mysql(self):
        conn = _get_mysql()
        try:
            cur = conn.cursor(dictionary=True)
            cur.execute("""
                SELECT * FROM visit_log_optblue_report ORDER BY id DESC LIMIT 500
            """)
            rows = cur.fetchall()
            self.search([]).unlink()
            for r in rows:
                vals = _row_to_vals(r, self._fields)
                if 'id' in r:
                    vals['mysql_id'] = r['id']
                self.create(vals)
            cur.close()
            action = self.env.ref('odoo_reports.action_icc_visit_log_optblue')
            if action:
                return action.read()[0]
        finally:
            conn.close()


# ============================================================
# 5. Visits Disposition
# ============================================================

class IccVisitsDisposition(models.Model):
    _name = 'icc.visits.disposition'
    _description = 'Visits Disposition Report'
    _order = 'entry_datesubmitted desc, id desc'

    mysql_id = fields.Integer(string='MySQL ID', readonly=True)
    merchantname = fields.Char(string='Merchant Name')
    sell_se_no = fields.Char(string='SE Number')
    merchantcity = fields.Char(string='City')
    type_of_merchant = fields.Char(string='Type')
    agentname = fields.Char(string='Agent')
    what_exterior_pop_seen_organic = fields.Char(string='Exterior POP Seen')
    was_organic_amex_pop_in_good_condition = fields.Char(string='POP Good Condition')
    was_contactless_pop_displayed = fields.Char(string='Contactless POP')
    did_you_speak_with_a_decision_maker = fields.Char(string='Spoke with Decision Maker')
    was_merchant_aware_of_amex_acceptance = fields.Char(string='Aware of Amex')
    did_the_merchant_indicate_they_were_acept = fields.Char(string='Merchant Accepts')
    merchant_gave_permission_to_test = fields.Char(string='Permission to Test')
    result_of_contactless_test = fields.Char(string='Contactless Result')
    acceptance_meter = fields.Char(string='Acceptance Meter')
    entry_status = fields.Char(string='Status')
    entry_datesubmitted = fields.Datetime(string='Date Submitted')

    def action_sync_from_mysql(self):
        conn = _get_mysql()
        try:
            cur = conn.cursor(dictionary=True)
            cur.execute("""
                SELECT * FROM visits_disposition_report ORDER BY id DESC LIMIT 500
            """)
            rows = cur.fetchall()
            self.search([]).unlink()
            for r in rows:
                vals = _row_to_vals(r, self._fields)
                if 'id' in r:
                    vals['mysql_id'] = r['id']
                self.create(vals)
            cur.close()
            action = self.env.ref('odoo_reports.action_icc_visits_disposition')
            if action:
                return action.read()[0]
        finally:
            conn.close()


# ============================================================
# 6. Call Disposition
# ============================================================

class IccCallDisposition(models.Model):
    _name = 'icc.call.disposition'
    _description = 'Call Disposition Report'
    _order = 'entry_datesubmitted desc, id desc'

    mysql_id = fields.Integer(string='MySQL ID', readonly=True)
    senumber = fields.Char(string='SE Number')
    name = fields.Char(string='Name')
    lead_list_name = fields.Char(string='Campaign')
    targetsegmentation = fields.Char(string='Segmentation')
    disposition = fields.Char(string='Disposition')
    if_agent_spoke_with_the_contact = fields.Char(string='Agent Spoke with Contact')
    is_merchant_aware_that_they_can_accept_american_express = fields.Char(string='Merchant Aware of Amex')
    will_merchant_agree_to_place_pop = fields.Char(string='Will Place POP')
    calls_agent = fields.Char(string='Agent')
    entry_status = fields.Char(string='Status')
    entry_datesubmitted = fields.Datetime(string='Date Submitted')

    def action_sync_from_mysql(self):
        conn = _get_mysql()
        try:
            cur = conn.cursor(dictionary=True)
            cur.execute("""
                SELECT * FROM call_disposition_report ORDER BY id DESC LIMIT 500
            """)
            rows = cur.fetchall()
            self.search([]).unlink()
            for r in rows:
                vals = _row_to_vals(r, self._fields)
                if 'id' in r:
                    vals['mysql_id'] = r['id']
                self.create(vals)
            cur.close()
            action = self.env.ref('odoo_reports.action_icc_call_disposition')
            if action:
                return action.read()[0]
        finally:
            conn.close()


# ============================================================
# 7. PR Calls Duration
# ============================================================

class IccPrCallsDuration(models.Model):
    _name = 'icc.pr.calls.duration'
    _description = 'PR Calls Duration Report'
    _order = 'date desc, time desc'

    mysql_id = fields.Integer(string='MySQL ID', readonly=True)
    date = fields.Date(string='Date')
    time = fields.Char(string='Time')
    talk_time = fields.Char(string='Talk Time')
    call_id = fields.Char(string='Call ID')
    se_number = fields.Char(string='SE Number')
    disposition = fields.Char(string='Disposition')
    lead_list_name = fields.Char(string='Campaign')
    target_segmentation = fields.Char(string='Segmentation')
    locale = fields.Char(string='Locale')
    mcc_code = fields.Char(string='MCC Code')
    wa_industry = fields.Char(string='Industry')
    agent_name = fields.Char(string='Agent')

    def action_sync_from_mysql(self):
        conn = _get_mysql()
        try:
            cur = conn.cursor(dictionary=True)
            cur.execute("""
                SELECT * FROM pr_calls_duration_report ORDER BY id DESC LIMIT 500
            """)
            rows = cur.fetchall()
            self.search([]).unlink()
            for r in rows:
                vals = _row_to_vals(r, self._fields)
                if 'id' in r:
                    vals['mysql_id'] = r['id']
                self.create(vals)
            cur.close()
            action = self.env.ref('odoo_reports.action_icc_pr_calls_duration')
            if action:
                return action.read()[0]
        finally:
            conn.close()


# ============================================================
# 8. PR Escalation
# ============================================================

class IccPrEscalation(models.Model):
    _name = 'icc.pr.escalation'
    _description = 'PR Calls and Visits Escalation'
    _order = 'date desc, id desc'

    mysql_id = fields.Integer(string='MySQL ID', readonly=True)
    date = fields.Date(string='Date')
    agent_name = fields.Char(string='Agent')
    program = fields.Char(string='Program')
    contact_name = fields.Char(string='Contact Name')
    phone_number = fields.Char(string='Phone')
    business_name = fields.Char(string='Business')
    se_number = fields.Char(string='SE Number')
    complaint_reason = fields.Char(string='Complaint Reason')
    issue_description = fields.Text(string='Issue Description')
    callback_requested = fields.Char(string='Callback Requested')
    call_resolution_notes = fields.Text(string='Resolution Notes')
    resolved = fields.Char(string='Resolved')

    def action_sync_from_mysql(self):
        conn = _get_mysql()
        try:
            cur = conn.cursor(dictionary=True)
            cur.execute("""
                SELECT * FROM pr_calls_visits_escalation ORDER BY id DESC LIMIT 500
            """)
            rows = cur.fetchall()
            self.search([]).unlink()
            for r in rows:
                vals = _row_to_vals(r, self._fields)
                if 'id' in r:
                    vals['mysql_id'] = r['id']
                self.create(vals)
            cur.close()
            action = self.env.ref('odoo_reports.action_icc_pr_escalation')
            if action:
                return action.read()[0]
        finally:
            conn.close()


# ============================================================
# 9. OPTOUTS-DNC
# ============================================================

class IccOptoutsDnc(models.Model):
    _name = 'icc.optouts.dnc'
    _description = 'OPTOUTS-DNC Report'
    _order = 'id desc'

    mysql_id = fields.Integer(string='MySQL ID', readonly=True)
    dnc_number = fields.Char(string='DNC Number')
    seller_id = fields.Char(string='Seller ID')
    created_at = fields.Datetime(string='Created At')

    def action_sync_from_mysql(self):
        conn = _get_mysql()
        try:
            cur = conn.cursor(dictionary=True)
            cur.execute("SELECT * FROM optouts_dnc ORDER BY id DESC LIMIT 500")
            rows = cur.fetchall()
            self.search([]).unlink()
            for r in rows:
                vals = _row_to_vals(r, self._fields)
                if 'id' in r:
                    vals['mysql_id'] = r['id']
                self.create(vals)
            cur.close()
            action = self.env.ref('odoo_reports.action_icc_optouts_dnc')
            if action:
                return action.read()[0]
        finally:
            conn.close()
