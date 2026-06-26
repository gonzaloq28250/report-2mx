import logging
from pathlib import Path

from odoo import http
from odoo.http import request
from werkzeug.utils import send_file

_logger = logging.getLogger(__name__)

_REPORTS_DIR = Path(__file__).resolve().parent.parent.parent / 'reports'


class IccReportController(http.Controller):

    @http.route('/icc/report/download/<int:report_id>', type='http', auth='user', methods=['GET'])
    def download_report(self, report_id, **kwargs):
        report = request.env['icc.report'].browse(report_id)
        if not report.exists() or report.state != 'generated':
            return request.not_found()

        filename = report.file_name or 'report.xlsx'
        file_path = _REPORTS_DIR / filename

        if not file_path.exists():
            return request.not_found()

        return send_file(
            str(file_path),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename,
        )
