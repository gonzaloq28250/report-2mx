import logging
from pathlib import Path

from odoo import http
from odoo.http import request

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

        _logger.info('Download: %s (exists=%s)', file_path, file_path.exists())

        if not file_path.exists():
            return request.not_found()

        with open(str(file_path), 'rb') as f:
            file_data = f.read()

        _logger.info('Serving %d bytes for %s', len(file_data), filename)

        headers = [
            ('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
            ('Content-Disposition', 'attachment; filename="%s"' % filename),
            ('Content-Length', str(len(file_data))),
        ]
        return request.make_response(file_data, headers)
