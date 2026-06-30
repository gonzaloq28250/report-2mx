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

        if not file_path.exists():
            return request.not_found()

        with open(str(file_path), 'rb') as f:
            file_data = f.read()

        headers = [
            ('Content-Type', 'application/octet-stream'),
            ('Content-Disposition', 'attachment; filename="%s"' % filename),
            ('Content-Length', len(file_data)),
            ('Content-Encoding', 'identity'),
            ('Cache-Control', 'no-cache, no-store, must-revalidate'),
            ('Pragma', 'no-cache'),
            ('Expires', '0'),
        ]
        return request.make_response(file_data, headers)

    @http.route('/icc/report/test/<filename>', type='http', auth='user', methods=['GET'])
    def test_download(self, filename, **kwargs):
        """Test endpoint: download any file from reports/ by name"""
        file_path = _REPORTS_DIR / filename
        if not file_path.exists():
            return request.not_found()

        with open(str(file_path), 'rb') as f:
            file_data = f.read()

        headers = [
            ('Content-Type', 'application/octet-stream'),
            ('Content-Disposition', 'attachment; filename="%s"' % filename),
            ('Content-Length', len(file_data)),
            ('Content-Encoding', 'identity'),
        ]
        return request.make_response(file_data, headers)
