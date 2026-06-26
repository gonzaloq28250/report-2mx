import logging
from pathlib import Path

from odoo import http
from odoo.http import request
from werkzeug.wrappers import Response as WerkzeugResponse

_logger = logging.getLogger(__name__)

_REPORTS_DIR = Path(__file__).resolve().parent.parent.parent / 'reports'


class IccReportController(http.Controller):

    @http.route('/icc/report/download/<int:report_id>', type='http', auth='user', methods=['GET'])
    def download_report(self, report_id, **kwargs):
        """Download a generated report file directly from disk."""
        report = request.env['icc.report'].browse(report_id)
        if not report.exists() or report.state != 'generated':
            return request.not_found()

        filename = report.file_name or 'report.xlsx'
        file_path = _REPORTS_DIR / filename

        if not file_path.exists():
            return request.not_found()

        def generate():
            with open(file_path, 'rb') as f:
                while True:
                    chunk = f.read(8192)
                    if not chunk:
                        break
                    yield chunk

        content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        if filename.endswith('.txt'):
            content_type = 'text/plain'

        return WerkzeugResponse(
            generate(),
            content_type=content_type,
            headers={
                'Content-Disposition': 'attachment; filename="%s"' % filename,
                'Cache-Control': 'no-cache',
            },
        )
