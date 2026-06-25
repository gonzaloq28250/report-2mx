import io
import logging

from odoo import http
from odoo.http import request, Response

_logger = logging.getLogger(__name__)


class IccReportController(http.Controller):

    @http.route('/icc/report/download/<int:report_id>', type='http', auth='user', methods=['GET'])
    def download_report(self, report_id, **kwargs):
        """Download a generated report file."""
        report = request.env['icc.report'].browse(report_id)
        if not report.exists() or report.state != 'generated' or not report.file_data:
            return request.not_found()

        file_data = report.file_data
        filename = report.file_name or 'report.xlsx'

        content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        if filename.endswith('.txt'):
            content_type = 'text/plain'

        return Response(
            body=file_data,
            headers=[
                ('Content-Type', content_type),
                ('Content-Disposition', 'attachment; filename="%s"' % filename),
                ('Content-Length', str(len(file_data))),
                ('Cache-Control', 'no-cache'),
            ],
        )
