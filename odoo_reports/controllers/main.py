import base64
import logging

from odoo import http
from odoo.http import request

_logger = logging.getLogger(__name__)


class IccReportController(http.Controller):

    @http.route('/icc/report/download/<int:report_id>', type='http', auth='user', methods=['GET'])
    def download_report(self, report_id, **kwargs):
        _logger.info('=== Download request for report %d ===', report_id)

        report = request.env['icc.report'].browse(report_id)
        if not report.exists() or report.state != 'generated':
            _logger.warning('Report %d not found or not generated', report_id)
            return request.not_found()

        filename = report.file_name or 'report.xlsx'
        _logger.info('File: %s, file_data type: %s, len: %s',
                      filename, type(report.file_data),
                      len(report.file_data) if report.file_data else 0)

        if not report.file_data:
            _logger.warning('No file_data for report %d', report_id)
            return request.not_found()

        file_data = report.file_data
        if isinstance(file_data, str):
            _logger.info('Decoding base64 string')
            file_data = base64.b64decode(file_data)

        _logger.info('Serving %d bytes', len(file_data))

        headers = [
            ('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
            ('Content-Disposition', 'attachment; filename="%s"' % filename),
            ('Content-Length', str(len(file_data))),
        ]
        return request.make_response(bytes(file_data), headers)
