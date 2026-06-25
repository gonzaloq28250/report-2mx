{
    'name': 'ICC Amex Reports',
    'version': '18.0.1.0.0',
    'summary': 'Generate and download ICC Amex Excel reports',
    'description': """
ICC Amex Reports
================
Triggers Python generation scripts for ICC Amex reports.
Generated Excel files are stored in Odoo and can be downloaded.

Reports:
- Daily Campaign Results v3
- Calls Consolidate Report
- Call Disposition Report
- Visit Log Optblue Report
- Visits Disposition Report
- PR Calls Duration Report
- PR Escalation Report
- Service Level Report
- OPTOUTS-DNC Report
    """,
    'category': 'Reporting',
    'author': 'ICC Amex',
    'website': '',
    'license': 'LGPL-3',
    'depends': ['base', 'web'],
    'data': [
        'security/ir.model.access.csv',
        'views/icc_report_views.xml',
        'views/icc_report_menu.xml',
    ],
    'assets': {
        'web.assets_backend': [
            'odoo_reports/static/src/js/report_download.js',
        ],
    },
    'installable': True,
    'application': True,
    'auto_install': False,
}
