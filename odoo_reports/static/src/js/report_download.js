/** @odoo-module **/
import { registry } from "@web/core/registry";
import { Component } from "@odoo/owl";

// Download button component for the form view
// The HTTP route /icc/report/download/<id> handles the actual download
// This is a placeholder for any client-side interactions if needed

registry.category("actions").add("icc_report_download", class IccReportDownload extends Component {
    setup() {
        // Download is handled via HTTP link in the form view
    }
});
