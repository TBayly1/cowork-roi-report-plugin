# Changelog

Notable changes to Cowork ROI Report - What Cowork Did For Me are recorded here.

## 1.0.0

Initial plugin release, adapted from **What Cowork Did For Me v48**.

- Added an Excel export for Power BI reporting. The bundled
  `to_xlsx.py` script converts the report payload into four structured
  worksheets: Sessions, Artifacts, Allocations, and References.
- Added optional email delivery of the generated HTML report and Excel
  workbook to a configured administrator. Email is disabled by default; when
  enabled, the recipient and attachments are presented for review in Cowork's
  native email approval workflow. The exporter stages the email plan and does
  not send email itself.
