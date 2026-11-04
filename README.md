# Cowork ROI Report

This repository contains the Cowork plugin package and a Power BI report
template. Use the two processes below to install the skill and load its
generated reporting data into Power BI.

## Process 1: Install the plugin

1. Download the [Cowork ROI Report Plugin Template v1.0.0 ZIP](Cowork%20ROI%20Report%20Plugin%20Template%20v1.0.0.zip).
2. Before uploading it, open the ZIP and edit
   `skills/cowork-roi-report-what-cowork-did-for-me/settings.json`.
3. Configure the administrator email settings:

   ```json
   {
     "email_enabled": false,
     "email_to": "Add Admin Email here"
   }
   ```

   To enable email, replace `"Add Admin Email here"` with one valid
   administrator email address and set `"email_enabled"` to `true`. Leave
   email disabled if you do not want the skill to prepare an email. The skill
   stages any email for review in Cowork's native approval workflow; it does
   not send email by itself.
4. Upload the plugin ZIP to your organization's skills registry, following
   that registry's import and approval process.

## Process 2: Load the Power BI report

1. Download the [Power BI report template](What%20Cowork%20Did%20For%20Us%20-%20PowerBi%20Report%20Template.pbit).
2. Open the `.pbit` file in Power BI Desktop.
3. When prompted, select the SharePoint site and folder that contain the
   reporting data, then authenticate with an account that has access.
4. Refresh the data and load it into the report. If Power BI prompts you to
   configure or confirm data-source settings, follow the prompts for the
   intended SharePoint site and folder.

The plugin's Excel export provides four reporting tables: `Sessions`,
`Artifacts`, `Allocations`, and `References`. Use the intended report data
location and ensure it contains the workbook produced by the plugin.

## Important

The report's time-savings, speed-multiplier, and value figures are modeled
estimates, not audited measurements, financial advice, or guarantees. Review
the assumptions and underlying data before relying on or sharing report
outputs. Your organization is responsible for configuring and approving the
plugin and report for its environment, including data access, email use,
privacy, security, and retention requirements.

See [PRIVACY.md](PRIVACY.md), [TERMS.md](TERMS.md), and
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) for additional information.
