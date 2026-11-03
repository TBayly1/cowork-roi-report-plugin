# Privacy Notice — Cowork ROI Report Plugin

**Effective date:** November 3, 2026

**Applies to:** Cowork ROI Report - What Cowork Did For Me plugin, version 1.0.0.

The short version: **The plugin prepares a report from Cowork activity and
artifact metadata available to the signed-in user. The project does not operate
a backend that collects report data, or include project-operated analytics or
telemetry.** The Cowork host, AI assistant provider, Microsoft 365 services,
and any configured connectors process information as needed to carry out the
user's request under their own terms and the user's organizational settings.
Email delivery is off by default and requires user review and approval.

## What the plugin is

Cowork ROI Report - What Cowork Did For Me is a Cowork skill package that
instructs an AI assistant to gather available activity and artifact metadata,
classify work, calculate modeled estimates, and generate an HTML report and an
Excel workbook. The bundled Python scripts transform the data supplied to
them; they do not provide a hosted report-generation service.

The report's time-savings, speed-multiplier, and professional-services-value
figures are estimates based on a methodology and assumptions. They are not
audited measurements, billing records, financial advice, or guaranteed results.

## Data collection

The plugin has no project-operated analytics, usage-tracking, or
crash-reporting endpoint. When the skill is run, information available through
the configured Cowork runtime and connectors may be processed, including:

- The signed-in user's display name and email address, for report identity and
  report metadata.
- Cowork activity metadata for the reporting period, such as dates, short
  session goals or request summaries, session identifiers, action/tool names,
  and available activity counts.
- Artifact metadata such as filenames, file extensions, timestamps, sizes, and
  source links or identifiers when available.
- The configured email recipient address, if email delivery is enabled.

The optional `mine_session.py` helper can read the current session transcript
to extract timestamps, turn counts, tool names, and output filenames. Its
compact telemetry record contains those extracted fields, not transcript
message bodies. The helper can also write that record to a path supplied when
it is run.

The skill directs the assistant to use permitted metadata and short
descriptions, and not to add document bodies, email contents, full transcripts,
or credentials merely to fill report fields. The Cowork host and connected
services may separately process prompts, tool inputs, and tool results under
their own terms and settings.

## Where your data goes

The data path depends on the Cowork host, the user's account and tenant, and
the connectors enabled for the session:

1. **Microsoft 365 services** may be accessed through the signed-in user's
   authorized Cowork connectors to find activity and artifact metadata.
2. **The Cowork host and AI assistant provider** process instructions, tool
   results, and report data needed to perform the requested task. The plugin
   does not control their processing or retention.
3. **The configured workspace** receives generated report files and may hold
   temporary files used by the report workflow. The documented output folder
   is `output/`; the host may back it with OneDrive or another workspace.
4. **Outlook/email** is used only if email is enabled and a recipient is
   configured. The HTML report and Excel workbook are staged for review through
   the host's native email workflow. The user must approve sending. An approved
   message shares the attachments with the configured recipient, who may be
   outside the user's organization.

The plugin's bundled scripts do not send report data to a server operated by
the project maintainers. This does not limit processing by the Cowork host,
model provider, Microsoft services, or other configured connectors.

Microsoft's handling is described in the
[Microsoft Privacy Statement](https://privacy.microsoft.com/privacystatement)
and may also be governed by the user's organization's agreements and policies.
Other providers' terms and policies apply to their services.

## What is stored in the workspace

Depending on the host and how the optional helper is used, files may be stored
in the Cowork workspace or its configured storage location:

| Data | Documented location | Purpose |
| --- | --- | --- |
| Generated HTML report and Excel workbook | Cowork run's configured output folder (documented default: `output/`) | Provide the user with the report and workbook |
| Temporary harvest, classification, and report payload files | Cowork run's working folder | Pass data between report-generation steps |
| Optional compact session telemetry | Path supplied to `mine_session.py`; the skill describes an optional durable log in the Cowork sessions area | Preserve session timing, tool counts, and artifact names for later reports |
| Email settings | `settings.json` beside the skill | Keep email disabled by default and, if enabled, identify one configured recipient |

These are paths described by the bundled workflow. The actual storage location
and whether temporary files are retained depend on the Cowork host and its
configuration.

## Third-party sharing

The project does not receive report data through a project-operated backend.
Information may be processed by the services involved in the user's session:

- Microsoft 365 services used to retrieve the user's activity or artifacts.
- The Cowork host and AI assistant/model provider used to run the skill.
- Other connectors enabled in the host configuration.
- Outlook and the configured recipient, if the user approves an email
  containing report attachments.

The plugin does not independently choose or verify the providers configured
for a Cowork session. Users should consult the applicable provider privacy
terms and their organization's policies.

## Data retention and deletion

The project maintainers do not operate a central store for report data and do
not set a retention period for files created in the user's workspace. The
Cowork host, storage provider, Microsoft 365 tenant policies, and any enabled
services determine retention of reports, temporary files, transcripts,
optional telemetry, and sent email.

Users may delete report files and optional telemetry they control, subject to
their organization's records-management and retention requirements. Deleting
local or workspace files does not delete copies held by service providers,
email recipients, or organizational backups.

## Security practices

The distributed `settings.json` has `email_enabled` set to `false` and uses a
placeholder instead of a real recipient. If email is enabled, the exporter
requires one valid email address, and the workflow stages the attachments for
review and approval in the host's native email action. The exporter itself
does not send email.

The project does not control the security safeguards of the Cowork host,
connected services, or workspace. Users should follow their organization's
access-control and data-handling requirements, protect generated reports, and
never put passwords, access tokens, or other secrets in plugin settings or
report inputs.

## Your choices and privacy requests

Email is disabled by default and can remain disabled. If enabled, check the
recipient and attachments before approving a send. Review generated reports
before sharing them; they may contain the user's name, activity dates, goal
summaries, filenames, and source references.

For activity, storage, access, or deletion requests concerning data held by
Microsoft 365 or another connected provider, use that provider's or your
organization's established privacy-request process. For questions about this
plugin or its source, use the
[project repository](https://github.com/TBayly1/cowork-roi-report-plugin).
Do not post personal, confidential, or sensitive information in a public
issue.

## Children's privacy

This plugin is designed for workplace productivity reporting and is not
directed to children.

## Changes to this notice

Changes to this notice will be published in the project repository. The
effective date above will be updated when the notice changes.

## Contact

For questions about the plugin source, contact the project maintainers through
the [GitHub repository](https://github.com/TBayly1/cowork-roi-report-plugin).
For privacy requests about information held by a service provider or your
organization, use that provider's or organization's privacy-request process.
