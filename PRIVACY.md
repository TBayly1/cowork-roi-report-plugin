# Privacy Notice — Cowork ROI Report Plugin

> **Draft template — not an approved privacy policy.** This notice describes
> behavior documented in the plugin source. It is not a guarantee about every
> Cowork host, connector, tenant, or deployment. The app owner must verify the
> actual data flows, complete every `[OWNER TO COMPLETE]` item, and obtain
> privacy/legal review before distribution. Do not treat this draft as final.

**Effective date:** [OWNER TO COMPLETE]

**Applies to:** Cowork ROI Report - What Cowork Did For Me plugin, version 1.0.0, as deployed by **[OWNER TO COMPLETE: legal name and organization]**.

The short version: **The bundled scripts prepare a report from Cowork activity and artifact metadata available to the signed-in user. The template does not include a publisher-operated report backend or analytics service.** The Cowork runtime and connected services may process information to fulfill the user's request. Email delivery is disabled in the template by default and, if enabled, requires review and approval in the native email workflow. The app owner must confirm these statements for the actual deployment.

## What the plugin is

Cowork ROI Report - What Cowork Did For Me is a Cowork skill package that
instructs an AI assistant to gather available activity and artifact metadata,
classify work, calculate modeled estimates, and generate an HTML report and an
Excel workbook. Its bundled scripts perform local transformations on the data
provided to them. Cowork, Microsoft 365, the AI assistant's provider, and
configured connectors may be involved in retrieving or processing that data.

The report's time-savings, speed-multiplier, and professional-services-value
figures are estimates based on a methodology and assumptions. They are not
audited measurements, billing records, or guaranteed results.

## Data collection

The plugin has no bundled publisher-operated telemetry, crash-reporting, or
usage-analytics endpoint. This does **not** mean that no information is
processed: when a user runs the skill, it may access and process information
available through the configured Cowork runtime and connectors, including:

- Account details needed to identify the signed-in user, such as display name
  and email address.
- Activity metadata for the report period, such as dates, short session goals
  or request summaries, session identifiers, action/tool names, and activity
  counts.
- Artifact metadata such as filenames, file extensions, timestamps, sizes, and
  source links or identifiers when available.
- The configured email recipient address if email delivery is enabled.
- If the optional `mine_session.py` helper is run, the current session
  transcript is read to extract timestamps, turn counts, tool names, and
  output filenames. The helper is designed to write a compact telemetry record,
  not transcript message bodies.

The skill instructs the assistant to use permitted metadata and short
descriptions, and not to add document bodies, email contents, full transcripts,
or credentials merely to fill report fields. The app owner must verify how the
deployed assistant and connectors handle prompts, tool inputs, and source data.

## Where your data goes

The exact data path depends on the Cowork host, connected services, account,
and tenant configuration. Based on the bundled workflow:

1. **Microsoft 365 services** may be accessed through the signed-in user's
   authorized Cowork connectors to find Cowork activity and artifact metadata.
   The plugin does not itself configure or expand the user's Microsoft
   permissions.
2. **The Cowork runtime and AI assistant provider** process the instructions,
   tool results, and report data needed to perform the requested task. Their
   own privacy terms, service settings, and organizational agreements apply.
   **[OWNER TO COMPLETE: identify the actual host, provider, and applicable
   data-processing arrangements.]**
3. **The configured output location** receives the generated HTML report and
   Excel workbook. The documented workflow uses an `output/` folder; in a
   Cowork deployment this may be backed by OneDrive or another host-managed
   workspace.
4. **Outlook/email** is involved only if an administrator enables email and
   configures a recipient. The files are staged for the native email preview;
   the user reviews and approves sending. A recipient may be outside the
   user's organization.

The bundled report-generation scripts do not implement a publisher-operated
backend for collecting report data. This statement does not cover processing
performed by the Cowork host, model provider, Microsoft services, or other
configured connectors.

## What is stored in the workspace

Depending on the host and how the optional helper is used, files may be stored
in the Cowork workspace or its configured OneDrive location:

| Data | Documented location | Purpose |
| --- | --- | --- |
| Generated HTML report and Excel workbook | Cowork run's configured output folder (documented default: `output/`) | Provide the user with the report and workbook |
| Temporary harvest, classification, and report payload files | Cowork run's working folder | Pass data between report-generation steps |
| Optional compact session telemetry | Path supplied to `mine_session.py`; the skill describes an optional durable log in the Cowork sessions area | Preserve session timing, tool counts, and artifact names for later reports |
| Email settings | `settings.json` beside the skill | Keep email disabled by default and, if enabled, identify one configured recipient |

These are paths described by the bundled workflow, not a guarantee of the
storage implementation or location in every host. **[OWNER TO COMPLETE:
confirm deployed workspace paths, whether temporary files are retained, and
whether logs or backups are created.]**

## Third-party sharing

The plugin's own scripts are designed to transform input data into report
files; they do not send those files to a publisher-operated server. Data may
nevertheless be processed or shared by:

- Microsoft 365 services used to retrieve the user's activity or artifacts.
- The Cowork host and AI assistant/model provider used to run the skill.
- Other connectors enabled in the host configuration.
- Outlook and the recipient, if the user approves an email containing the
  report attachments.

The app owner must confirm the complete list of providers and recipients for
the deployed environment: **[OWNER TO COMPLETE]**. Microsoft's handling is
described in the [Microsoft Privacy Statement](https://privacy.microsoft.com/privacystatement)
and may also be governed by the user's organization's agreements and policies.

## Data retention and deletion

This plugin template does not define a central data-retention period or operate
a central store for report data. Retention of reports, temporary files,
optional telemetry, transcripts, and sent email depends on the Cowork host,
storage provider, Microsoft 365 tenant policies, and any enabled services.
**[OWNER TO COMPLETE: specify retention periods, deletion instructions,
backup handling, and how deletion requests are handled for the deployment.]**

Users may delete report files and optional telemetry they control, subject to
their organization's records-management and retention requirements. Data held
by Microsoft or other service providers remains subject to those providers'
terms and the organization's applicable policies.

## Security practices

The template's email setting is `false` by default. When email is enabled, the
exporter validates that the configured recipient is one email address, and
the workflow requires the user to review and approve the native email action.
The settings file is configuration data and should not contain passwords,
tokens, or other secrets.

The repository does not establish the security controls of the Cowork host,
connected services, or deployed environment. **[OWNER TO COMPLETE: describe
the actual access controls, encryption, incident response, and security
contact. Do not claim controls that have not been verified.]**

## Your choices and privacy requests

Email is disabled by default; users can leave it disabled. If enabled, verify
the recipient and attachments before approving a send. Review generated
reports before sharing them because they may include the user's name, activity
dates, goal summaries, filenames, and source references.

For access, correction, deletion, or other privacy requests, contact **[OWNER
TO COMPLETE: privacy contact email or web form and request process]**. The app
owner should explain how requests involving connected services and backups are
handled.

## Children's privacy

This plugin is designed for workplace productivity reporting and is not
intended for children's use. **[OWNER TO COMPLETE: confirm the intended
audience and any applicable age-related requirements.]**

## Changes to this notice

The app owner should update this notice when the plugin or its data practices
change. **[OWNER TO COMPLETE: explain how material changes will be announced
and provide a new effective date.]**

## Contact

Privacy questions or concerns: **[OWNER TO COMPLETE: legal name, organization,
and privacy contact email or web form]**.
