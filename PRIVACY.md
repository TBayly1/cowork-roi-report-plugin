# Privacy Notice — Draft Template

> **Not a final or approved privacy notice.** This draft describes the plugin's
> intended behavior based on its bundled instructions and scripts. The app owner
> must verify it against the deployed configuration and have it reviewed by
> qualified privacy counsel before distributing or using the plugin. Replace
> every `[OWNER TO COMPLETE]` item and remove this warning only after approval.

**Last updated:** [OWNER TO COMPLETE]

## Who is responsible

The app owner responsible for this plugin is **[OWNER TO COMPLETE: legal name
and organization]**. For privacy questions or requests, contact **[OWNER TO
COMPLETE: privacy contact email or web form]**.

## What the plugin does

Cowork ROI Report - What Cowork Did For Me is intended to prepare a personal
impact report from the signed-in user's Cowork activity and related artifact
metadata. It can generate an HTML report and an Excel workbook. The report may
contain modeled time-savings and professional-services-equivalent estimates;
these are estimates, not audited measurements or billing records.

## Information used

Depending on the available sources and runtime configuration, a report run may
process:

- Account details needed to identify the signed-in user, such as display name
  and email address.
- Cowork activity metadata for the reporting period, including dates, short
  session goals or request summaries, session identifiers, action/tool names,
  and available activity counts.
- Artifact metadata such as filenames, file extensions, timestamps, sizes, and
  source links or identifiers when available.
- Email delivery settings, including the configured recipient address if email
  delivery has been enabled.
- Optional live-session telemetry. If the bundled telemetry helper is used, it
  reads the current session's local transcript to extract timing, turn counts,
  tool names, and output filenames. The helper is intended to write a compact
  telemetry record, not to include transcript message bodies in that record.

The workflow instructs the agent to use permitted metadata and short
descriptions, and not to add document bodies, email contents, full transcripts,
or credentials to the report merely to fill fields. The app owner must confirm
that this behavior matches the actual runtime and connector permissions before
approving this notice.

## How information is used and where it goes

The information is used to classify activity, calculate modeled estimates, and
create the requested report files. The bundled scripts are designed to
transform the supplied data into report files; they do not implement a
centralized service or analytics endpoint operated by this template's
publisher. However, the Cowork runtime, connected Microsoft 365 services,
model provider, and any enabled connectors may process information under their
own configurations and terms. **[OWNER TO COMPLETE: identify the actual
runtime, service providers, data flows, and any additional recipients.]**

Report files are written to the output location configured for the Cowork run
(the documented default is an `output/` folder). The report and workbook may
include personal or work-related details such as the user's name, activity
dates, goal summaries, filenames, and source references. Anyone with access to
those files may be able to view that information.

Email delivery is disabled in the distributed template by default. If an
administrator enables it and configures a recipient, the workflow stages the
HTML report and Excel workbook for the Cowork email skill's native Outlook
preview. The user must review and approve sending. The recipient may be outside
the user's organization; sending shares the attachments with that recipient and
is subject to organizational sharing controls.

## Storage and retention

This template does not define a central retention period or operate a central
store for report data. Report, telemetry, transcript, and email retention
depends on the Cowork runtime, the user's storage location, Microsoft 365
settings, and any enabled services. **[OWNER TO COMPLETE: state the actual
storage locations, retention periods, deletion behavior, and backup handling
for the deployed configuration.]**

Users should remove generated reports and optional telemetry they no longer
need, subject to their organization's records and retention policies.

## Choices and requests

Email is off by default. Users can avoid email delivery by leaving it disabled;
they should not enable it until the recipient is configured and the sharing is
appropriate. The report covers the fixed period described by the skill and
processes only sources available to the signed-in user and runtime.

For access, correction, deletion, or other privacy requests, contact **[OWNER
TO COMPLETE: privacy contact and request process]**. The app owner must explain
how requests are handled across report files, connected services, and backups.

## Security and children's privacy

The app owner must document the technical and organizational safeguards used
for the deployed environment: **[OWNER TO COMPLETE]**. Do not place credentials
or secrets in the plugin settings or report inputs.

This plugin is intended for workplace productivity reporting, not for
children's use. **[OWNER TO COMPLETE: confirm intended audience and any
applicable age-related requirements.]**

## Changes to this notice

The app owner should update this notice when the plugin's data practices or
deployment change. **[OWNER TO COMPLETE: explain how users will be notified of
material changes.]**
