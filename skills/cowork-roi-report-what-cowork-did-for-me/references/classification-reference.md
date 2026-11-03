# Classification reference — heuristics, counting discipline & raw-harvest schema

Reference — extension/category heuristics the classifier encodes:
| Signal in the session | Category key |
|---|---|
| Multi-source **synthesis** report — newsletter, weekly/biweekly recap, "Wrapped", briefing, ROI summary, status review (pull data → synthesize → write-up) | `analysis` |
| Data **analysis / validation / metrics / KPI mapping / catalog**, analytical spreadsheet | `analysis` |
| **Deck** (.pptx), **document** (.docx), guide, one-pager, written content, PDF report | `document` |
| **Interactive web app / dashboard / builder / hub** (.html app), or a **script** (.ps1/.py/.js) | `code` |
| **Prompt engineering / skill authoring / packaging** (prompt .md, skill .md, skill .zip), cross-system automation | `special` |
| **Email** sent via Cowork | `email` · **Teams** message | `comms` · **Meeting** scheduled/recapped | `meeting` |
| Quick **Q&A / formatting / lookup / short review** with no saved deliverable | `general` |

Counting discipline:
- Cap **~2 run tasks per session**. Fold supporting files (prompts, how-tos, design specs, READMEs, lock files, zips) into the primary task.
- Genuinely distinct deliverables (e.g., two different customer analyses in one session) → separate tasks.
- Categories with **no** artifacts in the window are reported as **zero** — this keeps totals a conservative floor.

The raw harvest you write to `working/cowork_raw.json` (input to `classify.py`):
```json
{ "meta": {"user":"<name>","email":"<mail>","generated":"<YYYY-MM-DD>",
           "window":{"from":"<local date 59 days before today>","to":"<today's local date>","label":"Last 60 days","months":2},
           "hourly_rate":72,
           "coverage":{"pagination_complete":<true|false|null>,"layouts_checked":[...],"pages_read":<int>}},
  "sessions": [ {"id":"<actual session key>","session_id":"<actual full source session ID when available>",
                 "folder_url":"<actual source folder URL when available>","date":"YYYY-MM-DD","hour":<0-23>,
                 "goal":"<short verb-first phrase>",
                 "inputs":  [{"name":"report-1.pdf","ext":"pdf"}, ...],
                 "outputs": [{"name":"deck.pptx","ext":"pptx","skills":["Presentation Design","Data Analysis"]}, ...],
                 "skills": ["Data Analysis"],
                 "professional_roles": ["Data Analyst","Management Consultant"],
                 "request":"<original ask>", "actions":["mcp__outlook__ListMessages","Write", ...],
                 "apps_accessed":["Outlook","Excel"], "sources_reviewed":<int>,
                 "has_folder":true, "exec_min":<measured minutes|null>}, ... ] }
```
`classify.py` adds the `tasks` array (categories) and writes `working/cowork_sessions.json`. Where a live
`session_telemetry.json` exists for a session, prefer its measured `exec_min`, tool counts and `produced_artifact`
flag over the file-timestamp estimate.

Preserve all available input/output metadata (drive and item IDs, `webUrl`, source path, timestamps,
creator application and size) in those objects. Preserve supplied session evidence and provenance
through classification; do not replace the real source keys with invented aliases or links.

## Four-table Power BI export contract (payload schema 2.0)

Output names use `cowork-roi-report - <User Name> - <YYYY-MM-DD>.html/.xlsx` from the signed-in user's
display name in `meta.user` and the run's local report date in `meta.generated`. The shared
`scripts/output_paths.py` helper sanitizes unsafe filename characters consistently. Optional CSV uses
the same suffix after `cowork-sessions`. The email plan points to the exact matching output paths.
Do not use the delivery recipient or a report-window start date as the filename's identity/date.
For future BI refresh, select the intended dated snapshot rather than blindly appending overlapping exports.

The same computed payload drives HTML and Excel. Keep all numeric facts at full precision and round
only their display. Export exactly four worksheets with structured tables: Sessions (`tblSessions`),
Artifacts (`tblArtifacts`), Allocations (`tblAllocations`) and References (`tblReferences`).
References is intentionally limited to the rows consumed by the current Power BI report; detailed facts
remain in the first three tables, while payload-audit and documentation rows are not exported.

The core allocation facts are:
- Session/category: `tasks`, one row for each session/category, with `runs`, `category_label`, `intent`,
  `minutes_low/typical/high`, `hours_low/typical/high`, and `value_low/typical/high`.
- Session/role: `session_roles`, one row per session/role with allocation `weight` and all three scenarios.
- Session/output/skill or chat-only session/skill: `skill_allocations`, with `deliverable_id`, `scope`,
  `skill`, `weight`, and all three scenarios.
- Input/output occurrences: stable occurrence IDs rather than filename keys. Preserve repeated
  filenames; `artifact_key` describes the separate headline deduplication rule.

The three allocation arrays stack into Allocations with `Allocation type = category|role|skill`.
Filter to one allocation type before aggregating. Each type independently covers the same session
effort; do not add them together. Inputs/outputs stack into Artifacts with `Direction = input|output`,
and only output rows carry attributed effort.

Keep the complete `cowork_fit` object and supplied action/app/source/telemetry metadata on session
facts. Inferred app surfaces are not verified tool activity. Missing telemetry stays unknown, not
zero; keep modeled assisted minutes separate from measured runtime.

`meta.run_id` identifies the snapshot. Sessions, Artifacts and Allocations contain collision-safe
`Session key` values; Artifacts and output-skill allocations also contain `Artifact occurrence key`.
Join Sessions to Artifacts and Allocations by `Session key`; use `Artifact occurrence key` only for
optional output-skill drilldown. These relationship instructions are documented here, not in References.
Each workbook is one snapshot of the selected window. Do not
append overlapping snapshots or sum session facts with their category/role/skill allocations.

References retains all `summary` and `activity` records, plus only the `configuration` fields
`email_to` and `Payload schema`. The existing Power BI query uses these rows for KPI values,
activity-by-hour data, report identity and schema identification. Keep the 13-column `tblReferences`
schema and table name stable; do not export other record types or configuration fields. Typed value
columns preserve text, number, Boolean and date values; `Number literal` remains available for exact
numeric literals. Do not sum arbitrary `Number value` rows.

Coverage and source identity must be reported honestly. Missing session URLs, blank JTBDs and default
process/pillar assignments remain visible as gaps. Export does not validate those business labels or
magically reconstruct metadata that an older report omitted.

**Workflow evidence (`request` / `actions` / `apps_accessed` / `sources_reviewed`).** These are the
action-grounded record of *what Cowork did* — mined by `mine_session.py` from the transcript. `compute.py`'s
Cowork-fit unions the **verified** `apps_accessed` with the apps **inferred** from outputs, so a cross-app
workflow (e.g. Outlook + Excel) grades **H** on evidence, not on the single file that landed in OneDrive.
**Populate them only when the action history is actually available; leave them ABSENT (not `[]`) when it is
not** — their absence is the signal that the assessment must fall back to **"Insufficient evidence"** rather
than assume a single-app workflow from a lone output.
