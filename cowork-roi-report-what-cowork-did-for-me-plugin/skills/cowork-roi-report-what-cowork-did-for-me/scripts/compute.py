#!/usr/bin/env python3
"""Cowork ROI — compute the methodology payload from a classified-sessions JSON.

v5 (artifact-scaled speed multiplier). The skill harvests the signed-in user's
own Cowork sessions, classifies each into run tasks AND records the distinct
input artifacts analyzed and output artifacts produced. This script applies a
two-clock model:

  Expert clock (unassisted)  = research-anchored analysis/general bands (v4)
                               + per-input reading/triage time
                               + per-output authoring time
  Assisted clock (your time) = a small fixed prompt/setup cost
                               + a per-artifact handling cost (modeled, not measured)

  Speed multiplier            = Expert clock / Assisted clock
  Professional-services value = Expert-clock hours x hourly rate

There is no ROI / seat-cost figure (credit & seat consumption is unavailable),
so value is framed as the professional-services equivalent at the selected rate.

INPUT (working/cowork_sessions.json):
{
  "meta": {"user","email","generated",
           "window":{"from","to","label","months"}, "hourly_rate":72},
  "sessions": [
     {"id","date","hour","goal",
      "inputs":  [{"name":"report-1.pdf","ext":"pdf"}, ...],   # analyzed
      "outputs": [{"name":"deck.pptx","ext":"pptx"}, ...],  # produced
      "tasks":   ["analysis","document"]},                  # category keys
     ...
  ]
}
Usage: python compute.py --in working/cowork_sessions.json --out working/cowork_roi_data.json
"""
import json, argparse, collections, copy, hashlib, math

# v4 research-anchored bands (min saved per run task): (low, typical, high, label)
# Category bands are template estimates; verify their research sources before relying on them.
#   Analysis & Research typical: 71 → 67  (Stanford-WB basket mean 335÷5=67)
#   Meeting workflows high:      45 → 43  (slide 7)
#   Communication workflows high: 6 → 11  (slide 8)
CATS = {
 "analysis":(30,67,92,"Analysis & Research"),
 "document":(12,24,42,"Document & content creation"),
 "email":(3,7,12,"Email workflows"),
 "meeting":(12,31,43,"Meeting workflows"),
 "comms":(2,4,11,"Communication workflows"),
 "special":(10,25,40,"Specialized workflows"),
 "code":(30,56,96,"Write or debug code"),
 "general":(2,5,8,"General assistance / Other"),
}
INTENT={"analysis":"Researching","document":"Building","email":"Communicating",
        "meeting":"Coordinating","comms":"Communicating","special":"Building",
        "code":"Building","general":"Researching"}
# Roles are no longer a fixed category map — they come from each session's professional_roles
# (LLM-tagged "roles a billing firm would charge"; keyword fallback in classify.py).

# ---- model constants ----
# Expert clock = sum of research-anchored CATS bands per task (no read/author heuristics).
# Assisted clock is a small MODELED hands-on estimate (OneDrive can't measure keystroke time).
ASSIST_FIXED   = 8    # min fixed prompt/setup per session (modeled)
ASSIST_PER_ART = 2    # min hands-on handling per artifact (in or out) — modeled
IMG_EXT = {"png","jpg","jpeg","gif","bmp","webp","heic"}
DATA_CAT_EXT = {"xlsx","xls","csv"}                                  # -> Analysis & Research
CODE_CAT_EXT = {"py","js","ps1","ipynb","sh","ts","go","sql"}        # -> Write or debug code
APP_CAT_EXT  = {"html","htm"}                                        # -> Specialized workflows

# friendly artifact-type labels for the Analyzed -> Produced section
TYPE_LABEL = {
 "pdf":"PDF","doc":"Document","docx":"Document","ppt":"Deck","pptx":"Deck",
 "xls":"Spreadsheet","xlsx":"Spreadsheet","csv":"Spreadsheet",
 "html":"Web page","htm":"Web page",
 "png":"Image","jpg":"Image","jpeg":"Image","gif":"Image","bmp":"Image","webp":"Image","heic":"Image",
 "md":"Text","txt":"Text",
 "py":"Script","js":"Script","ps1":"Script","sh":"Script","ts":"Script","go":"Script","sql":"Query","ipynb":"Notebook",
}
def type_label(ext): return TYPE_LABEL.get(ext.lower(),"File")

# Per-artifact AUTHORING anchor (min a professional would spend producing ONE deliverable),
# by friendly type — (low, typical, high). Scales the expert clock with the NUMBER of artifacts.
# Images use a deliberately small anchor so auto-generated QA screenshots/variants don't inflate.
AUTHOR_BAND = {
 "Deck":(30,45,60),"Document":(25,40,55),"Spreadsheet":(25,35,50),
 "Web page":(25,35,50),"Script":(25,35,50),"Notebook":(25,35,50),
 "Query":(15,25,35),"Image":(2,5,10),"Text":(10,20,30),"File":(15,30,45),
}
# Per-input READ/triage anchor (min to analyze ONE source) — scales with analyses performed.
READ_DOC=(8,12,18); READ_IMG=(2,5,8)
# Per LOGIC-LINE-OF-CODE anchor (min an unassisted professional spends per finished, reviewed SLOC).
# (low, typical, high) = 40 / 30 / 20 LOC per hour — industry range for production-quality code.
LOC_MIN=(1.5,2.0,3.0)
# Only SUBSTANTIVE produced artifacts earn an authoring anchor — Word docs, decks, spreadsheets, written
# specs. Code is valued separately by its logic LOC; generated web/HTML output, images, and QA/variant
# files are TRIVIAL and earn nothing (per user direction: ignore trivial artifacts).
SUBSTANTIVE_AUTHOR={"Document","Deck","Spreadsheet","Text"}
# Task bands that still apply on top (the analysis/coordination "thinking"); code & special are valued
# via LOC / authoring, so their flat bands are dropped to avoid double-counting.
KEEP_BANDS={"analysis","document","email","meeting","comms","general"}
READ_SOURCE_EXT=DATA_CAT_EXT|{"pdf","docx","doc","pptx","ppt"}
import re
REVISION_WEIGHT=0.3   # later versions of the SAME artifact = revisions, not fresh authoring
_VER_TAIL=re.compile(r'[-_ ]?v?\d+(\.\d+)?$')
_VER_WORD=re.compile(r'[-_ ](v\d+|\d+d|\d+day|sample|draft|final|copy|mockup|backup|old|new)$')
def art_base(name):
    """Normalize an artifact name to its version-independent base so v6/v11/v17 (or report-v1/v2/final)
    collapse to ONE artifact — used so iterative versions aren't counted as distinct deliverables."""
    n=str(name).lower().rsplit('.',1)[0]
    for _ in range(4):
        n2=_VER_WORD.sub('',_VER_TAIL.sub('',n))
        if n2==n: break
        n=n2
    return n.strip('-_ ') or str(name).lower()

def hrs(m): return m/60
def round_to_total(parts_min, total_hours):
    """Round each part (minutes) to 1-decimal hours so the parts sum EXACTLY to total_hours."""
    import math
    tt=round(total_hours*10)
    raw=[(m/60.0)*10 for m in parts_min]
    fl=[math.floor(x) for x in raw]
    rem=int(tt-sum(fl))
    order=sorted(range(len(raw)), key=lambda i: raw[i]-fl[i], reverse=True)
    for i in range(max(rem,0)): fl[order[i%len(order)]]+=1
    return [round(t/10.0,1) for t in fl]
def out_cat(ext):
    e=ext.lower()
    if e in DATA_CAT_EXT: return "Analysis & Research"
    if e in CODE_CAT_EXT: return "Write or debug code"
    if e in APP_CAT_EXT:  return "Specialized workflows"
    return "Document & content creation"

def _ext(a):
    if isinstance(a,dict): return (a.get("ext") or a.get("name","").split(".")[-1] or "file").lower()
    s=str(a); return (s.split(".")[-1] if "." in s else "file").lower()
def _name(a):
    return a.get("name","artifact") if isinstance(a,dict) else str(a)

def session_expert(runs, idx):
    """Deck-anchored expert clock (Cowork Time-Savings Methodology): the SUM over RUNS of each run's
    category band — minutes SAVED per run. Each band already sums the activity-instance chain inside ONE
    run (e.g. code 56 = write+test+debug = 18min x 3 steps; doc 24 = 6.1min x 4 instances), so we count
    RUNS and multiply by the band — never per-LOC, never a per-artifact authoring add-on (that would
    double-count the chain). `runs` is {category: run_count}. idx: 0=low, 1=typical, 2=high."""
    return sum(n * CATS.get(c, CATS["general"])[idx] for c, n in (runs or {}).items())

import os
COST_LOG = "/mnt/user-config/.claude/cowork-session-costs.json"
def load_cost_lookup():
    """Map session id (full UUID and 8-char prefix) -> actual $ cost captured by the
    statusLine hook. Empty if the log doesn't exist yet (older sessions have no cost)."""
    try:
        with open(COST_LOG) as f:
            d=json.load(f)
        out={}
        for k,v in (d.items() if isinstance(d,dict) else []):
            amt = v.get("total_cost_usd") if isinstance(v,dict) else None
            if isinstance(amt,(int,float)):
                out[k]=amt; out[k[:8]]=amt
        return out
    except Exception:
        return {}

CREDITS_LOG = "/mnt/user-config/.claude/cowork-session-credits.json"
def load_credits_lookup():
    """Map session id (full + 8-char) -> real AiCredits from /cost readings (durable ledger)."""
    try:
        with open(CREDITS_LOG) as f:
            d=json.load(f)
        out={}
        for k,v in (d.get("sessions",{}) or {}).items():
            amt = v.get("credits") if isinstance(v,dict) else v
            if isinstance(amt,(int,float)):
                out[k]=amt; out[k[:8]]=amt
        return out
    except Exception:
        return {}


# Live-session telemetry: total interactions per session (user + assistant turns).
# Present only when the telemetry hook recorded the session; keyed by id + 8-char id.
TELEMETRY_LOG = os.path.join(os.path.dirname(CREDITS_LOG), "cowork-session-telemetry.json") \
    if 'CREDITS_LOG' in dir() else "/mnt/user-config/.claude/cowork-session-telemetry.json"
def load_interactions_lookup():
    try:
        with open(TELEMETRY_LOG) as f:
            d=json.load(f)
        out={}
        for k,v in (d or {}).items():
            if not isinstance(v,dict): continue
            t=v.get("turns",{}) or {}
            tot=(t.get("user",0) or 0)+(t.get("assistant",0) or 0)
            if t.get("user") is not None or t.get("assistant") is not None:
                for key in (v.get("id"), v.get("session_id"), k):
                    if key: out[str(key)]=tot; out[str(key)[:8]]=tot
        return out
    except Exception:
        return {}


# ---------------------------------------------------------------------------
# Cowork-fit - the "single-surface test"
# Guiding question: could a single surface-specific / in-app Copilot have done
# this end-to-end? If yes, it did not need Cowork.
#   * build/package a skill, code, HTML app, or an EXECUTED automation/connector
#       /browser run  -> no in-app surface exists  -> H (Cowork-only)
#   * >= 2 in-app surfaces (cross-surface chain, e.g. Outlook + Excel)
#                                                   -> H (no single Copilot spans apps)
#   * exactly 1 in-app surface                      -> L (name the surface)
# The rule baseline is deterministic. An optional review can override it, subject
# to the evidence guard at the end of cowork_fit().
# ---------------------------------------------------------------------------
CF_IN_APP = {"xlsx":"Excel","xlsm":"Excel","xls":"Excel","csv":"Excel","tsv":"Excel",
             "docx":"Word","doc":"Word","pdf":"Word","rtf":"Word",
             "pptx":"PowerPoint","ppt":"PowerPoint"}
CF_CODE_EXT  = {"py","ps1","js","ts","ipynb","sh","html","htm"}
CF_BUILD_EXT = {"zip"}
CF_CONNECTORS = ("dynamics 365","ado boards","power bi","fabric","servicenow","salesforce"," sap ")
CF_SURFACE_ORDER = ["Outlook","Excel","Word","PowerPoint","Teams"]
# Roles that are NOT linked to a business-value-analytics job -> supporting signal
CF_OUT_OF_DOMAIN_ROLES = {"Software Engineer","DevOps Engineer","Frontend Developer",
                          "Backend Developer","Solutions Architect","Data Engineer",
                          "QA Engineer","Security Engineer","UX Designer"}

def cf_surfaces(goal, outputs, cats):
    """The in-app Copilot surfaces a piece of work would touch (Outlook/Excel/Word/PPT/Teams)."""
    g = " " + (goal or "").lower() + " "
    surfaces = set()
    for o in outputs:
        e=_ext(o).lower()
        if e in CF_IN_APP: surfaces.add(CF_IN_APP[e])
    if ("Email workflows" in cats) or ("email" in cats) or any(w in g for w in ("email","inbox","reply","triage","e-mail","shar","distribut","notify","newsletter","broadcast","circulate","outreach")):
        surfaces.add("Outlook")
    if ("Meeting workflows" in cats) or ("Communication workflows" in cats) or ("meeting" in cats) or ("comms" in cats) or any(w in g for w in ("meeting","transcript","recap","teams","standup")):
        surfaces.add("Teams")
    if any(w in g for w in ("deck","slide","presentation","chart","visual","layout")):
        surfaces.add("PowerPoint")
    return surfaces

# Grade-appropriate fallback labels used when an LLM review CHANGES the grade but
# supplies no label of its own. Without this, an upgraded High keeps the rule's
# original single-app label ("<app> Copilot could do it") and the hover then reads
# "<app> Copilot could do it: <High justification>" — which undercuts the High.
# A High label must justify the NEED for Cowork; Medium/Low state the lighter fit.
CF_REVIEW_LABELS = {
    "H": "Needs Cowork",
    "M": "Moderate fit",
    "L": "Single-app Copilot could do it",
    "?": "Insufficient evidence",
}

def cowork_fit(goal, outputs, inputs, cats, roles, extra_surfaces=None, review=None, evidence=None):
    g = " " + (goal or "").lower() + " "
    out_ext = [_ext(o).lower() for o in outputs]
    conversational = (len(outputs) == 0)

    # -- Cowork-only: builds/packages a skill, code, or HTML/app --
    build = (any(e in CF_BUILD_EXT for e in out_ext)
             or any(e in CF_CODE_EXT for e in out_ext)
             or (("skill" in g or "package" in g or "bundle" in g) and ("build" in g or "create" in g))
             or (any(w in g for w in ("html","dashboard","web app","webpage","web page",
                                      "website","microsite","landing page","interactive report"))
                 and any(w in g for w in ("build","create","generat","make","develop","design"))))
    # -- Cowork-only: an EXECUTED automation / connector / browser run --
    executed = ((("browser automation" in g) or ("connector" in g) or ("integrat" in g)
                 or any(c in g for c in CF_CONNECTORS) or ("sweep" in g))
                and (("run" in g) or ("sweep" in g) or ("execute" in g) or len(outputs) > 0))
    # -- >5 sources synthesized is also inherently agentic --
    big_synth = len(inputs) > 5
    # -- multi-file / multi-format signals. Heavy synthesis / several distinct OUTPUT
    #    formats are Cowork-worthy (H); lighter multi-format INPUT or sheer file volume is
    #    a MODERATE (M) fit. --
    in_ext = [_ext(i).lower() for i in (inputs or [])]
    in_fmts = set(e for e in in_ext if e)
    out_fmts = set(e for e in out_ext if e)
    n_in, n_out = len(inputs or []), len(outputs)
    multi_doc_synth = (n_in >= 3 and len(in_fmts) >= 2) or big_synth   # synthesize many docs across formats -> H
    multi_out_fmt = len(out_fmts) >= 2                                 # produces >=2 distinct output formats -> H
    multi_fmt_input = len(in_fmts) >= 2                               # reads >=2 formats (below synth) -> M
    many_files = (n_in >= 3) or (n_out >= 3)                          # sheer volume of files -> M
    moderate_complexity = multi_fmt_input or many_files

    # -- automation-like process (inbox triage, channel scan, sweep, workflow, recurring
    #    run). These orchestrate work across items, so they are never a single in-app
    #    Copilot task and can never be graded Low. --
    AUTOMATION_SIGNALS = ("triage","scan","sweep","automat","workflow","monitor",
        "recurring","batch","bulk","orchestrat","pipeline","auto-")
    automation_kind = next((w for w in AUTOMATION_SIGNALS if w in g), None)

    # -- WORKFLOW EVIDENCE. Assess what Cowork DID, not just the file it produced.
    #    `evidence` carries the action-grounded trace mined from the session transcript:
    #      apps    = M365 apps the ACTIONS actually touched (verified)
    #    Verified apps are unioned with the surfaces INFERRED from outputs/goal and with
    #    surfaces from sessions that share an output basename (extra_surfaces).
    #    `actions` and `sources_reviewed` are preserved upstream for transparency but do
    #    not currently change this rule baseline. --
    ev = evidence or {}
    supplied_apps = ev.get("apps")
    if supplied_apps is None:
        supplied_apps = []
    if not isinstance(supplied_apps, list):
        raise ValueError("Observed apps must be a list of names or app metadata objects")
    verified_apps = set()
    for app in supplied_apps:
        if isinstance(app, dict):
            app = next((app[key] for key in ("app", "name", "app_name", "displayName", "id", "app_id")
                        if isinstance(app.get(key), str) and app[key]), None)
        if not isinstance(app, str) or not app.strip():
            raise ValueError("Each observed app requires a nonempty name or ID")
        verified_apps.add(app)
    inferred_surfaces = cf_surfaces(goal, outputs, cats) | set(extra_surfaces or [])
    surfaces = inferred_surfaces | verified_apps
    def _order(ss):
        return [x for x in CF_SURFACE_ORDER if x in ss] + sorted(s for s in ss if s not in CF_SURFACE_ORDER)
    # verified cross-app = >=2 apps with at least one proven by the action trace
    verified_cross_app = (len(verified_apps) >= 2) or (len(verified_apps) >= 1 and len(surfaces) >= 2)

    # -- Specialized workflow = Cowork-native (automation, connectors, scheduled/recurring
    #    prompts, skill build/packaging). Per methodology, this is a High Cowork match. --
    is_special = ("Specialized workflows" in cats) or ("special" in cats)

    ood = [r for r in (roles or []) if r in CF_OUT_OF_DOMAIN_ROLES]

    # -- decide (order matters — the H/M/L/? hierarchy). `why` is a full, project-specific
    #    sentence shown on hover. --
    surf = " + ".join(_order(surfaces))
    automation_h = bool(executed or automation_kind or is_special)
    if len(surfaces) >= 2:                                   # (1) >=2 apps in play -> H
        grade = "H"
        label = "Needs " + surf + " Copilot"
        prov = "verified from the action trace" if verified_cross_app else "inferred from the outputs"
        why = "This spans %s (%s) — no single in-app Copilot works across apps, so it needs Cowork." % (surf, prov)
    elif build:                                              # (2) code / app / skill build -> H
        grade, label = "H", "Cowork-only build"
        why = "Cowork generated or packaged code / an app / a skill here — no in-app Copilot can do that."
    elif automation_h:                                       # (3) automation / workflow -> H
        grade, label = "H", "Automation / workflow"
        kind = automation_kind or ("connector run" if executed else "specialized workflow")
        why = ("Automation / workflow work (%s) — Cowork orchestrates across items or systems; "
               "no single in-app Copilot runs it." % kind)
    elif multi_doc_synth:                                    # (4) multi-document synthesis -> H
        grade, label = "H", "Multi-document synthesis"
        why = (("Synthesizes %d source documents across %d formats — multi-doc synthesis beyond any "
                "single-app Copilot." % (n_in, len(in_fmts))) if len(in_fmts) >= 2 else
               ("Synthesizes %d sources at once — beyond a single-app Copilot." % n_in))
    elif multi_out_fmt:                                      # (5) multiple different output formats -> H
        grade, label = "H", "Multi-format output"
        why = ("Produces %d different output formats (%s) — no single in-app Copilot emits across "
               "formats." % (len(out_fmts), ", ".join(sorted(f for f in out_fmts if f))))
    elif conversational:                                     # (6) purely conversational -> L
        grade, label = "L", "Copilot chat could do it"
        why = "A purely conversational task with no build and no other app — Copilot chat would have covered it."
    elif moderate_complexity:                                # (7) moderate multi-format / volume -> M
        grade, label = "M", "Moderate fit"
        if multi_fmt_input:
            why = ("Juggles %d different input formats into the output — multi-format handling that's "
                   "more than a quick single-app task, though short of full cross-app synthesis." % len(in_fmts))
        else:
            why = ("Juggles %d files at once — the volume makes it a moderate fit, more than a "
                   "single-app Copilot task." % max(n_in, n_out))
    elif len(surfaces) == 1:                                 # (8) single app a lone Copilot could do -> L
        only = next(iter(surfaces))
        grade, label = "L", "%s Copilot could do it" % only
        if verified_apps:
            why = "The action trace shows this stayed in one app (%s) — %s Copilot alone would have covered it." % (only, only)
        else:
            why = "This reads as single-app %s work — %s Copilot alone would likely have covered it." % (only, only)
    else:                                                    # (9) truly no signal -> ?
        grade, label = "?", "Insufficient evidence"
        seen = out_ext[0] if (out_ext and out_ext[0]) else "file"
        why = ("The saved %s output maps to no clear app and there's no code / automation / "
               "multi-format signal or action trace to place the work — insufficient evidence, "
               "not guessed." % seen)

    result = {"grade": grade, "rule_grade": grade, "label": label, "why": why, "method": "rule",
              "surfaces": _order(surfaces),
              "apps": _order(surfaces),
              "verified_apps": _order(verified_apps),
              "evidence": ("insufficient" if grade == "?" else
                           "verified" if verified_apps else "inferred"),
              "verified_cross_app": bool(verified_cross_app),
              "out_of_domain_roles": ood}
    # -- LLM-review layer -----------------------------------------------------
    # The deterministic grade above is a fast, reproducible PROXY for the real
    # question ("could one in-app Copilot have done this?"). The agent may pass a
    # `review = {"grade": "H|M|L|?", "why", "label", "resolves_evidence"|"conflict"}`.
    # It OVERRIDES the rule grade (kept as rule_grade), flagged method="AI-reviewed".
    # GUARD: the review may NOT silently downgrade hard evidence — a verified cross-app
    # H cannot drop to L/? when it is backed by verified cross-app, build, or automation
    # evidence unless the review EXPLICITLY resolves the conflict.
    if review and review.get("grade") in ("H", "M", "L", "?"):
        new_grade = review["grade"]
        result["review_grade"] = new_grade
        if review.get("why"):
            result["review_why"] = review["why"]
        resolved = bool(review.get("resolves_evidence") or review.get("conflict"))
        strong_h = verified_cross_app or automation_h or build
        blocked = (not resolved) and (grade == "H" and strong_h and new_grade in ("L", "?"))
        if blocked:
            result["method"] = "AI-review-rejected"
            result["review_grade"] = new_grade
            if review.get("why"): result["review_why"] = review["why"]
            return result
        result["rule_grade"] = grade
        result["grade"] = new_grade
        if review.get("why"): result["why"] = review["why"]
        if review.get("label"):
            result["label"] = review["label"]
        elif new_grade != grade:
            # Grade moved — regenerate the label so it matches the NEW grade.
            result["label"] = CF_REVIEW_LABELS.get(new_grade, result["label"])
        result["method"] = "AI-reviewed" if new_grade != grade else "AI-confirmed"
        result["evidence"] = ("insufficient" if new_grade == "?" else
                              "verified" if verified_apps else "inferred")
    return result


def nonnegative_number(value, field, optional=False):
    if value is None and optional:
        return None
    try:
        finite = math.isfinite(value) if type(value) in (int, float) else False
    except OverflowError:
        finite = False
    if not finite or value < 0:
        raise ValueError(f"{field} must be a finite non-negative number")
    return value


def unique_tags(values, field):
    if values is None:
        return []
    if not isinstance(values, list) or any(not isinstance(v, str) for v in values):
        raise ValueError(f"{field} must be a list of strings")
    return list(dict.fromkeys(v for v in values if v))


def session_runs(s):
    raw = s.get("runs")
    if raw is None:
        raw = {c: 1 for c in (unique_tags(s.get("tasks"), "tasks") or ["general"])}
    if not isinstance(raw, dict):
        raise ValueError(f"Session {s['id']!r}: runs must be an object of non-negative integer counts")
    runs = collections.Counter()
    for category, count in raw.items():
        if not isinstance(category, str) or not category.strip():
            raise ValueError(f"Session {s['id']!r}: run categories must be non-empty strings")
        if isinstance(count, bool) or not isinstance(count, int) or count < 0:
            raise ValueError(f"Session {s['id']!r}: run count for {category!r} must be a non-negative integer")
        runs[category if category in CATS else "general"] += count
    return dict(runs)


def allocation(band, rate):
    result = {}
    for label, minutes in zip(("low", "typical", "high"), band):
        result[f"minutes_{label}"] = minutes
        result[f"hours_{label}"] = hrs(minutes)
        result[f"value_{label}"] = hrs(minutes) * rate
    return result


def add_band(totals, key, band):
    for i, minutes in enumerate(band):
        totals[key][i] += minutes


def artifact_record(a, sid, date):
    item = copy.deepcopy(a) if isinstance(a, dict) else {}
    item.update({"session": sid, "name": _name(a), "ext": _ext(a)})
    item.setdefault("type", type_label(_ext(a)))
    item.setdefault("date", date)
    return item


def source_metric(s, key, lookup):
    if key in s:
        value = s[key]
    elif s["id"] in lookup:
        value = lookup[s["id"]]
    else:
        value = lookup.get(s["id"][:8])
    return nonnegative_number(value, f"Session {s['id']!r}: {key}", optional=True)


def measured_minutes(s):
    if "measured_exec_min" in s:
        value = s["measured_exec_min"]
    elif "model" in str(s.get("timing_basis", "")).lower():
        value = None
    else:
        value = s.get("exec_min")
    return nonnegative_number(value, f"Session {s['id']!r}: measured_exec_min", optional=True)


def build_payload(d, cost_lookup=None, credits_lookup=None, interactions_lookup=None):
    meta = copy.deepcopy(d["meta"])
    sessions = d["sessions"]
    rate = nonnegative_number(meta.get("hourly_rate", 72), "hourly_rate")
    cost_lookup = cost_lookup or {}
    credits_lookup = credits_lookup or {}
    interactions_lookup = interactions_lookup or {}
    seen_ids = set()
    normalized_runs = {}
    for s in sessions:
        sid = s.get("id")
        if (not isinstance(sid, str) or not sid.strip() or sid != sid.strip()
                or any(ord(c) < 32 for c in sid)):
            raise ValueError("Each session requires a non-empty string id without surrounding whitespace or control characters")
        if sid in seen_ids:
            raise ValueError(f"Duplicate session id: {sid!r}")
        seen_ids.add(sid)
        hour = s.get("hour", 12)
        if isinstance(hour, bool) or not isinstance(hour, int) or not 0 <= hour <= 23:
            raise ValueError(f"Session {sid!r}: hour must be an integer from 0 to 23")
        normalized_runs[sid] = session_runs(s)
        for key in ("exec_min", "measured_exec_min"):
            if key in s:
                nonnegative_number(s[key], f"Session {sid!r}: {key}", optional=True)

    canonical = json.dumps(d, sort_keys=True, ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    meta.update({"schema_version": "2.0",
                 "run_id": "sha256:" + hashlib.sha256(canonical.encode("utf-8")).hexdigest()})
    defaults = {"user": "User", "email": "", "generated": "",
                "window": {"label": "Window", "from": "", "to": "", "months": 1},
                "methodology": "Cowork Time-Savings Methodology v4 + v5 artifact-scaled speed multiplier",
                "hourly_rate_default": rate, "seat_cost_month": 0,
                "artifact_identity_rule": "art_base(name): lowercase, extension removed, trailing revision/version tokens removed",
                "timing_basis": "exec_min and exec_hours are modeled; measured_exec_min is source telemetry when supplied"}
    for key, value in defaults.items():
        meta.setdefault(key, value)
    meta["hourly_rate_default"] = rate

    tasks = []; goals = []; artifacts = []; inputs_rows = []
    session_roles = []; skill_allocations = []; deliverables = []
    ccount = collections.Counter(); icount = collections.Counter()
    heat = collections.Counter(); active = set()
    in_types = collections.Counter(); out_types = collections.Counter()
    proc_count = collections.Counter(); cauthored = collections.Counter()
    cat_bands = collections.defaultdict(lambda: [0, 0, 0])
    intent_bands = collections.defaultdict(lambda: [0, 0, 0])
    proc_bands = collections.defaultdict(lambda: [0, 0, 0])
    role_bands = collections.defaultdict(lambda: [0, 0, 0])
    skill_bands = collections.defaultdict(lambda: [0, 0, 0])
    skill_sessions = collections.defaultdict(set)
    skill_deliverables = collections.defaultdict(set)
    inv = collections.defaultdict(list)
    expert = [0, 0, 0]; assist_tot = 0; conv = 0
    measured_values = []

    # Preserve the existing cross-session surface inference and Cowork-fit rules.
    deliv_surfaces = collections.defaultdict(set)
    for s in sessions:
        surf = cf_surfaces(s.get("goal", ""), s.get("outputs") or [], s.get("tasks") or [])
        for a in s.get("outputs") or []:
            deliv_surfaces[art_base(_name(a)).lower()] |= surf

    for s in sessions:
        sid = s["id"]; date = s["date"]; hour = s.get("hour", 12)
        goal = s.get("goal", "Cowork session")
        process = s.get("process", "General Productivity")
        inputs = s.get("inputs") or []
        outputs = s.get("outputs") or []
        runs = normalized_runs[sid]
        cats = list(runs)
        band = tuple(session_expert(runs, i) for i in range(3))
        e_l, e_t, e_h = band
        assist = max(ASSIST_FIXED + ASSIST_PER_ART * (len(inputs) + len(outputs)), 4)
        for i, minutes in enumerate(band):
            expert[i] += minutes
        assist_tot += assist
        active.add(date)
        proc_count[process] += 1
        add_band(proc_bands, process, band)
        measured = measured_minutes(s)
        if measured is not None:
            measured_values.append(measured)

        prof_roles = unique_tags(s.get("professional_roles"), "professional_roles") or ["Unassigned"]
        for role in prof_roles:
            share = tuple(m / len(prof_roles) for m in band)
            session_roles.append({"session": sid, "role": role, "weight": 1 / len(prof_roles),
                                  **allocation(share, rate)})
            add_band(role_bands, role, share)

        category_hours = {}
        for category, count in runs.items():
            category_band = tuple(count * CATS[category][i] for i in range(3))
            label = CATS[category][3]
            intent = INTENT[category]
            tasks.append({"session": sid, "category": category, "category_label": label,
                          "intent": intent, "runs": count, **allocation(category_band, rate)})
            ccount[category] += count
            icount[intent] += count
            add_band(cat_bands, category, category_band)
            add_band(intent_bands, intent, category_band)
            category_hours[label] = hrs(category_band[1])

        for ordinal, a in enumerate(inputs, 1):
            item = artifact_record(a, sid, date)
            item["input_id"] = f"{sid}:input:{ordinal}"
            inputs_rows.append(item)
            in_types[type_label(_ext(a))] += 1

        sess_skills = unique_tags(s.get("skills"), "skills")
        for ordinal, a in enumerate(outputs, 1):
            cauthored[out_cat(_ext(a))] += 1
            out_types[type_label(_ext(a))] += 1
            deliverable_id = f"{sid}:output:{ordinal}"
            dskills = unique_tags(a.get("skills") if isinstance(a, dict) else None, "output skills") or sess_skills or ["Unassigned"]
            output_band = tuple(m / len(outputs) for m in band)
            item = artifact_record(a, sid, date)
            item.update({"deliverable_id": deliverable_id, "artifact_key": art_base(_name(a)),
                         "skills": dskills, **allocation(output_band, rate)})
            item.update({"hours": item["hours_typical"], "value": item["value_typical"]})
            deliverables.append(item)
            artifacts.append({**copy.deepcopy(item), "artifact_id": deliverable_id})
            inv[type_label(_ext(a))].append(copy.deepcopy(item))
            for skill in dskills:
                share = tuple(m / len(dskills) for m in output_band)
                skill_allocations.append({"session": sid, "deliverable_id": deliverable_id,
                                          "scope": "deliverable", "skill": skill,
                                          "weight": 1 / len(outputs) / len(dskills),
                                          **allocation(share, rate)})
                add_band(skill_bands, skill, share)
                skill_sessions[skill].add(sid)
                skill_deliverables[skill].add(deliverable_id)

        if not outputs:
            conv += 1
            for skill in sess_skills or ["Unassigned"]:
                count = len(sess_skills) or 1
                share = tuple(m / count for m in band)
                skill_allocations.append({"session": sid, "deliverable_id": None, "scope": "session",
                                          "skill": skill, "weight": 1 / count, **allocation(share, rate)})
                add_band(skill_bands, skill, share)
                skill_sessions[skill].add(sid)

        heat[(date, hour)] += sum(runs.values())
        credits = source_metric(s, "credits", credits_lookup)
        cost = (source_metric(s, "cost_usd", cost_lookup) if "cost_usd" in s or credits is None
                else credits * 0.01)
        record = copy.deepcopy(s)
        record.update({"session": sid, "date": date, "title": goal, "process": process,
                       "source_record": copy.deepcopy(s.get("source_record", s)),
                       "value_pillar": s.get("value_pillar", "Transformation"),
                       "pillar_css": s.get("pillar_css", "trans"),
                       "job": s.get("job", "Other"), "jtbd": s.get("jtbd", ""),
                       "credits": credits, "cost_usd": cost, **allocation(band, rate),
                       "categories": sorted(category_hours), "category_hours": category_hours,
                       "category_runs": runs, "runs": copy.deepcopy(s.get("runs", runs)),
                       "n_tasks": sum(runs.values()), "artifacts": [_name(a) for a in outputs],
                       "speed_x": e_t / assist, "speed_typical": e_t / assist,
                       "speed_low": e_l / assist, "speed_high": e_h / assist,
                       "exec_min": assist, "measured_exec_min": measured, "timing_basis": "modeled",
                       "professional_roles": prof_roles, "n_inputs": len(inputs), "n_outputs": len(outputs),
                       "inputs": copy.deepcopy(inputs), "outputs": copy.deepcopy(outputs), "skills": sess_skills,
                       "total_interactions": source_metric(s, "total_interactions", interactions_lookup),
                       "cowork_fit": cowork_fit(goal, outputs, inputs, cats, prof_roles,
                           extra_surfaces=set().union(*(deliv_surfaces.get(art_base(_name(o)).lower(), set()) for o in outputs)) if outputs else set(),
                           review=s.get("cowork_fit_review"),
                           evidence={"apps": s.get("apps_accessed"), "actions": s.get("actions"),
                                     "sources": s.get("sources_reviewed")}),
                       "conversational": not outputs})
        goals.append(record)

    totals = allocation(expert, rate)
    exp_l, exp_t, exp_h = expert
    H_t = totals["hours_typical"]
    ex_h = hrs(assist_tot)
    spd_t, spd_l, spd_h = (m / assist_tot if assist_tot else 0 for m in (exp_t, exp_l, exp_h))
    measured_total = sum(measured_values) if measured_values else None
    measured_hours = hrs(measured_total) if measured_total is not None else None
    categories = []
    for key in CATS:
        if key not in cat_bands:
            continue
        categories.append({"key": key, "label": CATS[key][3], "low_per_task": CATS[key][0],
                           "typical_per_task": CATS[key][1], "high_per_task": CATS[key][2],
                           "tasks": ccount[key], "authored_outputs": cauthored[CATS[key][3]],
                           **allocation(cat_bands[key], rate)})
    categories.sort(key=lambda c: -c["hours_typical"])
    skills_aug = []
    for skill, band in sorted(skill_bands.items(), key=lambda kv: -kv[1][1]):
        amounts = allocation(band, rate)
        skills_aug.append({"skill": skill, "hours": amounts["hours_typical"], "value": amounts["value_typical"],
                           "deliverables": len(skill_deliverables[skill]), "sessions": len(skill_sessions[skill]),
                           **amounts})
    roles = []
    for role, band in sorted(role_bands.items(), key=lambda kv: -kv[1][1]):
        amounts = allocation(band, rate)
        roles.append({"role": role, "hours": amounts["hours_typical"], "value": amounts["value_typical"], **amounts})

    INV_ORDER = ["Deck", "Document", "Spreadsheet", "Web page", "Script", "Notebook", "Query", "Image", "Text", "File"]
    INV_ICON = {"Deck": "\U0001F4CA", "Document": "\U0001F4C4", "Spreadsheet": "\U0001F4C8", "Web page": "\U0001F310",
                "Script": "⚙️", "Notebook": "\U0001F4D3", "Query": "\U0001F50D", "Image": "\U0001F5BC️",
                "Text": "\U0001F4DD", "File": "\U0001F4CE"}
    inventory = [{"type": label, "icon": INV_ICON.get(label, "\U0001F4CE"), "count": len(inv[label]),
                  "items": inv[label]} for label in INV_ORDER + [k for k in inv if k not in INV_ORDER] if inv.get(label)]
    n_in = sum(in_types.values()); n_out = sum(out_types.values())
    payload = {
        "meta": meta, "source_data": copy.deepcopy(d),
        "kpis": {"sessions": len(sessions), "run_tasks": sum(ccount.values()),
                 "artifacts": len({a["artifact_key"] for a in artifacts}), "active_days": len(active),
                 "hours_saved_typical": H_t, "hours_saved_low": totals["hours_low"], "hours_saved_high": totals["hours_high"],
                 "minutes_low": exp_l, "minutes_typical": exp_t, "minutes_high": exp_h,
                 "hours_per_active_day": H_t / (len(active) or 1), "speed_multiplier": spd_t,
                 "exec_min": assist_tot, "exec_hours": ex_h, "timing_basis": "modeled",
                 "timed_sessions": len(measured_values), "measured_timed_sessions": len(measured_values),
                 "modeled_sessions": len(sessions), "measured_exec_min": measured_total,
                 "measured_exec_hours": measured_hours, "conversational_sessions": conv},
        "value": {"hourly_rate": rate, **totals, "speed_typical": spd_t, "speed_low": spd_l, "speed_high": spd_h,
                  "human_equiv_hours": H_t, "exec_min": assist_tot, "exec_hours": ex_h, "timing_basis": "modeled",
                  "measured_exec_min": measured_total, "measured_exec_hours": measured_hours},
        "leverage": {"timed_sessions": len(measured_values), "modeled_sessions": len(sessions),
                    "human_equiv_hours": H_t, "exec_min": assist_tot, "exec_hours": ex_h,
                    "speed_multiplier": spd_t, "timing_basis": "modeled",
                    "measured_exec_min": measured_total, "measured_exec_hours": measured_hours},
        "io": {"inputs_total": n_in, "outputs_total": n_out, "expert_hours": H_t,
               "expert_hours_low": totals["hours_low"], "expert_hours_high": totals["hours_high"],
               "per_deliverable": n_in / n_out if n_out else None,
               "inputs_by_type": [{"label": label, "count": count} for label, count in in_types.most_common()],
               "outputs_by_type": [{"label": label, "count": count} for label, count in out_types.most_common()]},
        "categories": categories,
        "intents": [{"intent": intent, "tasks": count, **allocation(intent_bands[intent], rate)}
                    for intent, count in icount.most_common()],
        "processes": [{"process": process, "sessions": proc_count[process], **allocation(band, rate),
                       "pct_time": band[1] / exp_t * 100 if exp_t else 0}
                      for process, band in sorted(proc_bands.items(), key=lambda kv: -kv[1][1])],
        "roles": roles, "session_roles": session_roles, "skills_augmented": skills_aug,
        "skill_allocations": skill_allocations, "inputs": inputs_rows, "inventory": inventory,
        "deliverables": sorted(deliverables, key=lambda item: -item["hours"]),
        "heatmap": [{"date": date, "hour": hour, "count": count} for (date, hour), count in sorted(heat.items())],
        "goals": sorted(goals, key=lambda g: -g["minutes_typical"]),
        "cowork_fit_summary": {label: sum(g["cowork_fit"]["grade"] == grade for g in goals)
                               for label, grade in (("H", "H"), ("M", "M"), ("L", "L"), ("U", "?"))},
        "tasks": tasks, "artifacts": artifacts,
    }
    return payload


def main(inp, out):
    with open(inp, encoding="utf-8-sig") as f:
        d = json.load(f)
    payload = build_payload(d, load_cost_lookup(), load_credits_lookup(), load_interactions_lookup())
    with open(out, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=1, allow_nan=False)
    kpis = payload["kpis"]; value = payload["value"]
    print(f"Sessions={kpis['sessions']} Tasks={kpis['run_tasks']} Outputs={payload['io']['outputs_total']} ActiveDays={kpis['active_days']}")
    print(f"Expert(human-equiv)={value['hours_typical']:.2f}h  Modeled assisted={value['exec_hours']:.2f}h  Speed={value['speed_typical']:.1f}x")
    print(f"Professional-services value=${value['value_typical']:,.2f} @${value['hourly_rate']}/hr (range ${value['value_low']:,.2f}-${value['value_high']:,.2f})")
    print("wrote", out)

if __name__=="__main__":
    ap=argparse.ArgumentParser()
    ap.add_argument("--in",dest="inp",default="working/cowork_sessions.json")
    ap.add_argument("--out",default="working/cowork_roi_data.json")
    a=ap.parse_args(); main(a.inp,a.out)
