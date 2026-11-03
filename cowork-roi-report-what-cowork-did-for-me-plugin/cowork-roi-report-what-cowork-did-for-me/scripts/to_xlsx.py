#!/usr/bin/env python3
"""Export one Cowork ROI payload as four complete, data-only Excel tables.

python scripts/to_xlsx.py --data working/cowork_roi_data.json \
    --settings settings.json --email-plan working/cowork_email.json

Default output: output/cowork-roi-report - <User Name> - <YYYY-MM-DD>.xlsx
The optional settings/email-plan pair stages instructions only; no email is sent.
"""

import argparse
from collections import defaultdict
from datetime import date, datetime
from io import BytesIO
import json
import math
import os
from pathlib import Path
import re
import hashlib
import importlib.util
from uuid import uuid4
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo
from output_paths import report_output_path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.writer.excel import ExcelWriter
except ImportError as exc:
    raise ImportError(
        "XLSX export requires openpyxl. Install it with: python -m pip install openpyxl"
    ) from exc


EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_STRING_UNITS = 32_767
INTEGER = "#,##0"
DECIMAL = "#,##0.0"
MONEY = '"$"#,##0.00'
DATE = "yyyy-mm-dd"
SPEED = '0.0"x"'
_MISSING = object()
_ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")
_ISO_DATE = re.compile(r"[0-9]{4}-[0-9]{2}-[0-9]{2}")
_EMAIL = re.compile(
    r"[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+"
    r"(?:\.[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+)*@"
    r"[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?"
    r"(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+"
)
EMAIL_PLACEHOLDER = "Add Admin Email here"
EMAIL_BODY = (
    "<p>Your one-off 60-day Cowork impact report and Power BI data workbook are attached.</p>"
    "<p>Time savings and professional-services value are modeled estimates, "
    "not measured savings or billing records.</p>"
)

# Column definitions are (header, source key, type, number format, width).
# Optional numbers accept an absent key or null; other numeric fields require a
# key but preserve an explicit null as an unknown value, never a fabricated zero.
SHEETS = (
    ("Sessions", "tblSessions", "goals", (
        ("Date", "date", "date", DATE, 13),
        ("Session", "session", "id", None, 40),
        ("Goal", "title", "text", None, 65),
        ("Business process", "process", "text", None, 30),
        ("Value pillar", "value_pillar", "text", None, 25),
        ("JTBD", "jtbd", "text", None, 45),
        ("Categories", "categories", "pipe", None, 45),
        ("Run tasks", "n_tasks", "int", INTEGER, 14),
        ("Deliverables", "n_outputs", "int", INTEGER, 15),
        ("Expert minutes", "minutes_typical", "num", DECIMAL, 18),
        ("Expert hours", "hours_typical", "num", DECIMAL, 16),
        ("Value (USD)", None, "session_value", MONEY, 19),
        ("Speed multiplier", "speed_x", "optional_num", SPEED, 20),
        ("Hands-on minutes (est.)", "exec_min", "optional_num", DECIMAL, 25),
        ("Conversational (no saved file)", "conversational", "bool", None, 32),
    )),
    ("Deliverables", "tblDeliverables", "deliverables", (
        ("Date", "date", "date", DATE, 13),
        ("File name", "name", "text", None, 60),
        ("Type", "type", "text", None, 20),
        ("Extension", "ext", "text", None, 14),
        ("Session", "session", "id", None, 40),
        ("Skills", "skills", "pipe", None, 45),
        ("Expert hours", "hours", "num", DECIMAL, 16),
        ("Value (USD)", "value", "num", MONEY, 19),
    )),
    ("Categories", "tblCategories", "categories", (
        ("Category", "label", "text", None, 38),
        ("Run tasks", "tasks", "int", INTEGER, 14),
        ("Minutes saved per run (low)", "low_per_task", "num", DECIMAL, 31),
        ("Minutes saved per run (typical)", "typical_per_task", "num", DECIMAL, 34),
        ("Minutes saved per run (high)", "high_per_task", "num", DECIMAL, 32),
        ("Expert minutes", "minutes_typical", "num", DECIMAL, 18),
        ("Expert hours", "hours_typical", "num", DECIMAL, 16),
        ("Value (USD)", "value_typical", "num", MONEY, 19),
    )),
    ("Processes", "tblProcesses", "processes", (
        ("Business process", "process", "text", None, 38),
        ("Sessions", "sessions", "int", INTEGER, 14),
        ("Expert minutes", "minutes_typical", "num", DECIMAL, 18),
        ("Expert hours", "hours_typical", "num", DECIMAL, 16),
        ("Value (USD)", "value_typical", "num", MONEY, 19),
        ("% of time", "pct_time", "pct", "0.0", 14),
    )),
    ("Roles", "tblRoles", "roles", (
        ("Professional role", "role", "text", None, 38),
        ("Expert hours", "hours", "num", DECIMAL, 16),
        ("Value (USD)", "value", "num", MONEY, 19),
    )),
    ("Activity", "tblActivity", "heatmap", (
        ("Date", "date", "date", DATE, 13),
        ("Hour (local)", "hour", "hour", INTEGER, 16),
        ("Run tasks started", "count", "required_int", INTEGER, 21),
    )),
)
SUMMARY = (
    ("Window", "meta.window", "label", "text", None),
    ("From", "meta.window", "from", "date", DATE),
    ("To", "meta.window", "to", "date", DATE),
    ("Sessions", "kpis", "sessions", "int", INTEGER),
    ("Run tasks", "kpis", "run_tasks", "int", INTEGER),
    ("Deliverables (artifacts)", "kpis", "artifacts", "int", INTEGER),
    ("Active days", "kpis", "active_days", "int", INTEGER),
    ("Conversational sessions (no saved file)", "kpis", "conversational_sessions", "int", INTEGER),
    ("Expert-equivalent hours (typical)", "value", "hours_typical", "num", DECIMAL),
    ("Expert-equivalent hours (low)", "value", "hours_low", "num", DECIMAL),
    ("Expert-equivalent hours (high)", "value", "hours_high", "num", DECIMAL),
    ("Hands-on hours (estimated)", "value", "exec_hours", "optional_num", DECIMAL),
    ("Hours saved per active day", "kpis", "hours_per_active_day", "optional_num", DECIMAL),
    ("Hourly rate used (USD)", None, None, "rate", MONEY),
    ("Professional-services value (typical, USD)", "value", "value_typical", "num", MONEY),
    ("Professional-services value (low, USD)", "value", "value_low", "num", MONEY),
    ("Professional-services value (high, USD)", "value", "value_high", "num", MONEY),
    ("Speed multiplier (typical)", "value", "speed_typical", "optional_num", SPEED),
    ("Speed multiplier (low)", "value", "speed_low", "optional_num", SPEED),
    ("Speed multiplier (high)", "value", "speed_high", "optional_num", SPEED),
    ("Inputs analysed", "io", "inputs_total", "int", INTEGER),
    ("Outputs produced", "io", "outputs_total", "int", INTEGER),
)


def _object(value, path):
    if not isinstance(value, dict):
        raise ValueError(f"{path} must be an object")
    return value


def _text(value, path):
    if not isinstance(value, str):
        raise ValueError(f"{path} must be a string")
    if _ILLEGAL_XML.search(value):
        raise ValueError(f"{path} contains characters not supported by Excel/XML")
    if len(value.encode("utf-16-le")) // 2 > EXCEL_MAX_STRING_UNITS:
        raise ValueError(f"{path} exceeds Excel's 32767-character cell limit")
    return value


def _number(value, path, *, integer=False, nullable=True, maximum=None):
    if nullable and value is None:
        return None
    if type(value) not in (int, float):
        raise ValueError(f"{path} must be a finite {'integer' if integer else 'number'}")
    try:
        finite = math.isfinite(value)
    except OverflowError:
        finite = False
    if not finite or value < 0:
        raise ValueError(f"{path} must be a finite nonnegative number")
    if abs(value) > 1e308 or (type(value) is int and value > 999_999_999_999_999):
        raise ValueError(f"{path} exceeds Excel's supported numeric precision/range")
    if integer and (type(value) is not int):
        raise ValueError(f"{path} must be an integer")
    if maximum is not None and value > maximum:
        raise ValueError(f"{path} must be between 0 and {maximum}")
    return value


def _convert(value, kind, path):
    if kind == "text":
        return _text(value, path)
    if kind == "id":
        value = _text(value, path)
        if not value or value != value.strip() or any(ord(c) < 32 or ord(c) == 127 for c in value):
            raise ValueError(f"{path} must be a nonempty ID without surrounding whitespace or control characters")
        return value
    if kind == "pipe":
        if not isinstance(value, list):
            raise ValueError(f"{path} must be an array of strings")
        return _text("|".join(_text(v, f"{path}[{i}]") for i, v in enumerate(value) if v != ""), path)
    if kind == "date":
        if not isinstance(value, str) or not _ISO_DATE.fullmatch(value):
            raise ValueError(f"{path} must be a valid YYYY-MM-DD date")
        try:
            parsed = date.fromisoformat(value)
        except ValueError as exc:
            raise ValueError(f"{path} must be a valid YYYY-MM-DD date") from exc
        if parsed < date(1900, 1, 1):
            raise ValueError(f"{path} must be an Excel date on or after 1900-01-01")
        return parsed
    if kind == "bool":
        if type(value) is not bool:
            raise ValueError(f"{path} must be a Boolean")
        return value
    if kind == "optional_num" and value is _MISSING:
        return None
    return _number(
        value, path, integer=kind in ("int", "required_int", "hour"),
        nullable=kind not in ("required_int", "hour", "rate"),
        maximum=23 if kind == "hour" else (100 if kind == "pct" else None),
    )


def _legacy_rows_from(data):
    data = _object(data, "data")
    objects = {key: _object(data.get(key), key) for key in ("meta", "kpis", "value", "io")}
    objects["meta.window"] = _object(objects["meta"].get("window"), "meta.window")
    value = objects["value"]
    rate_path = "value.hourly_rate" if value.get("hourly_rate") is not None else "meta.hourly_rate_default"
    rate = _convert(
        value["hourly_rate"] if rate_path.startswith("value.") else objects["meta"].get("hourly_rate_default"),
        "rate", rate_path,
    )
    summary_rows = []
    for label, source, key, kind, fmt in SUMMARY:
        cell = rate if kind == "rate" else _convert(objects[source].get(key, _MISSING), kind, f"{source}.{key}")
        summary_rows.append((label, cell))
    if summary_rows[1][1] > summary_rows[2][1]:
        raise ValueError("meta.window.from must not be after meta.window.to")
    sheets = [("Summary", "tblSummary", ("Metric", "Value"), summary_rows, (55, 40), (None, None))]
    session_ids = set()
    for name, table, source, columns in SHEETS:
        entries = data.get(source)
        if not isinstance(entries, list):
            raise ValueError(f"{source} must be an array")
        if len(entries) + 1 > EXCEL_MAX_ROWS:
            raise ValueError(f"{source} exceeds Excel's {EXCEL_MAX_ROWS}-row sheet limit (including header)")
        rows = []
        for i, entry in enumerate(entries):
            path = f"{source}[{i}]"
            entry = _object(entry, path)
            row = []
            for header, key, kind, fmt, width in columns:
                if kind == "session_value":
                    hours = row[10]
                    amount = None if hours is None else hours * rate
                    amount = _number(amount, f"{path}.value_usd")
                    row.append(_number(entry.get("value_typical", amount), f"{path}.value_typical"))
                else:
                    row.append(_convert(entry.get(key, _MISSING), kind, f"{path}.{key}"))
            if source == "goals":
                sid = row[1]
                if sid in session_ids:
                    raise ValueError(f"{path}.session duplicates session ID {sid!r}")
                session_ids.add(sid)
            elif source == "deliverables" and row[4] not in session_ids:
                raise ValueError(f"{path}.session references missing session ID {row[4]!r}")
            rows.append(tuple(row))
        sheets.append((name, table, tuple(c[0] for c in columns), rows,
                       tuple(c[4] for c in columns), tuple(c[3] for c in columns)))
    return sheets, objects["meta.window"]


SCENARIOS = ("low", "typical", "high")
MEASURES = tuple(f"{unit}_{band}" for unit in ("minutes", "hours", "value") for band in SCENARIOS)
REL_TOL = ABS_TOL = 1e-9


class _SheetSpec(tuple):
    def __new__(cls, values, kinds):
        result = super().__new__(cls, values)
        result.kinds = kinds
        return result


def _col(header, key, kind="text?", fmt=None, width=24):
    return (header, key, kind, fmt, width)


ALLOCATIONS = tuple(
    _col(f"{unit.title()} ({band})" + (" (USD)" if unit == "value" else ""),
         f"{unit}_{band}", "num?", MONEY if unit == "value" else DECIMAL)
    for unit in ("minutes", "hours", "value") for band in SCENARIOS
)
PROVENANCE = tuple(_col(label, key) for label, key in (
    ("Source ID", "source_id"), ("Source URL", "source_url"), ("Web URL", "web_url"),
    ("Item ID", "item_id"), ("Drive ID", "drive_id"), ("App ID", "app_id"),
    ("Folder", "folder"), ("Folder ID", "folder_id"), ("Path", "path"),
    ("MIME type", "mime_type"), ("Created", "created"), ("Modified", "modified"),
))
SESSION_DETAIL = (
    _col("Job", "job"), _col("Professional roles", "professional_roles", "pipe?"),
    _col("Skills", "skills", "pipe?"), _col("Inputs", "n_inputs", "int?", INTEGER),
    _col("Timing basis", "timing_basis"),
    _col("Measured hands-on minutes", "measured_exec_min", "num?", DECIMAL),
    _col("Credits", "credits", "num?", DECIMAL),
    _col("Cost (USD)", "cost_usd", "num?", MONEY),
    _col("Total interactions", "total_interactions", "int?", INTEGER),
    _col("Request", "request"), _col("Process source", "process_source"),
    _col("Value pillar source", "value_pillar_source"), _col("JTBD source", "jtbd_source"),
    *PROVENANCE,
    *tuple(_col("Cowork fit " + field.replace("_", " "), "cowork_fit." + field, kind)
           for field, kind in (
               ("grade", "text?"), ("label", "text?"), ("why", "text?"), ("method", "text?"),
               ("rule_grade", "text?"), ("evidence", "display?"), ("apps", "pipe?"),
               ("surfaces", "pipe?"), ("verified_apps", "pipe?"),
               ("verified_cross_app", "bool?"), ("out_of_domain_roles", "pipe?"),
               ("review_applied", "bool?"), ("review_blocked", "bool?"),
               ("review_reason", "text?"), ("review_conflict", "text?"),
               ("review_grade", "text?"), ("review_why", "text?"),
           )),
    *ALLOCATIONS,
)


def _at(record, key):
    for part in key.split("."):
        if not isinstance(record, dict) or part not in record:
            return None
        record = record[part]
    return record


def _optional(value, kind, path):
    if value is None:
        return None
    base = kind.rstrip("?")
    if base == "display":
        if isinstance(value, str):
            return _text(value, path)
        if isinstance(value, list) and all(isinstance(item, str) for item in value):
            return _convert(value, "pipe", path)
        if isinstance(value, (dict, list)):
            return None
        raise ValueError(f"{path} must be text, an array, or an object")
    if base == "scalar":
        value = _scalar(value, path)
        if type(value) is int and abs(value) > 999_999_999_999_999:
            return str(value)
        return value
    if base == "pipe" and isinstance(value, (dict, list)):
        # Structured evidence is expanded into typed columns as well as RawData.
        if isinstance(value, dict) or any(isinstance(v, (dict, list)) for v in value):
            return None
    if base == "text" and isinstance(value, (dict, list)):
        return None
    return _convert(value, "optional_num" if base == "num" else base, path)


def _scalar(value, path):
    if value is None or type(value) is bool:
        return value
    if isinstance(value, str):
        return _text(value, path)
    if type(value) in (int, float):
        try:
            finite = math.isfinite(value)
        except OverflowError:
            finite = False
        if not finite or abs(value) > 1e308:
            raise ValueError(f"{path} must be finite and within Excel's numeric range")
        return value
    raise ValueError(f"{path} must contain only JSON objects, arrays, and scalars")


def _leaf_rows(value, pointer=""):
    """Typed JSON Pointer leaves, including container markers and numeric literals."""
    if isinstance(value, dict):
        if not value:
            yield (pointer, "object", None, None, None, None)
        for key, child in value.items():
            _text(key, f"{pointer} object key")
            escaped = key.replace("~", "~0").replace("/", "~1")
            yield from _leaf_rows(child, pointer + "/" + escaped)
    elif isinstance(value, list):
        if not value:
            yield (pointer, "array", None, None, None, None)
        for i, child in enumerate(value):
            yield from _leaf_rows(child, pointer + "/" + str(i))
    else:
        _scalar(value, pointer)
        if value is None:
            yield (pointer, "null", None, None, None, None)
        elif type(value) is bool:
            yield (pointer, "boolean", None, None, value, None)
        elif isinstance(value, str):
            yield (pointer, "string", value, None, None, None)
        else:
            # Excel cannot represent arbitrary JSON integers exactly. The literal
            # also makes every floating-point audit value exactly reconstructible.
            numeric = value if type(value) is float or abs(value) <= 999_999_999_999_999 else None
            yield (pointer, "integer" if type(value) is int else "number",
                   None, numeric, None, repr(value))


def _flatten(record, prefix="", excluded=()):
    fields = {}
    for key, value in record.items():
        if key in excluded:
            continue
        path = f"{prefix}.{key}" if prefix else key
        if isinstance(value, dict):
            fields.update(_flatten(value, path))
        elif isinstance(value, list):
            if all(isinstance(v, str) for v in value):
                fields[path] = _convert(value, "pipe", path)
            else:
                for i, child in enumerate(value):
                    index = f"{path}[{i}]"
                    if isinstance(child, dict):
                        fields.update(_flatten(child, index))
                    elif isinstance(child, list):
                        fields.update(_flatten({"items": child}, index))
                    else:
                        fields[index] = child
        else:
            fields[path] = value
    return fields


def _records_sheet(name, table, records, columns, run_id, *, extras=True, excluded=()):
    columns = list(columns)
    # All row tables are snapshot-scoped, including documentation.
    if not any(c[0] == "Run ID" for c in columns):
        columns.append(_col("Run ID", "__run_id"))
    known = {c[1] for c in columns}
    flattened = [_flatten(record, excluded=excluded) for record in records]
    if extras:
        dynamic = sorted({key for record in flattened for key in record if key not in known})
        used_headers = {c[0].casefold() for c in columns}
        for key in dynamic:
            header = "Detail: " + key
            while header.casefold() in used_headers:
                header = "Detail: " + header
            used_headers.add(header.casefold())
            columns.append(_col(header, key, "scalar?"))
    rows = []
    for i, (record, flat) in enumerate(zip(records, flattened)):
        flat["__run_id"] = run_id
        row = []
        for header, key, kind, fmt, width in columns:
            supplied = _at(record, key)
            value = supplied if kind.rstrip("?") in ("pipe", "display") and supplied is not None else (
                flat[key] if key in flat else supplied)
            row.append(_optional(value, kind, f"{name}[{i}].{key}"))
        rows.append(tuple(row))
    return _SheetSpec((name, table, tuple(c[0] for c in columns), rows,
                       tuple(c[4] for c in columns), tuple(c[3] for c in columns)),
                      tuple(c[2] for c in columns))


def _entries(data, key, *, required=False):
    if key not in data and not required:
        return []
    entries = data.get(key)
    if not isinstance(entries, list):
        raise ValueError(f"{key} must be an array")
    for i, entry in enumerate(entries):
        _object(entry, f"{key}[{i}]")
    return entries


def _identity(value, path):
    return _convert(value, "id", path)


def _close(actual, expected, path):
    _number(actual, path, nullable=False)
    _number(expected, path, nullable=False)
    if not math.isclose(actual, expected, rel_tol=REL_TOL, abs_tol=ABS_TOL):
        raise ValueError(f"{path} does not reconcile: {actual!r} != {expected!r}")


def _unique(entries, keys, path):
    seen = set()
    for i, row in enumerate(entries):
        values = tuple(row.get(key) for key in keys)
        for key, value in zip(keys, values):
            if not isinstance(value, (str, int)) or type(value) is bool or value == "":
                raise ValueError(f"{path}[{i}].{key} must be a nonempty key")
        if values in seen:
            raise ValueError(f"{path}[{i}] duplicates key {values!r}")
        seen.add(values)


def _validate_facts(data, rate, modern, quality):
    goals = data["goals"]
    for goal in goals:
        if goal.get("cowork_fit") is not None:
            _object(goal["cowork_fit"], f"goals.{goal['session']}.cowork_fit")
    sessions = {g["session"]: g for g in goals}
    deliveries, input_ids = {}, set()
    for source, id_key in (("deliverables", "deliverable_id"), ("inputs", "input_id")):
        for i, row in enumerate(_entries(data, source, required=modern)):
            path = f"{source}[{i}]"
            sid = _identity(row.get("session"), path + ".session")
            if sid not in sessions:
                raise ValueError(f"{path}.session references missing session {sid!r}")
            if modern or row.get(id_key) is not None:
                key = _identity(row.get(id_key), path + "." + id_key)
                target = deliveries if source == "deliverables" else input_ids
                if key in target:
                    raise ValueError(f"{path} duplicates {id_key} {key!r}")
                if source == "deliverables":
                    deliveries[key] = row
                else:
                    input_ids.add(key)
            if source == "deliverables" and modern:
                if not _text(row.get("artifact_key"), path + ".artifact_key"):
                    raise ValueError(f"{path}.artifact_key must be nonempty")
    for source, keys in (
        ("tasks", ("session", "category")),
        ("session_roles", ("session", "role")),
        ("skill_allocations", ("session", "deliverable_id", "scope", "skill")),
    ):
        entries = _entries(data, source, required=modern)
        if source not in data:
            quality("warning", "legacy_missing_fact", source, "Absent in legacy payload; no allocations invented.")
        seen = set()
        for i, row in enumerate(entries):
            path = f"{source}[{i}]"
            sid = _identity(row.get("session"), path + ".session")
            if sid not in sessions:
                raise ValueError(f"{path}.session references missing session {sid!r}")
            for key in keys:
                if key == "deliverable_id" and row.get(key) is None:
                    continue
                value = row.get(key)
                if key == "session":
                    continue
                if not isinstance(value, str) or not value:
                    raise ValueError(f"{path}.{key} must be nonempty text")
            key = tuple(row.get(k) for k in keys)
            if key in seen:
                raise ValueError(f"{path} duplicates fact key {key!r}")
            seen.add(key)
            if source == "skill_allocations":
                did, scope = row.get("deliverable_id"), row.get("scope")
                if scope not in ("session", "deliverable"):
                    raise ValueError(f"{path}.scope must be session or deliverable")
                if (scope == "deliverable") != (did is not None):
                    raise ValueError(f"{path} scope and deliverable_id disagree")
                if did is not None and (did not in deliveries or deliveries[did]["session"] != sid):
                    raise ValueError(f"{path}.deliverable_id references a missing or different session's deliverable")
            if modern and source in ("session_roles", "skill_allocations"):
                weight = _number(row.get("weight"), path + ".weight", nullable=False, maximum=1)
                if weight == 0:
                    raise ValueError(f"{path}.weight must be positive")
            if source == "tasks" and (modern or "runs" in row):
                _number(row.get("runs"), path + ".runs", integer=True, nullable=False)
                if modern:
                    for key in ("category_label", "intent"):
                        if not _text(row.get(key), path + "." + key):
                            raise ValueError(f"{path}.{key} must be nonempty")
    for source, keys in (("categories", ("key",) if modern else ("label",)),
                          ("processes", ("process",)), ("roles", ("role",)),
                          ("skills_augmented", ("skill",)), ("intents", ("intent",)),
                          ("heatmap", ("date", "hour"))):
        _unique(_entries(data, source), keys, source)
    if not modern:
        quality("warning", "legacy_schema", "meta.schema_version",
                "Legacy payload: precise/full facts and source coverage cannot be certified.")
        return
    _object(data.get("source_data"), "source_data")
    for source, keys in (("kpis", ("sessions", "run_tasks", "artifacts", "active_days", "conversational_sessions")),
                          ("io", ("inputs_total", "outputs_total"))):
        for key in keys:
            _number(data[source].get(key), source + "." + key, integer=True, nullable=False)
    for i, row in enumerate(data["categories"]):
        _number(row.get("tasks"), f"categories[{i}].tasks", integer=True, nullable=False)
    for source in ("intents", "inventory", "skills_augmented", "artifacts"):
        _entries(data, source, required=True)
    for source in ("goals", "tasks", "session_roles", "skill_allocations", "deliverables"):
        for i, row in enumerate(data[source]):
            for key in MEASURES:
                _number(row.get(key), f"{source}[{i}].{key}", nullable=False)
            for band in SCENARIOS:
                _close(row[f"hours_{band}"], row[f"minutes_{band}"] / 60, f"{source}[{i}].hours_{band}")
                _close(row[f"value_{band}"], row[f"hours_{band}"] * rate, f"{source}[{i}].value_{band}")
    for source in ("tasks", "session_roles", "skill_allocations"):
        grouped = defaultdict(list)
        for row in data[source]:
            grouped[row["session"]].append(row)
        for sid, goal in sessions.items():
            for key in MEASURES:
                _close(math.fsum(r[key] for r in grouped[sid]), goal[key], f"{source}.{sid}.{key}")
            if source in ("session_roles", "skill_allocations"):
                _close(math.fsum(r["weight"] for r in grouped[sid]), 1, f"{source}.{sid}.weight")
                for row in grouped[sid]:
                    for key in MEASURES:
                        _close(row[key], goal[key] * row["weight"], f"{source}.{sid}.{key}.weight")
    for sid, goal in sessions.items():
        outputs = [d for d in data["deliverables"] if d["session"] == sid]
        inputs = [d for d in data["inputs"] if d["session"] == sid]
        _close(len(outputs), goal["n_outputs"], f"goals.{sid}.n_outputs")
        _close(len(inputs), _number(goal.get("n_inputs"), f"goals.{sid}.n_inputs", integer=True, nullable=False),
               f"goals.{sid}.n_inputs")
        _close(sum(t["runs"] for t in data["tasks"] if t["session"] == sid),
               goal["n_tasks"], f"goals.{sid}.n_tasks")
        if goal["conversational"] != (not outputs):
            raise ValueError(f"goals.{sid}.conversational disagrees with deliverables")
        for key in MEASURES:
            if outputs:
                _close(math.fsum(d[key] for d in outputs), goal[key], f"deliverables.{sid}.{key}")
        for d in outputs:
            allocations = [a for a in data["skill_allocations"] if a.get("deliverable_id") == d["deliverable_id"]]
            for key in MEASURES:
                _close(math.fsum(a[key] for a in allocations), d[key],
                       f"skill_allocations.{d['deliverable_id']}.{key}")
    for band in SCENARIOS:
        for unit in ("hours", "value"):
            key = f"{unit}_{band}"
            expected = _number(data["value"].get(key), "value." + key, nullable=False)
            _close(math.fsum(g[key] for g in goals), expected, "goals.total." + key)
    total = data["value"]["hours_typical"]
    for source, hours, amount in (("categories", "hours_typical", "value_typical"),
                                   ("processes", "hours_typical", "value_typical"),
                                   ("roles", "hours", "value"),
                                   ("skills_augmented", "hours", "value")):
        rows = _entries(data, source, required=True)
        for key, expected in ((hours, total), (amount, data["value"]["value_typical"])):
            for i, row in enumerate(rows):
                _number(row.get(key), f"{source}[{i}].{key}", nullable=False)
            _close(math.fsum(row[key] for row in rows), expected, source + ".total." + key)
    for fact, fact_key, aggregate, aggregate_key in (
        ("tasks", "category", "categories", "key"),
        ("session_roles", "role", "roles", "role"),
        ("skill_allocations", "skill", "skills_augmented", "skill"),
        ("goals", "process", "processes", "process"),
    ):
        rollups = {row[aggregate_key]: row for row in data[aggregate]}
        grouped = defaultdict(list)
        for row in data[fact]:
            key = row[fact_key]
            if key not in rollups:
                if any(row[measure] != 0 for measure in MEASURES):
                    raise ValueError(f"{fact}.{key} references missing nonzero {aggregate} aggregate")
                quality("info", "zero_group_without_aggregate", f"{fact}.{key}",
                        f"Explicit zero allocation retained in facts; {aggregate} does not materialize this zero group.")
            grouped[key].append(row)
        for group, summary in rollups.items():
            if aggregate == "processes":
                _close(len(grouped[group]), summary.get("sessions"), f"processes.{group}.sessions")
            for measure in MEASURES:
                alias = ({"hours_typical": "hours", "value_typical": "value"}.get(measure)
                         if aggregate in ("roles", "skills_augmented") else None)
                if measure in summary:
                    expected = _number(summary[measure], f"{aggregate}.{group}.{measure}", nullable=False)
                    _close(math.fsum(row[measure] for row in grouped[group]), expected,
                           f"{aggregate}.{group}.{measure}")
                if alias is not None:
                    _close(math.fsum(row[measure] for row in grouped[group]), summary[alias],
                           f"{aggregate}.{group}.{alias}")
    for row in data["deliverables"]:
        for alias, measure in (("hours", "hours_typical"), ("value", "value_typical")):
            _close(row[alias], row[measure], f"deliverables.{row['deliverable_id']}.{alias}")
    for source in ("inputs", "deliverables"):
        for row in data[source]:
            entity = row.get("input_id", row.get("deliverable_id"))
            if not any(row.get(k) for k in ("source_id", "item_id", "id", "source_ids")):
                quality("warning", "missing_source_id", str(entity), "No supplied source artifact identity.")
            if not any(row.get(k) for k in ("source_url", "web_url", "webUrl", "url", "source_urls")):
                quality("warning", "missing_source_url", str(entity), "No supplied source artifact URL.")
    for source, rows in (("artifacts", data["artifacts"]),):
        ids = set()
        for i, row in enumerate(rows):
            sid = row.get("session")
            if sid is not None and sid not in sessions:
                raise ValueError(f"{source}[{i}] references missing session")
            did = row.get("deliverable_id")
            if did is not None and did not in deliveries:
                raise ValueError(f"{source}[{i}] references missing deliverable")
            aid = row.get("artifact_id")
            if aid is not None:
                _identity(aid, f"{source}[{i}].artifact_id")
                if aid in ids:
                    raise ValueError(f"{source}[{i}] duplicates artifact_id")
                ids.add(aid)
    for i, group in enumerate(data["inventory"]):
        if "items" in group:
            if not isinstance(group["items"], list):
                raise ValueError("inventory.items must be an array")
            if "count" in group:
                _close(len(group["items"]), _number(group["count"], "inventory.count", integer=True, nullable=False),
                       f"inventory[{i}].count")
            for item in group["items"]:
                if isinstance(item, dict) and item.get("deliverable_id") is not None:
                    if item["deliverable_id"] not in deliveries:
                        raise ValueError(f"inventory[{i}] references missing deliverable")
    for direction in ("inputs", "outputs"):
        rows = data["io"].get(direction + "_by_type")
        if not isinstance(rows, list):
            raise ValueError(f"io.{direction}_by_type must be an array")
        for i, row in enumerate(rows):
            _object(row, f"io.{direction}_by_type[{i}]")
            _number(row.get("count"), f"io.{direction}_by_type[{i}].count", integer=True, nullable=False)
        _unique(rows, ("label",), f"io.{direction}_by_type")
        _close(sum(row["count"] for row in rows), data["io"][direction + "_total"], f"io.{direction}_by_type.total")
    _close(sum(row["tasks"] for row in data["categories"]), data["kpis"]["run_tasks"], "categories.tasks")
    for i, row in enumerate(data["intents"]):
        _number(row.get("tasks"), f"intents[{i}].tasks", integer=True, nullable=False)
        _close(sum(t["runs"] for t in data["tasks"] if t["intent"] == row["intent"]),
               row["tasks"], f"intents.{row['intent']}.tasks")
    if {t["intent"] for t in data["tasks"] if t["runs"] != 0} - {r["intent"] for r in data["intents"]}:
        raise ValueError("tasks.intent references missing intents key")
    _close(sum(row["tasks"] for row in data["intents"]), data["kpis"]["run_tasks"], "intents.tasks")
    _close(sum(row["count"] for row in data["heatmap"]), data["kpis"]["run_tasks"], "heatmap.count")
    fit = _object(data.get("cowork_fit_summary"), "cowork_fit_summary")
    if set(fit) - {"H", "M", "L", "U", "?"} or ("U" in fit and "?" in fit):
        raise ValueError("cowork_fit_summary has unknown or duplicate normalized grade keys")
    for grade, count in fit.items():
        _number(count, "cowork_fit_summary." + grade, integer=True, nullable=False)
        _close(sum((g.get("cowork_fit") or {}).get("grade") == ("?" if grade == "U" else grade) for g in goals),
               count, "cowork_fit_summary." + grade)
    _close(sum(fit.values()), len(goals), "cowork_fit_summary.total")
    for actual, expected, path in (
        (len(goals), data["kpis"]["sessions"], "kpis.sessions"),
        (sum(t["runs"] for t in data["tasks"]), data["kpis"]["run_tasks"], "kpis.run_tasks"),
        (len({d["artifact_key"] for d in data["deliverables"]}), data["kpis"]["artifacts"], "kpis.artifacts"),
        (len(data["inputs"]), data["io"]["inputs_total"], "io.inputs_total"),
        (len(data["deliverables"]), data["io"]["outputs_total"], "io.outputs_total"),
    ):
        _close(actual, expected, path)
    quality("info", "reconciled", "facts",
            "All scenario allocations reconcile to sessions and Summary; relative/absolute tolerance 1e-9.")


def _reference_modules():
    # These are bundled local constants; importing does not run either CLI.
    modules = []
    for name in ("build_report", "compute"):
        spec = importlib.util.spec_from_file_location(
            "cowork_xlsx_reference_" + name, Path(__file__).with_name(name + ".py"))
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        modules.append(module)
    return tuple(modules)


def _detail_views_from(data):
    base, window = _legacy_rows_from(data)
    # Validate every leaf, including unknown future fields, before any writes.
    leaves = list(_leaf_rows(data))
    meta = data["meta"]
    version = meta.get("schema_version")
    if version is not None and version not in ("1.0", "2.0"):
        raise ValueError("meta.schema_version must be 1.0 or 2.0 when supplied")
    modern = version == "2.0"
    run_id = meta.get("run_id")
    if run_id is None and not modern:
        canonical = json.dumps(data, sort_keys=True, separators=(",", ":"), ensure_ascii=False)
        run_id = "legacy-" + hashlib.sha256(canonical.encode("utf-8")).hexdigest()
    _identity(run_id, "meta.run_id")
    rate = data["value"].get("hourly_rate")
    if rate is None:
        rate = meta["hourly_rate_default"]
    issues = []

    def quality(severity, code, entity, detail):
        issues.append({"severity": severity, "code": code, "entity": entity, "detail": detail})

    _validate_facts(data, rate, modern, quality)
    records = {}
    for name, table, source, columns in SHEETS:
        entries = _entries(data, source, required=True)
        merged = []
        for entry in entries:
            source_record = entry.get("source_record")
            if source_record is not None:
                _object(source_record, source + ".source_record")
            merged.append({**(source_record or {}), **entry})
        records[name] = merged
    summary = base[0]
    sheets = [(summary[0], summary[1], summary[2] + ("Run ID",),
               [row + (run_id,) for row in summary[3]], summary[4] + (30,), summary[5] + (None,))]
    excluded = ("source_record", "inputs", "outputs", "actions", "apps_accessed")
    for old, (name, table, source, columns) in zip(base[1:], SHEETS):
        extra_columns = SESSION_DETAIL if name == "Sessions" else (
            (_col("Deliverable ID", "deliverable_id"), _col("Artifact key", "artifact_key"),
             *PROVENANCE, *ALLOCATIONS) if name == "Deliverables" else
            ((_col("Category key", "key"), *ALLOCATIONS) if name == "Categories" else
             ALLOCATIONS if name in ("Processes", "Roles") else ())
        )
        prefix_keys = tuple(c[1] for c in columns if c[1] is not None)
        detail = _records_sheet(name, table, records[name], extra_columns, run_id,
                                excluded=prefix_keys + (excluded if name == "Sessions" else ()))
        # The old columns, including per-row number formats, remain an exact prefix.
        sheets.append(_SheetSpec((name, table, old[2] + detail[2],
                                  [a + b for a, b in zip(old[3], detail[3])],
                                  old[4] + detail[4], old[5] + detail[5]),
                                 tuple(c[2] for c in columns) + detail.kinds))
    generated = meta.get("generated")
    generated_date = None
    if isinstance(generated, str) and len(generated) >= 10 and _ISO_DATE.fullmatch(generated[:10]):
        generated_date = generated[:10]
    run = {**meta, "generated_date": generated_date, "rate": rate,
           "snapshot_policy": "One snapshot; do not append overlapping lookbacks.",
           "timing_disclosure": "Expert time and default hands-on clock are modeled, not measured.",
           "citation_status": "Bundled published claims; not newly verified by this export.",
           "kpis": data["kpis"], "value": data["value"], "leverage": data.get("leverage"),
           "io": {k: v for k, v in data["io"].items() if k not in ("inputs_by_type", "outputs_by_type")}}
    sheets.append(_records_sheet("Run", "tblRun", [run], (
        _col("User", "user"), _col("Email", "email"), _col("Generated date", "generated_date", "date?", DATE),
        _col("Generated timestamp", "generated"), _col("Window", "window.label"),
        _col("From", "window.from", "date?", DATE), _col("To", "window.to", "date?", DATE),
        _col("Hourly rate (USD)", "rate", "num?", MONEY), _col("Schema version", "schema_version"),
        _col("Methodology", "methodology"), _col("Snapshot policy", "snapshot_policy"),
        _col("Timing disclosure", "timing_disclosure"), _col("Citation status", "citation_status"),
    ), run_id))
    fact_specs = (
        ("SessionCategories", "tasks", (
            _col("Session", "session", "id?"), _col("Category key", "category"),
            _col("Category", "category_label"), _col("Intent", "intent"), _col("Runs", "runs", "int?", INTEGER),
            *ALLOCATIONS)),
        ("SessionRoles", "session_roles", (
            _col("Session", "session", "id?"), _col("Role", "role"), _col("Weight", "weight", "num?", DECIMAL),
            *ALLOCATIONS)),
        ("SkillAllocations", "skill_allocations", (
            _col("Session", "session", "id?"), _col("Deliverable ID", "deliverable_id", "id?"),
            _col("Scope", "scope"), _col("Skill", "skill"), _col("Weight", "weight", "num?", DECIMAL),
            *ALLOCATIONS)),
        ("Skills", "skills_augmented", (
            _col("Skill", "skill"), _col("Expert hours", "hours", "num?", DECIMAL),
            _col("Value (USD)", "value", "num?", MONEY), _col("Deliverables", "deliverables", "int?", INTEGER),
            _col("Sessions", "sessions", "int?", INTEGER), *ALLOCATIONS)),
        ("Inputs", "inputs", (
            _col("Input ID", "input_id", "id?"), _col("Session", "session", "id?"),
            _col("File name", "name"), _col("Extension", "ext"), _col("Type", "type"),
            _col("Date", "date", "date?", DATE), *PROVENANCE)),
    )
    for name, source, columns in fact_specs:
        sheets.append(_records_sheet(name, "tbl" + name, _entries(data, source), columns, run_id))
    apps, actions = [], []
    for goal in records["Sessions"]:
        sid = goal["session"]
        fit = goal.get("cowork_fit") or {}
        _object(fit, f"goals.{sid}.cowork_fit")
        for flag in ("verified_cross_app", "review_applied", "review_blocked"):
            if fit.get(flag) is not None:
                _convert(fit[flag], "bool", f"goals.{sid}.cowork_fit.{flag}")
        if fit.get("grade") not in (None, "H", "M", "L", "?"):
            raise ValueError(f"goals.{sid}.cowork_fit.grade must be H, M, L or ?")
        known_apps = {}
        for field, origin, verified in (
            ("apps", "cowork_fit.apps (inferred)", False),
            ("verified_apps", "cowork_fit.verified_apps", True),
            ("apps_accessed", "apps_accessed (observed)", True),
        ):
            values = goal.get(field) if field == "apps_accessed" else fit.get(field)
            if values is None:
                continue
            if not isinstance(values, list):
                raise ValueError(f"goals.{sid}.{field} must be an array")
            for item in values:
                if isinstance(item, str):
                    app, metadata = item, {}
                elif isinstance(item, dict):
                    app = next((item[k] for k in ("app", "name", "app_name", "displayName", "id", "app_id")
                                if isinstance(item.get(k), str) and item[k]), None)
                    if app is None:
                        quality("warning", "unidentified_app", sid,
                                "An apps_accessed entry has no supplied app name/ID; the compact References table omits raw payload details.")
                        continue
                    metadata = item
                else:
                    raise ValueError(f"goals.{sid}.{field} entries must be strings or objects")
                _text(app, f"goals.{sid}.{field}")
                key = app.casefold()
                row = known_apps.setdefault(key, {"session": sid, "app": app, "verified": False,
                                                   "origins": [], "observed": False})
                row["verified"] = row["verified"] or verified
                row["observed"] = row["observed"] or field == "apps_accessed"
                if origin not in row["origins"]:
                    row["origins"].append(origin)
                if metadata:
                    row.setdefault("supplied_metadata", []).append(metadata)
        apps.extend(known_apps.values())
        supplied_actions = goal.get("actions")
        if supplied_actions is not None:
            if not isinstance(supplied_actions, list):
                raise ValueError(f"goals.{sid}.actions must be an array")
            for i, action in enumerate(supplied_actions, 1):
                if not isinstance(action, (str, dict)):
                    raise ValueError(f"goals.{sid}.actions entries must be strings or objects")
                row = {"action": action} if isinstance(action, str) else {**action}
                row.update(session=sid, ordinal=i)
                actions.append(row)
        for key, label in (("source_id", "source IDs"), ("source_url", "source URLs")):
            if not goal.get(key):
                quality("warning", "missing_" + key, sid,
                        f"No canonical {label} supplied; alternate provenance remains in session details, not References.")
        if not goal.get("jtbd"):
            quality("warning", "blank_jtbd", sid, "JTBD absent or blank; not inferred by exporter.")
        for key in ("process_source", "value_pillar_source", "jtbd_source"):
            value = goal.get(key)
            if value is None or "default" in str(value).lower() or "fallback" in str(value).lower():
                quality("warning", "classification_provenance", sid,
                        f"{key} is absent/default/fallback; supplied classification is not verified source telemetry.")
        missing = [k for k in ("measured_exec_min", "credits", "cost_usd", "total_interactions")
                   if goal.get(k) is None]
        if missing:
            quality("warning", "missing_telemetry", sid, "Unknown: " + ", ".join(missing))
        if goal.get("apps_accessed") is None or goal.get("actions") is None or goal.get("sources_reviewed") is None:
            quality("warning", "unknown_activity_coverage", sid,
                    "One or more app/action/source lists were not supplied; absence is not zero observed activity.")
    quality("info", "source_coverage_unknown", "source_data",
            "Export completeness applies to the supplied payload only; upstream source/discovery coverage is unknown.")
    sheets.append(_records_sheet("SessionApps", "tblSessionApps", apps, (
        _col("Session", "session", "id?"), _col("App", "app"), _col("Verified", "verified", "bool?"),
        _col("Observed apps_accessed", "observed", "bool?"), _col("Evidence origin", "origins", "pipe?"),
    ), run_id))
    sheets.append(_records_sheet("SessionActions", "tblSessionActions", actions, (
        _col("Session", "session", "id?"), _col("Ordinal", "ordinal", "int?", INTEGER),
        _col("Action", "action"), _col("Name", "name"), _col("App", "app"),
        _col("Target", "target"), _col("Status", "status"), _col("Timestamp", "timestamp"),
    ), run_id))
    sheets.append(_records_sheet("Intents", "tblIntents", _entries(data, "intents"), (
        _col("Intent", "intent"), _col("Run tasks", "tasks", "int?", INTEGER),
    ), run_id))
    inventory = []
    for group in _entries(data, "inventory"):
        items = group.get("items")
        if items is None:
            inventory.append({k: v for k, v in group.items() if k != "items"})
        else:
            if not isinstance(items, list):
                raise ValueError("inventory.items must be an array")
            for i, item in enumerate(items, 1):
                item = {"name": item} if isinstance(item, str) else _object(item, "inventory.items")
                inventory.append({"type": group.get("type"), "ordinal": i, **item})
    sheets.append(_records_sheet("Inventory", "tblInventory", inventory, (
        _col("Type", "type"), _col("Ordinal in type", "ordinal", "int?", INTEGER),
        _col("File name", "name"), _col("Date", "date", "date?", DATE),
    ), run_id))
    io_types = []
    for direction in ("inputs", "outputs"):
        values = data["io"].get(direction + "_by_type", [])
        if not isinstance(values, list):
            raise ValueError(f"io.{direction}_by_type must be an array")
        for entry in values:
            io_types.append({**_object(entry, "io type"), "direction": direction})
    sheets.append(_records_sheet("IOTypes", "tblIOTypes", io_types, (
        _col("Direction", "direction"), _col("Type", "label"), _col("Count", "count", "int?", INTEGER),
    ), run_id))
    fit_summary = _object(data.get("cowork_fit_summary", {}), "cowork_fit_summary")
    fit_rows = [{"grade": "?" if grade == "U" else grade, "count": count}
                for grade, count in fit_summary.items()]
    sheets.append(_records_sheet("FitSummary", "tblFitSummary", fit_rows, (
        _col("Cowork-fit grade", "grade"), _col("Sessions", "count", "int?", INTEGER),
    ), run_id))
    report, compute = _reference_modules()
    by_label = {v[3]: k for k, v in compute.CATS.items()}
    active = {row.get("key", by_label.get(row.get("label"))): row.get("tasks") for row in data["categories"]}
    if modern:
        for row in data["tasks"]:
            if row["category"] not in compute.CATS:
                raise ValueError(f"tasks.category references unknown methodology key {row['category']!r}")
            if row["category_label"] != compute.CATS[row["category"]][3]:
                raise ValueError("tasks.category_label disagrees with the bundled methodology")
    methodology = []
    for label, low, typical, high in report.ALL_CATS:
        key = by_label[label]
        if tuple(compute.CATS[key][:3]) != (low, typical, high):
            raise ValueError(f"Bundled research/model bands disagree for {label}")
        methodology.append({"key": key, "label": label, "low": low, "typical": typical, "high": high,
                            "unit": "minutes per run task", "tasks": active.get(key, 0),
                            "active": bool(active.get(key, 0)), "rationale": report.CAT_SOURCES[label],
                            "citation_status": "Bundled research claim; not newly verified."})
    sheets.append(_records_sheet("Methodology", "tblMethodology", methodology, (
        _col("Category key", "key"), _col("Category", "label"), _col("Low", "low", "num?", DECIMAL),
        _col("Typical", "typical", "num?", DECIMAL), _col("High", "high", "num?", DECIMAL),
        _col("Unit", "unit"), _col("Run tasks", "tasks", "int?", INTEGER),
        _col("Active", "active", "bool?"), _col("Research rationale", "rationale"),
        _col("Citation status", "citation_status"),
    ), run_id, extras=False))
    sources = [{"category": label, "title": title, "url": url}
               for label, links in report.CAT_LINKS.items() for title, url in links]
    sources.append({"category": None, "title": "Project repository", "url": report.GITHUB_URL})
    sheets.append(_records_sheet("Sources", "tblSources", sources, (
        _col("Category", "category"), _col("Published title", "title"), _col("URL", "url"),
    ), run_id, extras=False))
    sheets.append(_records_sheet("Glossary", "tblGlossary",
                                [{"term": term, "definition": definition} for term, definition in report.GLOSSARY_TERMS],
                                (_col("Term", "term"), _col("Definition", "definition")), run_id, extras=False))
    assumptions = (
        ("Expert clock", "Sum(runs × bundled category minutes); hours = minutes / 60. Modeled, not measured."),
        ("Value", "Expert-equivalent hours × hourly rate; not billing records or measured cash savings."),
        ("Hourly rate (USD)", rate),
        ("Assisted fixed minutes", compute.ASSIST_FIXED),
        ("Assisted minutes per artifact", compute.ASSIST_PER_ART),
        ("Assisted clock", "max(fixed + per-artifact × (input occurrences + output occurrences), 4) minutes; modeled."),
        ("Measured clock", "measured_exec_min is separate nullable telemetry; missing is unknown, not zero."),
        ("Speed", "Expert modeled minutes / modeled hands-on minutes; rate independent."),
        ("Role allocation", "Equal session shares across distinct supplied roles; no tags => Unassigned."),
        ("Deliverable allocation", "Equal session shares across output occurrences, not deduplicated artifacts."),
        ("Skill allocation", "Equal shares of each deliverable across its skills (session fallback); chat-only at session grain; no tags => Unassigned. Weight is a fraction of the entire session, as for role weights."),
        ("Deduplication", "Artifact key supplied by compute.art_base(): lower-case name without extension, iterative version/final/draft suffixes stripped by the bundled rule. Only distinct-artifact KPI is deduplicated; occurrence facts are not."),
        ("Percent of time", "Processes.% of time is a number on the 0..100 scale, not an Excel fraction."),
        ("Rounding", "No early rounding of allocations or money; Excel formats round display only. Reconciliation tolerance: relative/absolute 1e-9."),
        ("Snapshot", "One run per workbook. Never append overlapping lookbacks and then sum snapshots."),
        ("Research", "Citations are bundled published claims, not newly verified by this export."),
        ("Payload audit reconstruction", "References payload-audit records use JSON Pointer paths to retain array indices; null/object/array types preserve null/empty containers. A string-typed blank Text value is the empty string. Use Number literal for exact JSON numeric reconstruction; large integers may have blank Number value."),
        ("Missing activity", "Absent apps/actions/sources is unknown coverage, not observed zero. Inferred apps are not verified unless separately supplied as verified/apps_accessed."),
        ("Safety", "Data-only literal text; no formulas, macros, hyperlinks or external links."),
    )
    sheets.append(_records_sheet("Assumptions", "tblAssumptions",
                                [{"term": term, "text": v if isinstance(v, str) else None,
                                  "number": v if type(v) in (int, float) else None} for term, v in assumptions],
                                (_col("Assumption", "term"), _col("Explanation", "text"),
                                 _col("Number", "number", "num?", DECIMAL)), run_id, extras=False))
    relationships = []
    session_children = ("Deliverables", "SessionCategories", "SessionRoles", "SkillAllocations", "Inputs",
                        "SessionApps", "SessionActions")
    for sheet in sheets:
        if sheet[0] != "Run" and sheet[0] not in session_children:
            relationships.append({"from": sheet[0], "fk": "Run ID", "to": "Run", "pk": "Run ID",
                                  "note": "Many-to-one; filter Run → child only. Aggregate tables remain separate measures."})
    for name in session_children:
        relationships.append({"from": name, "fk": "Run ID + Session", "to": "Sessions",
                              "pk": "Run ID + Session",
                              "note": "Many-to-one; Sessions → fact only. Create a collision-safe composite key in Power Query; Session alone is unsafe across snapshots."})
    relationships.append({"from": "SkillAllocations", "fk": "Run ID + Deliverable ID (non-null only)",
                          "to": "Deliverables", "pk": "Run ID + Deliverable ID",
                          "note": "Optional alternative join, inactive when Sessions path is active; never create two active filter paths. Session-scope/chat rows have no deliverable."})
    sheets.append(_records_sheet("Relationships", "tblRelationships", relationships, (
        _col("Child table", "from"), _col("Foreign key", "fk"), _col("Parent table", "to"),
        _col("Unique key", "pk"), _col("Cardinality", "cardinality"),
        _col("Filter direction and cautions", "note"),
    ), run_id, extras=False))
    # Set the declared cardinality without creating Excel relationships or formulas.
    rel = sheets[-1]
    cardinality_column = rel[2].index("Cardinality")
    rel[3][:] = [row[:cardinality_column] + ("many-to-one",) + row[cardinality_column + 1:] for row in rel[3]]
    sheets.append(_records_sheet("DataQuality", "tblDataQuality", issues, (
        _col("Severity", "severity"), _col("Code", "code"), _col("Entity", "entity"), _col("Detail", "detail"),
    ), run_id, extras=False))
    raw_records = [{"path": p, "type": t, "text": text, "number": n, "boolean": b, "literal": literal}
                   for p, t, text, n, b, literal in leaves]
    sheets.append(_records_sheet("RawData", "tblRawData", raw_records, (
        _col("Path (JSON Pointer)", "path"), _col("Type", "type"), _col("Text value", "text"),
        _col("Number value", "number", "scalar?"), _col("Boolean value", "boolean", "bool?"),
        _col("Number literal", "literal"),
    ), run_id, extras=False))
    dictionary_columns = (
        _col("Worksheet", "sheet"), _col("Table", "table"), _col("Column", "column"),
        _col("Data type", "type"), _col("Unit", "unit"), _col("Grain / key", "grain"),
        _col("Description", "description"),
    )
    dictionary = _records_sheet("DataDictionary", "tblDataDictionary", [], dictionary_columns, run_id, extras=False)
    sheets.append(dictionary)
    grains = {
        "Summary": "One legacy metric per snapshot; Metric",
        "Run": "One snapshot; Run ID",
        "Sessions": "One session; Run ID + Session",
        "Deliverables": "One output occurrence; Run ID + Deliverable ID",
        "SessionCategories": "One category per session; Run ID + Session + Category key",
        "SessionRoles": "One role per session; Run ID + Session + Role",
        "SkillAllocations": "One skill per session/deliverable scope; Run ID + Session + Deliverable ID + Scope + Skill",
        "Inputs": "One input occurrence; Run ID + Input ID",
        "SessionApps": "One distinct app per session; Run ID + Session + App (case-insensitive)",
        "SessionActions": "One supplied action in original order; Run ID + Session + Ordinal",
        "RawData": "One typed scalar leaf or empty container; Run ID + Path (JSON Pointer)",
        "DataDictionary": "One column description; Worksheet + Table + Column",
        "Categories": "One category aggregate; Run ID + Category key; do not sum with SessionCategories",
        "Processes": "One process aggregate; Run ID + Business process; do not sum with Sessions",
        "Roles": "One role aggregate; Run ID + Professional role; do not sum with SessionRoles",
        "Skills": "One skill aggregate; Run ID + Skill; do not sum with SkillAllocations",
        "Activity": "One activity bucket; Run ID + Date + Hour (local)",
        "Intents": "One intent aggregate; Run ID + Intent",
        "Inventory": "One supplied inventory occurrence; Run ID + Type + Ordinal in type; group counts retained in RawData",
        "IOTypes": "One direction/type aggregate; Run ID + Direction + Type",
        "FitSummary": "One grade aggregate; Run ID + Cowork-fit grade",
        "Methodology": "One bundled category band; Run ID + Category key",
        "Sources": "One bundled source/category association; Run ID + Category + Published title + URL",
        "Glossary": "One bundled glossary definition; Run ID + Term",
        "Assumptions": "One assumption; Run ID + Assumption",
        "Relationships": "One proposed relationship; Child table + Foreign key + Parent table",
        "DataQuality": "One diagnostic per entity/code; Run ID + Entity + Code (multiple details possible)",
    }
    dictionary_rows = []
    for spec in sheets:
        sheet_name, table, headers, rows, widths, formats = spec
        for i, (header, fmt) in enumerate(zip(headers, formats)):
            types = {("boolean" if type(r[i]) is bool else "date" if isinstance(r[i], date)
                      else "number" if type(r[i]) in (int, float) else "text")
                     for r in rows if r[i] is not None}
            kind = getattr(spec, "kinds", ("text",) * len(headers))[i].rstrip("?")
            declared = ("boolean" if kind == "bool" else "date" if kind == "date"
                        else "number" if kind in ("num", "int", "required_int", "hour", "pct", "session_value")
                        else "scalar (text/number/boolean)" if kind == "scalar" else "text")
            dtype = " / ".join(sorted(types)) or declared
            unit = ("USD" if fmt == MONEY else "percent (0..100)" if header == "% of time"
                    else "fraction of session (0..1)" if header == "Weight"
                    else "minutes per run task" if sheet_name == "Methodology" and header in ("Low", "Typical", "High")
                    else "hours" if "hours" in header.lower()
                    else "minutes" if "minutes" in header.lower()
                    else "multiplier" if fmt == SPEED else "date" if fmt == DATE else None)
            description = "Supplied field; null/blank means unknown or not applicable. Literal text, no formulas."
            if sheet_name in ("Categories", "Processes", "Roles", "Skills", "Intents", "Inventory", "IOTypes", "FitSummary", "Activity"):
                description += " Aggregate/view: do not add to allocation facts or join on nonunique labels."
            if header == "Run ID":
                description = "Snapshot hash; use with occurrence/session IDs for joins across non-overlapping runs."
            if header == "% of time":
                description += " Stored on 0..100 scale, not 0..1."
            if "est." in header or header == "Timing basis":
                description += " Modeled clock is not measured telemetry."
            if fmt in (MONEY, DECIMAL, SPEED):
                description += " Full precision stored; rounding is display only."
            dictionary_rows.append({"sheet": sheet_name, "table": table, "column": header, "type": dtype,
                                    "unit": unit, "grain": grains.get(sheet_name, "Snapshot-scoped view/documentation; see named columns"),
                                    "description": description})
    sheets[-1] = _records_sheet("DataDictionary", "tblDataDictionary", dictionary_rows,
                                dictionary_columns, run_id, extras=False)
    for name, table, headers, rows, widths, formats in sheets:
        if len(rows) + 1 > EXCEL_MAX_ROWS:
            raise ValueError(f"{name} exceeds Excel's row limit (including header)")
        if len(headers) > 16_384:
            raise ValueError(f"{name} exceeds Excel's column limit")
        if len(set(h.casefold() for h in headers)) != len(headers):
            raise ValueError(f"{name} contains duplicate table headers")
        for header in headers:
            _text(header, f"{name} header")
    return sheets, window


REFERENCE_COLUMNS = (
    _col("Reference ID", "reference_id"),
    _col("Record type", "record_type"),
    _col("Record ID", "record_id"),
    _col("Field", "field"),
    _col("Value type", "value_type"),
    _col("Text value", "text"),
    _col("Number value", "number", "scalar?", "#,##0.############"),
    _col("Boolean value", "boolean", "bool?"),
    _col("Date value", "date", "date?", DATE),
    _col("Number literal", "literal"),
    _col("Path (JSON Pointer)", "pointer"),
    _col("Session", "session", "id?"),
    _col("Run ID", "run_id"),
)
REFERENCE_TYPES = {
    "Summary": "summary",
    "Categories": "category-summary",
    "Processes": "process-summary",
    "Roles": "role-summary",
    "Activity": "activity",
    "Run": "run",
    "Skills": "skill-summary",
    "SessionApps": "session-app",
    "SessionActions": "session-action",
    "Intents": "intent-summary",
    "Inventory": "inventory",
    "IOTypes": "io-type",
    "FitSummary": "fit-summary",
    "Methodology": "methodology",
    "Sources": "source",
    "Glossary": "glossary",
    "Assumptions": "assumption",
    "DataQuality": "data-quality",
}


def _key(*parts):
    return json.dumps(parts, ensure_ascii=False, separators=(",", ":"))


def _record_id(*parts):
    return hashlib.sha256(_key(*parts).encode("utf-8")).hexdigest()


def _stack_views(name, views, prefix_columns, prefix_values):
    """Union already-validated rows without converting their native cell types."""
    columns = list(prefix_columns)
    seen = {column[0].casefold(): i for i, column in enumerate(columns)}
    for _, view in views:
        kinds = getattr(view, "kinds", ("scalar?",) * len(view[2]))
        for header, kind, fmt, width in zip(view[2], kinds, view[5], view[4]):
            if header.casefold() not in seen:
                seen[header.casefold()] = len(columns)
                columns.append(_col(header, header, kind, fmt, width))
    rows = []
    for discriminator, view in views:
        for ordinal, values in enumerate(view[3], 1):
            source = dict(zip(view[2], values))
            source.update(prefix_values(discriminator, source, ordinal))
            rows.append(tuple(source.get(column[0]) for column in columns))
    return _SheetSpec(
        (name, "tbl" + name, tuple(c[0] for c in columns), rows,
         tuple(c[4] for c in columns), tuple(c[3] for c in columns)),
        tuple(c[2] for c in columns),
    )


def _reference_value(value):
    result = {"text": None, "number": None, "boolean": None, "date": None, "literal": None}
    if value is None:
        result["value_type"] = "null"
    elif type(value) is bool:
        result.update(value_type="boolean", boolean=value)
    elif isinstance(value, date):
        result.update(value_type="date", date=value)
    elif type(value) in (int, float):
        result.update(value_type="integer" if type(value) is int else "number",
                      number=value if type(value) is float or abs(value) <= 999_999_999_999_999 else None,
                      literal=repr(value))
    elif isinstance(value, str):
        result.update(value_type="string", text=value)
    else:
        raise ValueError("Reference values must be validated scalar or date cells")
    return result


def _reference_spec(records):
    # Date cells have already been parsed in the domain views; keep them native.
    return _SheetSpec(
        ("References", "tblReferences", tuple(c[0] for c in REFERENCE_COLUMNS),
         [tuple(record.get(c[1]) for c in REFERENCE_COLUMNS) for record in records],
         tuple(c[4] for c in REFERENCE_COLUMNS), tuple(c[3] for c in REFERENCE_COLUMNS)),
        tuple(c[2] for c in REFERENCE_COLUMNS),
    )


def _rows_from(data, *, settings=None):
    views, window = _detail_views_from(data)
    by_name = {view[0]: view for view in views}
    run_id = by_name["Run"][3][0][by_name["Run"][2].index("Run ID")]
    sessions = by_name["Sessions"]
    source_sessions = {goal["session"]: goal for goal in data["goals"]}
    columns = (_col("Session key", "session_key"), _col("Hour (local)", "hour", "int?", INTEGER))
    session_rows = []
    for values in sessions[3]:
        row = dict(zip(sessions[2], values))
        hour = source_sessions[row["Session"]].get("hour")
        if hour is not None:
            _number(hour, f"goals.{row['Session']}.hour", integer=True, nullable=False, maximum=23)
        session_rows.append((_key(run_id, row["Session"]), hour, *values))
    sessions = _SheetSpec(
        ("Sessions", "tblSessions", tuple(c[0] for c in columns) + sessions[2],
         session_rows, tuple(c[4] for c in columns) + sessions[4],
         tuple(c[3] for c in columns) + sessions[5]),
        tuple(c[2] for c in columns) + sessions.kinds,
    )

    def artifact_values(direction, row, ordinal):
        supplied_id = row.get("Input ID" if direction == "input" else "Deliverable ID")
        occurrence = supplied_id if supplied_id is not None else _key(row["Session"], direction, ordinal)
        return {
            "Artifact ID": occurrence, "Direction": direction,
            "Session key": _key(run_id, row["Session"]),
            "Artifact occurrence key": _key(run_id, direction, occurrence),
            "Occurrence ID basis": "supplied occurrence ID" if supplied_id is not None else
                                   "legacy session/direction/ordinal; not a source item ID",
        }

    artifacts = _stack_views(
        "Artifacts", [("input", by_name["Inputs"]), ("output", by_name["Deliverables"])],
        (_col("Artifact ID", "artifact_id"), _col("Direction", "direction"),
         _col("Session key", "session_key"), _col("Artifact occurrence key", "artifact_key"),
         _col("Occurrence ID basis", "id_basis")), artifact_values,
    )

    def allocation_values(kind, row, ordinal):
        sid, did = row["Session"], row.get("Deliverable ID")
        key = (row.get("Category key") if kind == "category" else row.get("Role") if kind == "role" else
               _key(did, row.get("Scope"), row.get("Skill")))
        return {
            "Allocation ID": _record_id(run_id, kind, sid, key), "Allocation type": kind,
            "Session key": _key(run_id, sid),
            "Artifact occurrence key": _key(run_id, "output", did) if did is not None else None,
        }

    allocations = _stack_views(
        "Allocations", [(kind, by_name[name]) for kind, name in
                        (("category", "SessionCategories"), ("role", "SessionRoles"),
                         ("skill", "SkillAllocations"))],
        (_col("Allocation ID", "allocation_id"), _col("Allocation type", "allocation_type"),
         _col("Session key", "session_key"), _col("Artifact occurrence key", "artifact_key")),
        allocation_values,
    )
    records = []

    def add_fields(record_type, record_id, fields, session=None):
        for field, value in fields.items():
            if field == "Run ID":
                continue
            records.append({
                "reference_id": _record_id(run_id, record_type, record_id, field),
                "record_type": record_type, "record_id": record_id, "field": field,
                "session": session, "run_id": run_id, **_reference_value(value),
            })

    for name, record_type in REFERENCE_TYPES.items():
        view = by_name[name]
        for ordinal, values in enumerate(view[3], 1):
            fields = dict(zip(view[2], values))
            add_fields(record_type, str(ordinal), fields, fields.get("Session"))
    raw = by_name["RawData"]
    for values in raw[3]:
        row = dict(zip(raw[2], values))
        pointer = row["Path (JSON Pointer)"]
        records.append({
            "reference_id": _record_id(run_id, "payload-audit", pointer),
            "record_type": "payload-audit", "record_id": pointer, "field": "Value",
            "value_type": row["Type"], "text": row["Text value"], "number": row["Number value"],
            "boolean": row["Boolean value"], "literal": row["Number literal"],
            "pointer": pointer, "run_id": run_id,
        })
    add_fields("configuration", "workbook", {
        "Workbook schema": "four-tables-1.0", "Payload schema": data["meta"].get("schema_version"),
        "Lookback days": 60, "Execution": "once",
        "Email settings supplied": settings is not None,
        **(settings or {}),
    })
    relationships = (
        ("Artifacts", "Session key", "Sessions", "Session key", "active",
         "Single direction Sessions -> Artifacts. Input/output records share this session relation."),
        ("Allocations", "Session key", "Sessions", "Session key", "active",
         "Single direction Sessions -> Allocations. Filter Allocation type before summing any measure."),
        ("Allocations", "Artifact occurrence key", "Artifacts", "Artifact occurrence key", "inactive",
         "Optional output-skill drilldown only. Keep inactive while the session paths are active."),
    )
    for ordinal, (child, fk, parent, pk, active, note) in enumerate(relationships, 1):
        add_fields("relationship", str(ordinal), {
            "Child table": child, "Foreign key": fk, "Parent table": parent, "Unique key": pk,
            "Cardinality": "many-to-one", "Active state": active, "Filter direction and cautions": note,
        })
    for ordinal, (term, explanation) in enumerate((
        ("Allocation type", "category, role and skill are separate decompositions of the same session effort. "
         "Each type reconciles independently; never SUM all types together."),
        ("References", "Long-form typed fields. Filter Record type and pivot Field by Run ID + Record ID "
         "to recover each reference dataset. Do not SUM arbitrary Number value rows."),
        ("Join keys", "Session key and Artifact occurrence key are collision-safe JSON arrays. "
         "Use these exact keys in Power BI; do not join titles or filenames."),
        ("Artifact grain", "Inputs carry no allocated effort. Output occurrences retain precise shares and "
         "duplicates. Use Direction=output and distinct Artifact key for the deduplicated artifact KPI."),
        ("Audit reconstruction", "Filter Record type=payload-audit. Path (JSON Pointer) and Value type retain "
         "null/empty containers and array order. A blank string cell means empty string; Number literal "
         "preserves exact numbers that Excel cannot store precisely."),
    ), 1):
        add_fields("model-guidance", str(ordinal), {"Topic": term, "Explanation": explanation})
    specs = [sessions, artifacts, allocations, _reference_spec(records)]
    grains = {
        "Sessions": "One session per snapshot; Session key",
        "Artifacts": "One input/output occurrence; Artifact occurrence key",
        "Allocations": "One category/role/skill allocation; Allocation ID; filter Allocation type",
        "References": "One typed field per reference record or payload leaf; Reference ID",
    }
    for spec in specs:
        for header, kind, fmt in zip(spec[2], spec.kinds, spec[5]):
            base_kind = kind.rstrip("?")
            dtype = ("Boolean" if base_kind == "bool" else "Date" if base_kind == "date" else
                     "Number" if base_kind in ("int", "num", "hour", "required_int", "pct", "session_value") else
                     "Typed scalar" if base_kind == "scalar" else "Text")
            unit = ("USD" if fmt == MONEY else "hours" if "hours" in header.lower() else
                    "minutes" if "minutes" in header.lower() else "multiplier" if fmt == SPEED else
                    "local hour (0..23)" if header == "Hour (local)" else None)
            add_fields("data-dictionary", _key(spec[0], header), {
                "Worksheet": spec[0], "Table": spec[1], "Column": header, "Data type": dtype,
                "Unit": unit, "Grain / key": grains[spec[0]],
                "Description": ("Full precision retained; formats affect display only. "
                                "Unknown/not-applicable values remain blank."),
            })
    # Preserve only the EAV slices read by the existing Power BI query.
    report_reference_fields = [
        record for record in records
        if record["record_type"] in {"summary", "activity"}
        or (record["record_type"] == "configuration"
            and record["field"] in {"email_to", "Payload schema"})
    ]
    specs[-1] = _reference_spec(report_reference_fields)
    for name, table, headers, rows, widths, formats in specs:
        if len(rows) + 1 > EXCEL_MAX_ROWS or len(headers) > 16_384:
            raise ValueError(f"{name} exceeds Excel's row/column limit")
        if len(set(h.casefold() for h in headers)) != len(headers):
            raise ValueError(f"{name} contains duplicate table headers")
        for header in headers:
            _text(header, f"{name} header")
        for row in rows:
            for value in row:
                if isinstance(value, str):
                    _text(value, f"{name} cell")
    return specs, window


def validate_settings(settings):
    """Return validated settings; accept only an explicit Boolean and recipient."""
    settings = _object(settings, "settings")
    expected = {"email_enabled", "email_to"}
    if set(settings) != expected:
        raise ValueError("settings must contain exactly email_enabled and email_to (no unknown keys)")
    if type(settings["email_enabled"]) is not bool:
        raise ValueError("settings.email_enabled must be a Boolean")
    recipient = settings["email_to"]
    if not settings["email_enabled"] and recipient == EMAIL_PLACEHOLDER:
        return dict(settings)
    if (
        not isinstance(recipient, str)
        or len(recipient) > 254
        or not _EMAIL.fullmatch(recipient)
        or len(recipient.split("@")[0]) > 64
    ):
        raise ValueError("settings.email_to must be one plain valid email address (no whitespace, display name, or list)")
    return dict(settings)


def _email_plan(settings, window, out_path):
    return {
        "enabled": settings["email_enabled"],
        "to": settings["email_to"],
        "subject": f"My Copilot Cowork impact - Last 60 days ({window['from']} to {window['to']})",
        "body": EMAIL_BODY,
        "content_type": "HTML",
        "direct_attachment_file_paths": [
            out_path.with_suffix(".html").as_posix(), out_path.as_posix(),
        ],
    }


def _workbook_bytes(sheets):
    workbook = Workbook()
    workbook.remove(workbook.active)
    workbook.properties.creator = "Cowork ROI exporter"
    workbook.properties.lastModifiedBy = "Cowork ROI exporter"
    workbook.properties.created = datetime(2000, 1, 1)
    workbook.properties.modified = datetime(2000, 1, 1)
    workbook.calculation = None
    for name, table_name, headers, rows, widths, formats in sheets:
        sheet = workbook.create_sheet(name)
        sheet.freeze_panes = "A2"
        sheet.sheet_view.showGridLines = False
        for r, values in enumerate((headers, *rows), 1):
            for c, value in enumerate(values, 1):
                cell = sheet.cell(r, c, value)
                if isinstance(value, str):
                    # openpyxl otherwise treats "=" as a formula and "#N/A" as an error.
                    cell.data_type = "s"
                if r == 1:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="1565A8")
                    cell.alignment = Alignment(vertical="center")
                elif formats[c - 1]:
                    cell.number_format = formats[c - 1]
            if name == "Summary" and r > 1:
                sheet.cell(r, 2).number_format = SUMMARY[r - 2][4] or "General"
        sheet.row_dimensions[1].height = 30
        for i, width in enumerate(widths, 1):
            sheet.column_dimensions[get_column_letter(i)].width = width
        table = Table(displayName=table_name, ref=f"A1:{get_column_letter(len(headers))}{len(rows) + 1}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        sheet.add_table(table)
    raw = BytesIO()
    # ExcelWriter avoids Workbook.save's automatic current-time metadata.
    with ZipFile(raw, "w", ZIP_DEFLATED) as archive:
        ExcelWriter(workbook, archive).save()
    workbook.close()
    stable = BytesIO()
    with ZipFile(raw) as source, ZipFile(stable, "w") as destination:
        for name in sorted(source.namelist()):
            info = ZipInfo(name, date_time=(2000, 1, 1, 0, 0, 0))
            info.create_system = 0
            info.external_attr = 0o600 << 16
            destination.writestr(info, source.read(name), compress_type=ZIP_DEFLATED, compresslevel=9)
    return stable.getvalue()


def _stage(path, content):
    staged = path.with_name(f".{path.name}.{uuid4().hex}.stage")
    try:
        with staged.open("xb") as handle:
            handle.write(content)
            handle.flush()
            os.fsync(handle.fileno())
    except OSError:
        staged.unlink(missing_ok=True)
        raise
    return staged


def _atomic_write(outputs):
    staged, backups, replaced = {}, {}, []
    recovery_files = set()
    try:
        for path, content in outputs:
            path.parent.mkdir(parents=True, exist_ok=True)
            staged[path] = _stage(path, content)
            backups[path] = _stage(path, path.read_bytes()) if path.exists() else None
        for path, content in outputs:
            os.replace(staged[path], path)
            replaced.append(path)
    except OSError as error:
        recovery_errors = []
        for path in reversed(replaced):
            try:
                if backups[path] is None:
                    path.unlink(missing_ok=True)
                else:
                    os.replace(backups[path], path)
            except OSError as recovery_error:
                backup = backups[path]
                if backup is not None:
                    recovery_files.add(backup)
                recovery_errors.append(f"{path}: {recovery_error}; original backup: {backup}")
        if recovery_errors:
            raise OSError("Output commit and rollback failed: " + "; ".join(recovery_errors)) from error
        raise
    finally:
        for path in (*staged.values(), *backups.values()):
            if path is not None and path not in recovery_files:
                path.unlink(missing_ok=True)


def build(data, out_path, *, settings=None, email_plan_path=None):
    """Validate, then atomically replace the latest workbook and optional plan.

    All payload/settings validation and serialization happens before any output
    is written. If committing the pair fails, previous snapshots are restored.
    """
    if (settings is None) != (email_plan_path is None):
        raise ValueError("settings and email_plan_path must be provided together")
    if settings is not None:
        settings = validate_settings(settings)
    out_path = Path(out_path)
    plan_path = Path(email_plan_path) if email_plan_path is not None else None
    if plan_path is not None and out_path.resolve() == plan_path.resolve():
        raise ValueError("workbook and email plan must use different output paths")
    sheets, window = _rows_from(data, settings=settings)
    outputs = [(out_path, _workbook_bytes(sheets))]
    if plan_path is not None:
        plan = _email_plan(settings, window, out_path)
        outputs.append((plan_path, (json.dumps(plan, indent=2, ensure_ascii=False) + "\n").encode("utf-8")))
    _atomic_write(outputs)


def _json_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key: {key}")
        result[key] = value
    return result


def _reject_constant(value):
    raise ValueError(f"JSON contains non-finite number: {value}")


def _read_json(path):
    with Path(path).open(encoding="utf-8-sig") as handle:
        return json.load(handle, object_pairs_hook=_json_object, parse_constant=_reject_constant)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data", default="working/cowork_roi_data.json")
    parser.add_argument("--out", default=None,
                        help="explicit output path; default adds the user's display name and local report date")
    parser.add_argument("--settings")
    parser.add_argument("--email-plan")
    args = parser.parse_args(argv)
    if (args.settings is None) != (args.email_plan is None):
        parser.error("--settings and --email-plan must be supplied together")
    try:
        settings = _read_json(args.settings) if args.settings is not None else None
        if settings is not None:
            validate_settings(settings)
        data = _read_json(args.data)
        out_path = args.out or report_output_path(data, "xlsx")
        build(data, out_path, settings=settings, email_plan_path=args.email_plan)
    except (ValueError, OSError) as exc:
        parser.exit(2, f"error: {exc}\n")
    scope = ("schema 2.0 report snapshot" if data["meta"].get("schema_version") == "2.0"
             else "legacy report snapshot; missing fields remain unavailable")
    print(f"wrote {out_path} (4 data-only tables: Sessions, Artifacts, Allocations, References; {scope})")
    if args.email_plan is not None:
        print(f"staged {args.email_plan}; no email was sent")


if __name__ == "__main__":
    main()
