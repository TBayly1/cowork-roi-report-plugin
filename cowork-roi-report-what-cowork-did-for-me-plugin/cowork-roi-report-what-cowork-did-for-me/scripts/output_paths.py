"""Shared report filenames derived from the invoking user's run metadata."""

from datetime import date, datetime
from pathlib import Path
import re
import unicodedata


_UNSAFE_FILENAME = re.compile(r'[<>:"/\\|?*\x00-\x1f\x7f]')
_DATE = re.compile(r"\d{4}-\d{2}-\d{2}")


def filename_user(value):
    if not isinstance(value, str) or not value.strip():
        raise ValueError("meta.user must contain the signed-in user's display name for output filenames")
    value = unicodedata.normalize("NFC", value)
    safe = _UNSAFE_FILENAME.sub("-", value)
    safe = " ".join(safe.split()).strip(" .")
    if not safe or not any(character.isalnum() for character in safe):
        raise ValueError("meta.user does not contain a usable display name for output filenames")
    try:
        units = len(safe.encode("utf-16-le")) // 2
    except UnicodeEncodeError as exc:
        raise ValueError("meta.user contains invalid Unicode for an output filename") from exc
    if units > 120:
        raise ValueError("The display name exceeds the 120-character output filename allowance")
    return safe


def report_date(value):
    if not isinstance(value, str) or not value.strip():
        raise ValueError("meta.generated must contain this run's local report date")
    value = value.strip()
    try:
        if _DATE.fullmatch(value):
            return date.fromisoformat(value).isoformat()
        if not re.match(r"\d{4}-\d{2}-\d{2}[T ]", value):
            raise ValueError("Expected an ISO date or datetime")
        # Keep the date supplied by the run, without converting to the machine's time zone.
        return datetime.fromisoformat(value.replace("Z", "+00:00")).date().isoformat()
    except ValueError as exc:
        raise ValueError("meta.generated must be a valid ISO local report date or datetime") from exc


def report_output_path(data, extension, *, directory="output", anonymous=False):
    if extension not in ("html", "xlsx", "csv"):
        raise ValueError("Report output extension must be html, xlsx or csv")
    if not isinstance(data, dict) or not isinstance(data.get("meta"), dict):
        raise ValueError("Report metadata is required for output filenames")
    meta = data["meta"]
    user = "Cowork user" if anonymous else filename_user(meta.get("user"))
    generated = report_date(meta.get("generated"))
    prefix = "cowork-sessions" if extension == "csv" else "cowork-roi-report"
    return Path(directory) / f"{prefix} - {user} - {generated}.{extension}"
