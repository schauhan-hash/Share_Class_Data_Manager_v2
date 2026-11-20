from copy import deepcopy
from datetime import date, datetime
import re

from .schema import DATE_FIELDS, FIELD_ORDER, NUMERIC_FIELDS, PERCENTAGE_FIELDS, REQUIRED_FIELDS

ISIN_PATTERN = re.compile(r"^GF555555555(\d+)$")


def clean_value(value):
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
    return str(value).strip()


def normalize_records(records):
    normalized = []
    for incoming in records:
        record = {field: clean_value(incoming.get(field, "")) for field in FIELD_ORDER}
        normalized.append(record)
    return normalized


def next_isin(records, reserved=None):
    values = [record.get("ISIN", "") for record in records]
    values.extend(reserved or [])
    suffixes = [int(match.group(1)) for value in values if (match := ISIN_PATTERN.fullmatch(value.strip()))]
    next_suffix = max(suffixes, default=0) + 1
    return f"GF555555555{next_suffix}"


def fill_missing_isins(records, existing=None):
    result = normalize_records(deepcopy(records))
    allocation_base = normalize_records(existing or [])
    reserved = []
    for record in result:
        if not record["ISIN"] or record["ISIN"].lower() == "n/a":
            generated = next_isin(allocation_base + result, reserved)
            record["ISIN"] = generated
            reserved.append(generated)
    return result


def _parse_date(value):
    if not value or value.lower() == "n/a":
        return None
    try:
        return datetime.fromisoformat(value[:10]).date()
    except ValueError:
        return False


def validate_record(record):
    issues = []
    for field in REQUIRED_FIELDS:
        if not clean_value(record.get(field, "")) or clean_value(record.get(field, "")).lower() == "n/a":
            issues.append({"severity": "Error", "field": field, "message": "Required value is missing."})

    for field in NUMERIC_FIELDS:
        value = clean_value(record.get(field, ""))
        if value and value.lower() != "n/a":
            try:
                number = float(value.replace(",", "").replace("$", ""))
                if number < 0:
                    issues.append({"severity": "Error", "field": field, "message": "Value cannot be negative."})
                if field in PERCENTAGE_FIELDS and number > 1:
                    issues.append(
                        {
                            "severity": "Warning",
                            "field": field,
                            "message": "Value is above 100%; confirm whether it is stored as a decimal.",
                        }
                    )
            except ValueError:
                issues.append({"severity": "Error", "field": field, "message": "Expected a numeric value."})

    for field in DATE_FIELDS:
        value = clean_value(record.get(field, ""))
        if value and value.lower() != "n/a" and _parse_date(value) is False:
            issues.append({"severity": "Error", "field": field, "message": "Use an ISO date such as 2026-05-01."})

    inception = _parse_date(clean_value(record.get("Share Class Inception Date", "")))
    close = _parse_date(clean_value(record.get("Share Class Close Date", "")))
    if inception not in (None, False) and close not in (None, False) and close < inception:
        issues.append(
            {
                "severity": "Error",
                "field": "Share Class Close Date",
                "message": "Close date cannot be before inception date.",
            }
        )

    soft_lock = clean_value(record.get("Soft Lock", "")).lower()
    hard_lock = clean_value(record.get("Hard Lock", "")).lower()
    lock_period = clean_value(record.get("# of Months Under Lock", ""))
    if (soft_lock == "yes" or hard_lock == "yes") and not lock_period:
        issues.append(
            {
                "severity": "Warning",
                "field": "# of Months Under Lock",
                "message": "A lock period is expected when a lock is enabled.",
            }
        )
    return issues


def validate_dataset(records):
    normalized = normalize_records(records)
    issues = []
    for index, record in enumerate(normalized, start=2):
        for issue in validate_record(record):
            issues.append({"row": index, "record": record.get("TPA Share Class ID") or record.get("ISIN"), **issue})

    for field in ("ISIN", "TPA Share Class ID"):
        values = {}
        for index, record in enumerate(normalized, start=2):
            value = clean_value(record.get(field, ""))
            if not value or value.lower() == "n/a":
                continue
            values.setdefault(value, []).append(index)
        for value, rows in values.items():
            if len(rows) > 1:
                issues.append(
                    {
                        "row": ", ".join(map(str, rows)),
                        "record": value,
                        "severity": "Error",
                        "field": field,
                        "message": f"Duplicate value: {value}.",
                    }
                )
    return issues


def diff_records(before, after):
    changes = []
    for field in FIELD_ORDER:
        old = clean_value(before.get(field, ""))
        new = clean_value(after.get(field, ""))
        if old != new:
            changes.append({"field": field, "before": old, "after": new})
    return changes

