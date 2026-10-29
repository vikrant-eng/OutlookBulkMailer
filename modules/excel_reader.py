"""
Excel import and recipient management module.
Reads .xlsx files, validates email addresses, and detects duplicates.
"""

import re
import openpyxl
from typing import Optional


_EMAIL_RE = re.compile(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$")


def validate_email(email: str) -> bool:
    return bool(_EMAIL_RE.match(str(email).strip()))


def load_excel(filepath: str) -> dict:
    """
    Load recipients from an Excel file.

    Returns a dict:
        {
            "recipients": [{"email": ..., "name": ...}, ...],
            "duplicates": ["email1", ...],
            "invalid":    ["bad@", ...],
            "errors":     ["row N: ...", ...],
        }
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active

    headers = []
    recipients = []
    duplicates = []
    invalid = []
    errors = []
    seen_emails = set()
    email_col = name_col = None

    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row_idx == 1:
            # Detect header row
            headers = [str(c).strip().lower() if c else "" for c in row]
            for i, h in enumerate(headers):
                if "email" in h:
                    email_col = i
                elif "name" in h:
                    name_col = i
            if email_col is None:
                errors.append("No 'Email' column found in the spreadsheet.")
                break
            continue

        if all(c is None for c in row):
            continue  # skip blank rows

        raw_email = str(row[email_col]).strip() if row[email_col] else ""
        raw_name = str(row[name_col]).strip() if (name_col is not None and row[name_col]) else ""

        if not raw_email:
            errors.append(f"Row {row_idx}: Empty email — skipped.")
            continue

        if not validate_email(raw_email):
            invalid.append(raw_email)
            errors.append(f"Row {row_idx}: Invalid email '{raw_email}' — skipped.")
            continue

        email_lower = raw_email.lower()
        if email_lower in seen_emails:
            duplicates.append(raw_email)
            errors.append(f"Row {row_idx}: Duplicate email '{raw_email}' — skipped.")
            continue

        seen_emails.add(email_lower)
        recipients.append({"email": raw_email, "name": raw_name})

    wb.close()
    return {
        "recipients": recipients,
        "duplicates": duplicates,
        "invalid": invalid,
        "errors": errors,
    }


def merge_recipients(existing: list[dict], new: list[dict]) -> tuple[list[dict], int]:
    """
    Merge new recipients into an existing list, skipping duplicates.
    Returns (merged_list, count_added).
    """
    seen = {r["email"].lower() for r in existing}
    added = 0
    result = list(existing)
    for r in new:
        if r["email"].lower() not in seen:
            result.append(r)
            seen.add(r["email"].lower())
            added += 1
    return result, added
