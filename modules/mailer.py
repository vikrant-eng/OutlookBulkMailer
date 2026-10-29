"""
Email sender module for Outlook Bulk Mailer.
Calls POST https://graph.microsoft.com/v1.0/me/sendMail

Security:
  - Access token is passed only in the Authorization header — never logged.
  - No client secret is used or stored.

Error handling:
  - 202 Accepted  → success
  - 401           → TokenExpiredError (caller should re-login)
  - 403           → PermissionError  (missing Mail.Send scope or admin block)
  - 429           → ThrottleError    (caller should back off)
  - 5xx           → retried with exponential back-off
  - Timeout/Connection errors → retried

References:
  https://learn.microsoft.com/en-us/graph/api/user-sendmail
  https://learn.microsoft.com/en-us/graph/errors
"""

import base64
import os
import time
import logging
import requests
from typing import Optional

log = logging.getLogger(__name__)

GRAPH_SEND_URL = "https://graph.microsoft.com/v1.0/me/sendMail"

DEFAULT_DELAY_SECONDS = 2   # polite inter-send delay
MAX_RETRIES           = 3
# Exponential back-off delays between retries (seconds)
RETRY_BACKOFF         = [5, 15, 45]

# HTTP status codes that are worth retrying (transient server errors)
_RETRYABLE_STATUS = {500, 502, 503, 504}


# ──────────────────────────────────────────────────────────────────────────────
# Personalisation
# ──────────────────────────────────────────────────────────────────────────────

def personalise(template: str, name: str) -> str:
    """Replace {{Name}} / {{name}} placeholders with the recipient's name."""
    result = template.replace("{{Name}}", name)
    result = result.replace("{{name}}", name)
    return result


# ──────────────────────────────────────────────────────────────────────────────
# Attachment encoding
# ──────────────────────────────────────────────────────────────────────────────

def _encode_attachment(filepath: str) -> dict:
    """Base64-encode a file for the Graph API fileAttachment format."""
    filename = os.path.basename(filepath)
    with open(filepath, "rb") as f:
        content = base64.b64encode(f.read()).decode("utf-8")
    return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": filename,
        "contentBytes": content,
    }


# ──────────────────────────────────────────────────────────────────────────────
# Payload builder
# ──────────────────────────────────────────────────────────────────────────────

def build_message(to_email: str, to_name: str, subject: str, body: str,
                  attachment_paths: Optional[list[str]] = None) -> dict:
    """
    Build the sendMail JSON payload for the Microsoft Graph API.

    Args:
        to_email:         Recipient email address.
        to_name:          Recipient display name (used for {{Name}} substitution).
        subject:          Email subject (may contain {{Name}}).
        body:             Email HTML body (may contain {{Name}}).
        attachment_paths: Optional list of local file paths to attach.

    Returns:
        dict suitable for json= parameter of requests.post()
    """
    attachments: list[dict] = []
    if attachment_paths:
        for path in attachment_paths:
            try:
                attachments.append(_encode_attachment(path))
            except Exception as exc:
                log.warning("Could not encode attachment '%s': %s", path, exc)

    payload: dict = {
        "message": {
            "subject": personalise(subject, to_name),
            "body": {
                "contentType": "HTML",
                "content": personalise(body, to_name),
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": to_email,
                        "name": to_name or to_email,
                    }
                }
            ],
        },
        "saveToSentItems": True,
    }
    if attachments:
        payload["message"]["attachments"] = attachments
    return payload


# ──────────────────────────────────────────────────────────────────────────────
# Error helpers
# ──────────────────────────────────────────────────────────────────────────────

def _parse_graph_error(resp: requests.Response) -> str:
    """Extract the human-readable error message from a Graph error response."""
    try:
        body = resp.json()
        err = body.get("error", {})
        code = err.get("code", "")
        msg  = err.get("message", resp.text[:200])
        return f"{code}: {msg}" if code else msg
    except Exception:
        return resp.text[:200]


# ──────────────────────────────────────────────────────────────────────────────
# Main send function
# ──────────────────────────────────────────────────────────────────────────────

def send_email(
    access_token: str,
    to_email: str,
    to_name: str,
    subject: str,
    body: str,
    attachment_paths: Optional[list[str]] = None,
    delay: float = DEFAULT_DELAY_SECONDS,
) -> dict:
    """
    Send a single personalised email via POST /v1.0/me/sendMail.

    Returns:
        {"success": True}
        {"success": False, "error": "<human-readable message>", "code": <int>}

    Special "code" values callers should handle:
        401  → token expired; re-login required
        403  → permission denied (Mail.Send scope missing or admin blocked)
        429  → throttled; "retry_after" key contains seconds to wait
    """
    # NOTE: access_token only goes into the header — never printed or logged.
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    payload = build_message(to_email, to_name, subject, body, attachment_paths)

    for attempt in range(MAX_RETRIES):
        try:
            resp = requests.post(
                GRAPH_SEND_URL,
                headers=headers,
                json=payload,
                timeout=30,
            )

            # ── 202 Accepted → email queued successfully ─────────────
            if resp.status_code == 202:
                time.sleep(delay)
                return {"success": True}

            # ── 401 Unauthorized → token expired ────────────────────
            if resp.status_code == 401:
                return {
                    "success": False,
                    "error": (
                        "Access token expired (HTTP 401).\n"
                        "Please go to Settings and sign in again."
                    ),
                    "code": 401,
                }

            # ── 403 Forbidden → missing scope or admin policy ────────
            if resp.status_code == 403:
                detail = _parse_graph_error(resp)
                return {
                    "success": False,
                    "error": (
                        f"Permission denied (HTTP 403): {detail}\n\n"
                        "Possible causes:\n"
                        "• 'Mail.Send' permission not added to the app registration\n"
                        "• Admin consent required in your organization\n"
                        "• Sending blocked by a mailbox or tenant policy"
                    ),
                    "code": 403,
                }

            # ── 429 Too Many Requests → throttled ───────────────────
            if resp.status_code == 429:
                retry_after = int(resp.headers.get("Retry-After", RETRY_BACKOFF[attempt]))
                if attempt < MAX_RETRIES - 1:
                    log.info("Throttled (429). Waiting %ds before retry.", retry_after)
                    time.sleep(retry_after)
                    continue
                return {
                    "success": False,
                    "error": (
                        f"Sending throttled (HTTP 429). "
                        f"Microsoft asks to wait {retry_after}s. "
                        "Increase the delay between emails in the Send page."
                    ),
                    "code": 429,
                    "retry_after": retry_after,
                }

            # ── 5xx retryable server errors ──────────────────────────
            if resp.status_code in _RETRYABLE_STATUS and attempt < MAX_RETRIES - 1:
                wait = RETRY_BACKOFF[attempt]
                log.warning("Server error %d. Retrying in %ds.", resp.status_code, wait)
                time.sleep(wait)
                continue

            # ── All other errors → report immediately ────────────────
            detail = _parse_graph_error(resp)
            return {
                "success": False,
                "error": f"HTTP {resp.status_code}: {detail}",
                "code": resp.status_code,
            }

        except requests.exceptions.Timeout:
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_BACKOFF[attempt])
                continue
            return {"success": False, "error": "Request timed out after retries.", "code": 0}

        except requests.exceptions.ConnectionError as exc:
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_BACKOFF[attempt])
                continue
            return {"success": False, "error": f"Network error: {exc}", "code": 0}

        except Exception as exc:
            return {"success": False, "error": str(exc), "code": 0}

    return {"success": False, "error": "Maximum retries exceeded.", "code": 0}


# ──────────────────────────────────────────────────────────────────────────────
# Test-email helper (used by Settings page)
# ──────────────────────────────────────────────────────────────────────────────

def send_test_email(access_token: str, to_email: str) -> dict:
    """Send a simple test email to verify Graph API connectivity."""
    return send_email(
        access_token=access_token,
        to_email=to_email,
        to_name="Test",
        subject="✅ Outlook Bulk Mailer — Connection Test",
        body=(
            "<p>This is an automated test email from <b>Outlook Bulk Mailer</b>.</p>"
            "<p>If you received this, your Microsoft Graph API connection is working correctly.</p>"
            "<p style='color:#666;font-size:12px;'>You can ignore or delete this message.</p>"
        ),
        delay=0,
    )
