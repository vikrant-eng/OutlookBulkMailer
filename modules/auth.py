"""
Authentication module for Outlook Bulk Mailer.
Uses MSAL (Microsoft Authentication Library) for secure OAuth2 login.

Security model:
  - Uses the Device Code Flow (RFC 8628) — no redirect server needed.
  - Tokens are cached via MSAL's SerializableTokenCache — passwords are NEVER stored.
  - Access tokens are never written to logs or settings files.
  - Client secrets are not used (public-client app).

References:
  - https://learn.microsoft.com/en-us/entra/identity-platform/scenario-desktop-app-registration
  - https://learn.microsoft.com/en-us/azure/active-directory/develop/msal-client-applications
"""

import os
import json
import logging
import msal
import requests
from typing import Optional

log = logging.getLogger(__name__)

# ──────────────────────────────────────────────────────────────────────────────
# Constants  (CLIENT_ID is loaded from settings at runtime — never hard-coded)
# ──────────────────────────────────────────────────────────────────────────────
_PLACEHOLDER_ID = "YOUR_AZURE_APP_CLIENT_ID"
CLIENT_ID: str = _PLACEHOLDER_ID

# "common" accepts both personal MSA and work/school AAD accounts.
AUTHORITY = "https://login.microsoftonline.com/common"

# Minimum required Graph scopes.  offline_access enables refresh tokens.
SCOPES = ["Mail.Send", "User.Read", "offline_access"]

GRAPH_ME_URL = "https://graph.microsoft.com/v1.0/me"

_DATA_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data"
)
_TOKEN_CACHE_PATH = os.path.join(_DATA_DIR, "token_cache.bin")
_SETTINGS_PATH    = os.path.join(_DATA_DIR, "settings.json")

# ──────────────────────────────────────────────────────────────────────────────
# Token cache helpers
# ──────────────────────────────────────────────────────────────────────────────

def _build_token_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if os.path.exists(_TOKEN_CACHE_PATH):
        try:
            with open(_TOKEN_CACHE_PATH, "r", encoding="utf-8") as f:
                cache.deserialize(f.read())
        except Exception as exc:
            log.warning("Could not load token cache: %s", exc)
    return cache


def _save_token_cache(cache: msal.SerializableTokenCache) -> None:
    os.makedirs(_DATA_DIR, exist_ok=True)
    if cache.has_state_changed:
        try:
            with open(_TOKEN_CACHE_PATH, "w", encoding="utf-8") as f:
                f.write(cache.serialize())
        except Exception as exc:
            log.warning("Could not save token cache: %s", exc)


def _build_app(cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
    """Build a PublicClientApplication (no secret — never embed a client secret)."""
    if CLIENT_ID == _PLACEHOLDER_ID or not CLIENT_ID:
        raise ConfigError(
            "Azure App Client ID is not configured.\n\n"
            "Open Settings, paste your Application (client) ID, and click Save."
        )
    return msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
    )


# ──────────────────────────────────────────────────────────────────────────────
# Custom exceptions
# ──────────────────────────────────────────────────────────────────────────────

class ConfigError(RuntimeError):
    """Client ID is missing or is a placeholder."""

class ConsentError(RuntimeError):
    """User cancelled or denied consent."""

class TokenExpiredError(RuntimeError):
    """Cached token could not be silently refreshed — re-login needed."""

class ThrottleError(RuntimeError):
    """Graph API returned 429; retry_after is set."""
    def __init__(self, msg: str, retry_after: int = 60):
        super().__init__(msg)
        self.retry_after = retry_after


# ──────────────────────────────────────────────────────────────────────────────
# Public API
# ──────────────────────────────────────────────────────────────────────────────

def get_access_token() -> Optional[str]:
    """
    Try to return a valid access token silently from the cache.

    Returns:
        Access token string, or None if interactive login is required.
    Raises:
        ConfigError  — if CLIENT_ID has not been configured.
    """
    cache = _build_token_cache()
    app = _build_app(cache)          # raises ConfigError if not set
    accounts = app.get_accounts()
    if not accounts:
        return None

    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    _save_token_cache(cache)

    if result and "access_token" in result:
        return result["access_token"]

    # Silent refresh failed (token revoked, password changed, etc.)
    err = (result or {}).get("error", "")
    if err in ("invalid_grant", "interaction_required"):
        raise TokenExpiredError(
            "Your session has expired. Please sign in again."
        )
    return None


def login_interactive() -> dict:
    """
    Initiate the Device Code Flow.

    Returns a dict with keys:
        flow   — the MSAL device flow dict (contains user_code, verification_uri)
        app    — PublicClientApplication instance
        cache  — SerializableTokenCache instance

    Raises:
        ConfigError  — CLIENT_ID not set
        RuntimeError — MSAL could not start the flow
    """
    cache = _build_token_cache()
    app = _build_app(cache)           # raises ConfigError if not configured

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        err_desc = flow.get("error_description", flow.get("error", "Unknown error"))
        raise RuntimeError(
            f"Could not start sign-in flow.\n\n"
            f"Details: {err_desc}\n\n"
            "Common causes:\n"
            "• Invalid Client ID\n"
            "• App registration is in a restricted tenant\n"
            "• Network connectivity issue"
        )
    return {"flow": flow, "app": app, "cache": cache}


def complete_device_flow(app, flow, cache) -> dict:
    """
    Block until the user completes sign-in, then return the token result.

    Call this from a background QThread so the UI stays responsive.

    Raises:
        ConsentError  — user cancelled or denied consent
        RuntimeError  — any other auth failure
    """
    result = app.acquire_token_by_device_flow(flow)
    _save_token_cache(cache)

    if "access_token" not in result:
        err   = result.get("error", "")
        desc  = result.get("error_description", "Unknown error")

        if err in ("authorization_declined", "access_denied"):
            raise ConsentError(
                "Sign-in was cancelled or consent was denied.\n\n"
                "To grant permission, sign in again and click 'Accept' "
                "when Microsoft asks for Mail.Send and User.Read access."
            )
        if err == "code_expired":
            raise RuntimeError(
                "The sign-in code expired before you completed sign-in.\n"
                "Please click 'Sign in with Microsoft' to get a new code."
            )
        raise RuntimeError(f"Authentication failed: {desc}")

    return result


def get_user_info(access_token: str) -> dict:
    """
    Call GET /v1.0/me to verify the token and return the user's profile.

    Returns:
        {displayName, userPrincipalName, mail, ...}
    Raises:
        requests.HTTPError on non-2xx response
        TokenExpiredError  on 401
        RuntimeError       on 403
    """
    # NOTE: access_token is only used in an Authorization header — never logged.
    headers = {"Authorization": f"Bearer {access_token}"}
    resp = requests.get(GRAPH_ME_URL, headers=headers, timeout=15)

    if resp.status_code == 401:
        raise TokenExpiredError("Access token is invalid or expired. Please sign in again.")
    if resp.status_code == 403:
        raise RuntimeError(
            "Permission denied (403).\n\n"
            "Ensure your app registration has the 'User.Read' permission "
            "and that you consented during sign-in."
        )
    resp.raise_for_status()
    return resp.json()


def verify_connection(access_token: str) -> dict:
    """
    Lightweight connectivity + auth check.

    Returns:
        {"ok": True,  "display_name": ..., "email": ...}
        {"ok": False, "error": "..."}
    """
    try:
        info = get_user_info(access_token)
        return {
            "ok": True,
            "display_name": info.get("displayName", ""),
            "email": info.get("userPrincipalName", info.get("mail", "")),
        }
    except TokenExpiredError as e:
        return {"ok": False, "error": str(e), "expired": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def logout() -> None:
    """Delete the local token cache, effectively signing the user out."""
    if os.path.exists(_TOKEN_CACHE_PATH):
        try:
            os.remove(_TOKEN_CACHE_PATH)
        except Exception as exc:
            log.warning("Could not delete token cache: %s", exc)


def is_logged_in() -> bool:
    """Return True if there is at least one cached account."""
    try:
        cache = _build_token_cache()
        app   = _build_app(cache)
        return len(app.get_accounts()) > 0
    except ConfigError:
        return False


def get_client_id() -> str:
    return CLIENT_ID


def set_client_id(new_id: str) -> None:
    """
    Update the in-memory CLIENT_ID and persist it to settings.json.
    The value is a UUID-format string — never a secret.
    """
    global CLIENT_ID
    CLIENT_ID = new_id.strip()
    _persist_setting("client_id", CLIENT_ID)


def load_settings() -> None:
    """Load persisted settings on startup (called once from MainWindow)."""
    global CLIENT_ID
    if not os.path.exists(_SETTINGS_PATH):
        return
    try:
        with open(_SETTINGS_PATH, "r", encoding="utf-8") as f:
            settings = json.load(f)
        if "client_id" in settings and settings["client_id"]:
            CLIENT_ID = settings["client_id"]
    except Exception as exc:
        log.warning("Could not load settings: %s", exc)


def _persist_setting(key: str, value) -> None:
    os.makedirs(_DATA_DIR, exist_ok=True)
    try:
        settings: dict = {}
        if os.path.exists(_SETTINGS_PATH):
            with open(_SETTINGS_PATH, "r", encoding="utf-8") as f:
                settings = json.load(f)
        settings[key] = value
        with open(_SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except Exception as exc:
        log.warning("Could not save setting '%s': %s", key, exc)


def is_valid_client_id(cid: str) -> bool:
    """Validate UUID-format client ID (8-4-4-4-12 hex groups)."""
    import re
    return bool(re.fullmatch(
        r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}",
        cid.strip()
    ))
