"""
Microsoft Graph Authentication Utility (Delegated OAuth2 Authorization Code Flow)

Why:
- Tenant admin granted *delegated* permissions only (e.g., Files.Read.All, Sites.Read.All),
  so client-credentials (app-only) tokens won't work.

What this provides:
- build_authorization_url(): redirect user to Microsoft login
- handle_auth_callback(): exchange authorization code for delegated access token
- getGraphAccessToken(): return a valid delegated access token (refresh when expired)

Storage:
- In-memory only (single-user / single-process). No secrets are logged.
"""

from __future__ import annotations

import base64
import hashlib
import logging
import os
import secrets
import threading
import time
from typing import Dict, Optional, Tuple
from urllib.parse import urlencode

import requests

logger = logging.getLogger(__name__)

# Refresh a bit early to avoid edge-of-expiry failures due to clock skew/network latency.
_EXPIRY_SKEW_SECONDS = 60

_LOCK = threading.Lock()
_ENV_LOGGED = False

# PKCE verifier store keyed by state
_PKCE_BY_STATE: Dict[str, str] = {}

# Token cache (delegated)
_ACCESS_TOKEN: Optional[str] = None
_REFRESH_TOKEN: Optional[str] = None
_EXPIRES_AT_EPOCH: float = 0.0


def _log_env_presence_once() -> None:
    global _ENV_LOGGED
    if _ENV_LOGGED:
        return
    _ENV_LOGGED = True

    tenant_set = bool(os.getenv("TENANT_ID", "").strip())
    client_set = bool(os.getenv("CLIENT_ID", "").strip())
    secret_set = bool(os.getenv("CLIENT_SECRET", "").strip())
    redirect_set = bool(os.getenv("REDIRECT_URI", "").strip())

    # Use warning level so it's visible even if log level is default WARNING.
    logger.warning(
        "Graph env vars present: TENANT_ID=%s CLIENT_ID=%s CLIENT_SECRET=%s REDIRECT_URI=%s",
        "yes" if tenant_set else "no",
        "yes" if client_set else "no",
        "yes" if secret_set else "no",
        "yes" if redirect_set else "no",
    )


def _env() -> Tuple[str, str, str, str]:
    tenant_id = os.getenv("TENANT_ID", "").strip()
    client_id = os.getenv("CLIENT_ID", "").strip()
    client_secret = os.getenv("CLIENT_SECRET", "").strip()
    redirect_uri = os.getenv("REDIRECT_URI", "").strip() or "http://localhost:8000/auth/callback"
    return tenant_id, client_id, client_secret, redirect_uri


def _authority_base(tenant_id: str) -> str:
    return f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0"


def _scopes() -> str:
    # Delegated scopes (v2 requires full resource scopes)
    # Include offline_access for refresh tokens.
    return " ".join(
        [
            "https://graph.microsoft.com/Files.Read.All",
            "https://graph.microsoft.com/Sites.Read.All",
            "offline_access",
            "openid",
            "profile",
        ]
    )


def _pkce_pair() -> Tuple[str, str]:
    verifier = secrets.token_urlsafe(48)
    digest = hashlib.sha256(verifier.encode("utf-8")).digest()
    challenge = base64.urlsafe_b64encode(digest).decode("utf-8").rstrip("=")
    return verifier, challenge


def hasDelegatedGraphToken() -> bool:
    return bool(_ACCESS_TOKEN) and (time.time() < _EXPIRES_AT_EPOCH)


def build_authorization_url() -> str:
    """
    Build Microsoft Identity authorize URL for delegated auth-code flow (with PKCE).
    """
    _log_env_presence_once()
    tenant_id, client_id, _, redirect_uri = _env()
    if not tenant_id or not client_id:
        logger.error("Microsoft Graph authentication failed")
        return ""

    verifier, challenge = _pkce_pair()
    state = secrets.token_urlsafe(24)
    with _LOCK:
        _PKCE_BY_STATE[state] = verifier

    params = {
        "client_id": client_id,
        "response_type": "code",
        "redirect_uri": redirect_uri,
        "response_mode": "query",
        "scope": _scopes(),
        "state": state,
        "code_challenge": challenge,
        "code_challenge_method": "S256",
    }
    return f"{_authority_base(tenant_id)}/authorize?{urlencode(params)}"


def _log_token_error(resp: requests.Response, *, context: str) -> None:
    error = ""
    error_description = ""
    error_codes = None
    trace_id = ""
    correlation_id = ""
    ts = ""
    try:
        body = resp.json() or {}
        error = str(body.get("error") or "")
        error_description = str(body.get("error_description") or "")
        error_codes = body.get("error_codes")
        trace_id = str(body.get("trace_id") or "")
        correlation_id = str(body.get("correlation_id") or "")
        ts = str(body.get("timestamp") or "")
    except Exception:
        pass

    logger.error(
        "%s (status=%s error=%s error_description=%s error_codes=%s trace_id=%s correlation_id=%s timestamp=%s)",
        context,
        resp.status_code,
        error,
        error_description,
        error_codes,
        trace_id,
        correlation_id,
        ts,
    )


def handle_auth_callback(code: str, state: str) -> bool:
    """
    Exchange authorization code for delegated access token.
    Returns True on success, False on failure.
    """
    global _ACCESS_TOKEN, _REFRESH_TOKEN, _EXPIRES_AT_EPOCH
    _log_env_presence_once()
    tenant_id, client_id, client_secret, redirect_uri = _env()
    if not tenant_id or not client_id or not client_secret or not redirect_uri:
        logger.error("Microsoft Graph authentication failed")
        return False

    with _LOCK:
        verifier = _PKCE_BY_STATE.pop(state, "")

    if not verifier:
        logger.error("Microsoft Graph authentication failed")
        return False

    token_url = f"{_authority_base(tenant_id)}/token"
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "authorization_code",
        "code": code,
        "redirect_uri": redirect_uri,
        "code_verifier": verifier,
        "scope": _scopes(),
    }

    try:
        resp = requests.post(token_url, data=data, timeout=20)
        if resp.status_code != 200:
            _log_token_error(resp, context="Microsoft Graph authentication failed")
            return False

        payload = resp.json() or {}
        access_token = payload.get("access_token")
        refresh_token = payload.get("refresh_token")
        expires_in = payload.get("expires_in")

        if not access_token or not isinstance(access_token, str):
            logger.error("Microsoft Graph authentication failed")
            return False

        try:
            expires_in_seconds = int(expires_in) if expires_in is not None else 0
        except Exception:
            expires_in_seconds = 0
        if expires_in_seconds <= 0:
            expires_in_seconds = 45 * 60

        with _LOCK:
            _ACCESS_TOKEN = access_token
            _REFRESH_TOKEN = str(refresh_token) if refresh_token else _REFRESH_TOKEN
            _EXPIRES_AT_EPOCH = time.time() + max(0, expires_in_seconds - _EXPIRY_SKEW_SECONDS)

        logger.info("Microsoft Graph authentication succeeded")
        return True
    except Exception as e:
        logger.exception("Microsoft Graph authentication failed (%s)", str(e))
        return False


def _refresh_access_token() -> bool:
    global _ACCESS_TOKEN, _REFRESH_TOKEN, _EXPIRES_AT_EPOCH
    tenant_id, client_id, client_secret, redirect_uri = _env()
    if not tenant_id or not client_id or not client_secret or not redirect_uri:
        logger.error("Microsoft Graph authentication failed")
        return False

    with _LOCK:
        refresh_token = _REFRESH_TOKEN

    if not refresh_token:
        return False

    token_url = f"{_authority_base(tenant_id)}/token"
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "redirect_uri": redirect_uri,
        "scope": _scopes(),
    }

    try:
        resp = requests.post(token_url, data=data, timeout=20)
        if resp.status_code != 200:
            _log_token_error(resp, context="Microsoft Graph authentication failed")
            return False

        payload = resp.json() or {}
        access_token = payload.get("access_token")
        new_refresh_token = payload.get("refresh_token")
        expires_in = payload.get("expires_in")
        if not access_token or not isinstance(access_token, str):
            return False

        try:
            expires_in_seconds = int(expires_in) if expires_in is not None else 0
        except Exception:
            expires_in_seconds = 0
        if expires_in_seconds <= 0:
            expires_in_seconds = 45 * 60

        with _LOCK:
            _ACCESS_TOKEN = access_token
            if new_refresh_token:
                _REFRESH_TOKEN = str(new_refresh_token)
            _EXPIRES_AT_EPOCH = time.time() + max(0, expires_in_seconds - _EXPIRY_SKEW_SECONDS)
        return True
    except Exception as e:
        logger.exception("Microsoft Graph authentication failed (%s)", str(e))
        return False


def getGraphAccessToken() -> str:
    """
    Return a valid *delegated* Microsoft Graph access token (string).

    - Requires user to sign in via /auth/login once.
    - Refreshes token using refresh_token when expired.
    - On failure, logs "Microsoft Graph authentication failed" and returns "".
    """
    _log_env_presence_once()

    with _LOCK:
        token = _ACCESS_TOKEN
        expires_at = _EXPIRES_AT_EPOCH

    if token and time.time() < expires_at:
        return token

    # Try refresh if possible
    if _refresh_access_token():
        with _LOCK:
            return _ACCESS_TOKEN or ""

    logger.error("Microsoft Graph authentication failed")
    return ""


def clear_delegated_token() -> None:
    global _ACCESS_TOKEN, _REFRESH_TOKEN, _EXPIRES_AT_EPOCH
    with _LOCK:
        _ACCESS_TOKEN = None
        _REFRESH_TOKEN = None
        _EXPIRES_AT_EPOCH = 0.0


# Emit env presence once at import time (best-effort)
try:
    _log_env_presence_once()
except Exception:
    pass

