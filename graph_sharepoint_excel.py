"""
OneDrive Excel Reader via Microsoft Graph (read-only, delegated)

Flow used by the backend Excel loader:
1) List files in configured OneDrive folder
4) Find a specific Excel file by exact name (EXCEL_FILE_NAME)
5) Download file bytes
6) Load into pandas with pd.read_excel(sheet_name=0)

Auth:
Uses delegated access token via graph_auth.getGraphAccessToken()
"""

from __future__ import annotations

import logging
from datetime import datetime, timezone
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote

import pandas as pd
import requests

from graph_auth import getGraphAccessToken
from config.sharepoint_config import (
    EXCEL_FILE_NAME,
    SHAREPOINT_FOLDER_PATH,
)

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _parse_graph_dt(dt_str: str) -> datetime:
    """
    Parse Graph datetime like '2026-01-01T10:11:12Z' or with offset.
    """
    s = (dt_str or "").strip()
    if not s:
        return datetime.fromtimestamp(0, tz=timezone.utc)
    # Normalize trailing Z
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return datetime.fromtimestamp(0, tz=timezone.utc)


def _graph_get(
    path: str, *, params: Optional[Dict[str, Any]] = None, context: str = "Microsoft Graph request"
) -> Dict[str, Any]:
    token = getGraphAccessToken()
    if not token:
        raise RuntimeError("Microsoft Graph authentication failed")

    url = path if path.startswith("http") else f"{GRAPH_BASE}{path}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }
    resp = requests.get(url, headers=headers, params=params, timeout=20)
    if resp.status_code >= 400:
        # Keep errors clear but avoid leaking secrets/response bodies.
        if resp.status_code in (401, 403):
            raise RuntimeError("Microsoft Graph access denied")
        raise RuntimeError(f"{context} failed ({resp.status_code})")
    return resp.json() or {}


def _graph_download_bytes(path: str, *, context: str = "Microsoft Graph download") -> bytes:
    """
    Download raw bytes from a Graph endpoint (e.g., /content). Follows redirects.
    Reads response as a stream.
    """
    token = getGraphAccessToken()
    if not token:
        raise RuntimeError("Microsoft Graph authentication failed")

    url = path if path.startswith("http") else f"{GRAPH_BASE}{path}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "*/*",
    }
    resp = requests.get(url, headers=headers, timeout=60, allow_redirects=True, stream=True)
    if resp.status_code >= 400:
        if resp.status_code in (401, 403):
            raise RuntimeError("Microsoft Graph access denied")
        raise RuntimeError(f"{context} failed ({resp.status_code})")
    chunks: List[bytes] = []
    for chunk in resp.iter_content(chunk_size=1024 * 256):
        if chunk:
            chunks.append(chunk)
    return b"".join(chunks)


def _list_onedrive_folder_children(folder_path: str) -> List[Dict[str, Any]]:
    """
    List files in the signed-in user's OneDrive folder:
    GET /me/drive/root:/<folder_path>:/children
    """
    encoded_path = quote(folder_path)
    graph_path = f"/me/drive/root:/{folder_path}:/children"
    try:
        payload = _graph_get(
            f"/me/drive/root:/{encoded_path}:/children",
            context="Unable to list OneDrive folder",
        )
    except Exception as e:
        # Clear log for folder missing / traversal issues
        logger.error("Unable to list OneDrive folder: %s", folder_path)
        msg = str(e)
        if "failed (404)" in msg:
            raise RuntimeError(f"OneDrive folder not found at Graph path: {graph_path}") from None
        raise

    items = payload.get("value") or []
    if not isinstance(items, list):
        return []
    return [i for i in items if isinstance(i, dict)]


def _pick_latest_excel_file(items: List[Dict[str, Any]]) -> Tuple[str, str, str]:
    """
    Returns (file_id, file_name, last_modified_str)
    """
    if not items:
        raise RuntimeError("OneDrive folder is empty")

    excel_items: List[Dict[str, Any]] = []
    for it in items:
        name = str(it.get("name") or "")
        if not name:
            continue
        lower = name.lower()
        # Requirement: filter only .xlsx files
        if not lower.endswith(".xlsx"):
            continue
        # Ensure it's a file (not a folder)
        if "file" not in it:
            continue
        excel_items.append(it)

    if not excel_items:
        raise RuntimeError("No Excel file found in OneDrive folder")

    excel_items.sort(
        key=lambda x: _parse_graph_dt(str(x.get("lastModifiedDateTime") or "")),
        reverse=True,
    )
    latest = excel_items[0]
    file_id = str(latest.get("id") or "")
    file_name = str(latest.get("name") or "")
    last_modified = str(latest.get("lastModifiedDateTime") or "")
    if not file_id:
        raise RuntimeError("Unable to resolve latest Excel file id")

    logger.info("OneDrive latest file: %s | lastModifiedDateTime=%s", file_name, last_modified)
    return file_id, file_name, last_modified


def _find_named_excel_file(items: List[Dict[str, Any]], file_name: str) -> Tuple[str, str]:
    """
    Find an Excel file with name exactly matching `file_name` in the folder listing.
    Returns (file_id, last_modified_str).
    """
    if not items:
        raise RuntimeError("OneDrive folder is empty")

    for it in items:
        name = str(it.get("name") or "")
        if name != file_name:
            continue
        # Ensure it's a file (not a folder)
        if "file" not in it:
            continue
        file_id = str(it.get("id") or "")
        last_modified = str(it.get("lastModifiedDateTime") or "")
        if not file_id:
            raise RuntimeError("Unable to resolve Excel file id")

        logger.info("OneDrive selected file: %s | lastModifiedDateTime=%s", name, last_modified)
        return file_id, last_modified

    raise RuntimeError(f"Excel file not found in OneDrive folder: {file_name}")


def fetch_latest_excel_first_sheet_as_dataframe():
    """
    Fetch the most recently modified Excel file from the configured OneDrive folder,
    download it, and load the first worksheet (index 0) into a DataFrame.
    Returns (DataFrame, filename) tuple.
    """
    items = _list_onedrive_folder_children(SHAREPOINT_FOLDER_PATH)
    _, latest_name, _ = _pick_latest_excel_file(items)

    encoded_full_path = quote(f"{SHAREPOINT_FOLDER_PATH}/{latest_name}")
    content = _graph_download_bytes(
        f"/me/drive/root:/{encoded_full_path}:/content",
        context="Unable to download Excel file",
    )

    bio = BytesIO(content)
    engine = "openpyxl" if latest_name.lower().endswith(".xlsx") else None
    df = pd.read_excel(bio, sheet_name=0, engine=engine)
    if df is None:
        return pd.DataFrame(), latest_name
    return df, latest_name


def fetch_configured_excel_first_sheet_as_dataframe() -> pd.DataFrame:
    """
    Fetch a specific Excel file (EXCEL_FILE_NAME) from the configured SharePoint folder,
    download its content, and load the first worksheet (index 0) into a DataFrame.

    - Sheet name can be anything (only first sheet is read)
    - If the sheet is empty -> returns empty DataFrame
    """
    # Per requirement: use /me/drive/root paths (personal OneDrive), no site-id, no SharePoint APIs.
    items = _list_onedrive_folder_children(SHAREPOINT_FOLDER_PATH)
    _find_named_excel_file(items, EXCEL_FILE_NAME)  # validates presence & logs lastModified

    graph_path = f"/me/drive/root:/{SHAREPOINT_FOLDER_PATH}/{EXCEL_FILE_NAME}:/content"
    encoded_full_path = quote(f"{SHAREPOINT_FOLDER_PATH}/{EXCEL_FILE_NAME}")
    try:
        content = _graph_download_bytes(
            f"/me/drive/root:/{encoded_full_path}:/content",
            context="Unable to download Excel file",
        )
    except Exception as e:
        msg = str(e)
        if "failed (404)" in msg:
            raise RuntimeError(f"Excel file not found at Graph path: {graph_path}") from None
        raise

    bio = BytesIO(content)
    # Always read first sheet; sheet name may be anything
    engine = "openpyxl" if EXCEL_FILE_NAME.lower().endswith(".xlsx") else None
    df = pd.read_excel(bio, sheet_name=0, engine=engine)
    if df is None:
        return pd.DataFrame()
    return df

