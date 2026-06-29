import sys
try:
    if sys.platform == 'win32':
        import truststore
        truststore.inject_into_ssl()  # Use Windows certificate store (fixes corporate SSL inspection)
except ImportError:
    pass  # truststore not required on Linux

import database as db

import pandas as pd
import json
from typing import Dict, Any, Optional
from datetime import datetime, timedelta
import os
import logging
from email_team_map import EMAIL_TEAM_MAP

# =============================
# Environment (.env) support (startup only)
# =============================
# Load env vars from a local .env file (placed in project root).
# Must run before any Microsoft Graph authentication is triggered.
from dotenv import load_dotenv

load_dotenv()

logger = logging.getLogger(__name__)

SNAPSHOT_FILE = "data_snapshot.json"
CLOSED_FILE = "recently_closed.json"

def load_json(path):
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r") as f:
            return json.load(f)
    except (json.JSONDecodeError, ValueError):
        logger.warning("Corrupted JSON file %s — resetting to empty list.", path)
        return []

def _json_default(obj):
    """Serialize types that standard json cannot handle (e.g. pandas Timestamp, numpy types)."""
    import pandas as pd
    import numpy as np
    if isinstance(obj, (pd.Timestamp,)):
        return obj.isoformat() if not pd.isna(obj) else None
    if isinstance(obj, float) and (obj != obj):  # NaN
        return None
    if isinstance(obj, np.integer):
        return int(obj)
    if isinstance(obj, np.floating):
        return float(obj)
    if isinstance(obj, np.bool_):
        return bool(obj)
    return str(obj)

def save_json(path, data):
    with open(path, "w") as f:
        json.dump(data, f, indent=2, default=_json_default)



from fastapi import FastAPI, HTTPException, Request, UploadFile, File
from fastapi.responses import JSONResponse, HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from starlette.templating import Jinja2Templates
from pydantic import BaseModel
import pandas as pd
import io

app = FastAPI(title="SLA Monitoring API", version="1.0.0")


@app.on_event("startup")
async def _graph_auth_startup_check():
    """
    Minimal, non-blocking Graph auth validation.
    - Logs success/failure without leaking secrets.
    - Never crashes the app.
    """
    # Log whether env vars are present (never log values)
    tenant_set = bool(os.getenv("TENANT_ID", "").strip())
    client_set = bool(os.getenv("CLIENT_ID", "").strip())
    secret_set = bool(os.getenv("CLIENT_SECRET", "").strip())
    logger.info(
        "Graph env vars present: TENANT_ID=%s CLIENT_ID=%s CLIENT_SECRET=%s",
        "yes" if tenant_set else "no",
        "yes" if client_set else "no",
        "yes" if secret_set else "no",
    )

    # NOTE: Delegated auth requires an interactive user sign-in.
    # We intentionally do NOT attempt token acquisition at startup.
    logger.info("Graph delegated auth ready (visit /auth/login to sign in)")

    # Initialise PostgreSQL schema (no-op if DATABASE_URL not set)
    db.init_db()


# =============================
# Delegated Auth Endpoints (OAuth2 Authorization Code Flow)
# =============================
@app.get("/auth/login")
async def auth_login():
    from graph_auth import build_authorization_url

    url = build_authorization_url()
    if not url:
        return HTMLResponse(
            "<h3>Microsoft login is not configured</h3>"
            "<p>Please set TENANT_ID, CLIENT_ID, CLIENT_SECRET, and REDIRECT_URI environment variables.</p>"
            "<p>Then refresh this page.</p>",
            status_code=500,
        )

    return RedirectResponse(url)


@app.get("/auth/callback")
async def auth_callback(code: Optional[str] = None, state: Optional[str] = None, error: Optional[str] = None):
    # Microsoft may return error params instead of code
    if error:
        logger.error("Microsoft Graph authentication failed (error=%s)", error)
        return HTMLResponse(
            f"<h3>Microsoft login failed</h3><p>Error: {error}</p><p><a href='/auth/login'>Try again</a></p>",
            status_code=400,
        )

    if not code or not state:
        return HTMLResponse(
            "<h3>Microsoft login failed</h3><p>Missing authorization code/state.</p><p><a href='/auth/login'>Try again</a></p>",
            status_code=400,
        )

    from graph_auth import handle_auth_callback

    ok = handle_auth_callback(code, state)
    if not ok:
        return HTMLResponse(
            "<h3>Microsoft login failed</h3><p>Token exchange failed. Check server logs for details.</p><p><a href='/auth/login'>Try again</a></p>",
            status_code=400,
        )

    return RedirectResponse("/")


@app.get("/auth/logout")
async def auth_logout():
    from graph_auth import clear_delegated_token

    clear_delegated_token()
    return RedirectResponse("/auth/login")


# =============================
# Debug: OneDrive root folder listing (delegated Graph)
# =============================
@app.get("/debug/onedrive-root-folders")
async def debug_onedrive_root_folders():
    """
    Temporary debug endpoint:
    - Uses delegated token (user must be signed in).
    - Calls Microsoft Graph: GET /me/drive/root/children
    - Returns only folder names (ignores files).
    """
    from graph_auth import getGraphAccessToken

    token = getGraphAccessToken()
    if not token:
        raise HTTPException(status_code=401, detail="Microsoft Graph authentication failed")

    import requests

    url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    try:
        resp = requests.get(url, headers=headers, timeout=20)
        if resp.status_code in (401, 403):
            raise HTTPException(status_code=resp.status_code, detail="Microsoft Graph access denied")
        if resp.status_code >= 400:
            raise HTTPException(status_code=500, detail=f"Microsoft Graph request failed ({resp.status_code})")

        payload = resp.json() or {}
        items = payload.get("value") or []
        folder_names = []
        for it in items:
            if not isinstance(it, dict):
                continue
            if "folder" not in it:
                continue
            name = str(it.get("name") or "").strip()
            if name:
                folder_names.append(name)

        logger.info("OneDrive root folders: %s", folder_names)
        return {"folders": folder_names}
    except HTTPException:
        raise
    except Exception as e:
        logger.exception("OneDrive root folder debug failed (%s)", str(e))
        raise HTTPException(status_code=500, detail="OneDrive root folder debug failed")


# =============================
# Debug: OneDrive root folder listing (API namespace)
# =============================
@app.get("/api/debug/onedrive-root")
async def api_debug_onedrive_root():
    """
    Temporary debug endpoint:
    - Uses delegated token (user must be signed in).
    - Calls Microsoft Graph: GET /me/drive/root/children
    - Returns only folder names (ignores files).
    """
    try:
        from graph_auth import getGraphAccessToken
        import requests

        token = getGraphAccessToken()
        if not token:
            return {"error": "Microsoft Graph authentication failed"}

        url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        resp = requests.get(url, headers=headers, timeout=20)
        if resp.status_code >= 400:
            return {"error": f"Microsoft Graph request failed ({resp.status_code})"}

        payload = resp.json() or {}
        items = payload.get("value") or []
        folders = []
        for it in items:
            if isinstance(it, dict) and "folder" in it:
                name = str(it.get("name") or "").strip()
                if name:
                    folders.append(name)

        logger.info("OneDrive root folders (api): %s", folders)
        return {"folders": folders}
    except Exception as e:
        return {"error": str(e)}

# =============================
# Frontend setup
# =============================
templates = Jinja2Templates(directory="templates")
app.mount("/static", StaticFiles(directory="static"), name="static")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# =============================
# Excel configuration
# =============================
EXCEL_FILE_PATH = r"D:\Onedrive_Reju\OneDrive - Titan Company Limited\Sla_data\sla_pending(1).xlsx"

# =============================
# In-memory uploaded file store
# =============================
# Holds the bytes of the most recently uploaded Excel/CSV file.
# Set by POST /api/upload, consumed by read_excel().
_uploaded_file: Optional[dict] = None  # {"bytes": bytes, "filename": str}


COLUMNS = {
    "number": "Number",
    "priority": "Priority",
    "state": "State",
    "created": "Created",
    "resolved": "Resolved",
    "short_description": "Short description",
    "contact_email": "Contact Email ID",
    "description": "Description text",
    "hours": "Hours Outstanding",
    "sla": "SLA",
    "team": "Team",
    "assignment_group": "Assignment group"
}

# =============================
# Excel reader (NO CACHE)
# =============================
def read_excel() -> pd.DataFrame:
    global _uploaded_file
    try:
        # ── NEW: Use uploaded file if one has been provided ───────────────
        if _uploaded_file is not None:
            fname = _uploaded_file["filename"]
            raw = io.BytesIO(_uploaded_file["bytes"])
            if fname.lower().endswith(".csv"):
                df = pd.read_csv(raw)
            else:
                df = pd.read_excel(raw, sheet_name=0)
            _excel_filename = fname
            df.columns = df.columns.str.strip()
        else:
            # ── LEGACY: SharePoint / Microsoft Graph path ─────────────────
            from graph_sharepoint_excel import fetch_latest_excel_first_sheet_as_dataframe
            df, _excel_filename = fetch_latest_excel_first_sheet_as_dataframe()
            df.columns = df.columns.str.strip()

        df = df.where(pd.notnull(df), None)

        # =============================
        # Data cleaning (backend only)
        # - Drop rows that are completely empty
        # - Strip whitespace from string fields
        # - Reset DataFrame index
        # =============================
        obj_cols = df.select_dtypes(include=["object"]).columns
        for col in obj_cols:
            df[col] = df[col].map(lambda v: v.strip() if isinstance(v, str) else v)
            # Treat whitespace-only strings as empty
            df[col] = df[col].map(lambda v: None if v == "" else v)

        df = df.dropna(how="all").reset_index(drop=True)

        if COLUMNS["created"] in df.columns:
            # Compute Resolution Time (hrs) before converting Created to string
            created_dt = pd.to_datetime(df[COLUMNS["created"]], errors="coerce")
            if COLUMNS["resolved"] in df.columns:
                resolved_dt = pd.to_datetime(df[COLUMNS["resolved"]], errors="coerce")
                resolution_secs = (resolved_dt - created_dt).dt.total_seconds()
                df["Resolution Time (hrs)"] = (resolution_secs / 3600).round(2)
                # Keep None where resolution time is invalid/missing
                df["Resolution Time (hrs)"] = df["Resolution Time (hrs)"].where(
                    pd.notnull(df["Resolution Time (hrs)"]), None
                )
                # Format Resolved column as a readable string
                df[COLUMNS["resolved"]] = resolved_dt.dt.strftime("%Y-%m-%d %H:%M:%S")

            df[COLUMNS["created"]] = created_dt.dt.strftime("%Y-%m-%d %H:%M:%S")

        # =============================
        # Team Resolution Logic (Precedence)
        # 1. Existing Team value (if non-empty)
        # 2. Email mapping (if match found)
        # 3. Default to "Others"
        # =============================
        def resolve_team(row):
            # 1. IF Team column exists AND value is non-null AND non-empty
            team_val = row.get(COLUMNS["team"])
            if team_val is not None and not pd.isna(team_val):
                team_str = str(team_val).strip()
                # Treat empty / "nan" as missing
                if team_str and team_str.lower() != "nan":
                    return team_str
            
            # 2. ELSE IF Contact Email ID (or legacy Email) exists in EMAIL_TEAM_MAP
            email = row.get(COLUMNS["contact_email"]) or row.get("Email")
            # If Email is blank/NaN, keep default "Others"
            if email is not None and not pd.isna(email):
                # Some rows contain multiple email IDs separated by commas/semicolons.
                # Requirement: use ONLY the first email (everything before the first comma).
                email_str = str(email).strip()
                if not email_str:
                    return "Others"
                # Split on comma first (primary requirement), then semicolon as fallback.
                if "," in email_str:
                    email_str = email_str.split(",", 1)[0]
                elif ";" in email_str:
                    email_str = email_str.split(";", 1)[0]

                email_key = email_str.strip().lower()
                if email_key in EMAIL_TEAM_MAP:
                    return EMAIL_TEAM_MAP[email_key]
            
            # 3. ELSE: Assign team as "Others"
            return "Others"

        df[COLUMNS["team"]] = df.apply(resolve_team, axis=1)
        # Ensure Team is always filled (never NaN)
        df[COLUMNS["team"]] = df[COLUMNS["team"]].where(pd.notnull(df[COLUMNS["team"]]), "Others")

        return df, _excel_filename

    except HTTPException:
        raise
    except Exception as e:
        msg = str(e)
        # Graceful auth/access handling for delegated Graph tokens
        if "Microsoft Graph authentication failed" in msg:
            raise HTTPException(status_code=401, detail=msg)
        if "Microsoft Graph access denied" in msg:
            raise HTTPException(status_code=403, detail=msg)
        # Clear "not found" scenarios from SharePoint traversal
        if (
            "SharePoint folder is empty" in msg
            or "OneDrive folder is empty" in msg
            or "Unable to list OneDrive folder" in msg
            or "OneDrive folder not found" in msg
            or "No Excel file found" in msg
            or "Excel file not found in OneDrive folder" in msg
            or "Excel file not found at Graph path" in msg
            or "Excel file not found in SharePoint folder" in msg
        ):
            raise HTTPException(status_code=404, detail=msg)

        raise HTTPException(status_code=500, detail=msg)


# =============================
# Health
# =============================
@app.get("/health")
async def health():
    return "OK"


# =============================
# 🔥 DELETED TICKET DETECTION
# =============================
def detect_deleted_tickets(current_records):
    prev_snapshot = load_json(SNAPSHOT_FILE)
    closed = load_json(CLOSED_FILE)

    prev_map = {
        str(t[COLUMNS["number"]]): t for t in prev_snapshot
        if COLUMNS["number"] in t
    }

    current_numbers = {
        str(t[COLUMNS["number"]]) for t in current_records
        if COLUMNS["number"] in t
    }

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for number, ticket in prev_map.items():
        if number not in current_numbers:
            ticket["closed_at"] = now
            closed.insert(0, ticket)

    save_json(CLOSED_FILE, closed[:100])
    save_json(SNAPSHOT_FILE, current_records)


# =============================
# Tickets API
# =============================
@app.get("/api/tickets")
async def get_tickets():
    import math
    import numpy as np

    df, _excel_filename = read_excel()
    records = df.to_dict("records")

    # Sanitize every value — converts NaN/Inf/numpy types to JSON-safe Python types.
    # Must happen before detect_deleted_tickets (save_json) AND before JSONResponse.
    def _sanitize(val):
        if val is None:
            return None
        import pandas as pd
        # NaT must be checked before Timestamp (NaT is a subtype)
        if val is pd.NaT:
            return None
        if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
            return None
        if isinstance(val, np.integer):
            return int(val)
        if isinstance(val, np.floating):
            return None if (math.isnan(val) or math.isinf(val)) else float(val)
        if isinstance(val, np.bool_):
            return bool(val)
        if isinstance(val, pd.Timestamp):
            return None if pd.isna(val) else val.isoformat()
        return val

    records = [{k: _sanitize(v) for k, v in row.items()} for row in records]

    # Apply manual ticket-state overrides from the database
    overrides = db.get_overrides()  # {ticket_number: {state, resolved_at, comment}}
    for record in records:
        num = str(record.get("Number") or "")
        if num in overrides:
            record["State"] = overrides[num]["state"]
            record["_overridden"] = True
            record["_resolved_at"] = overrides[num]["resolved_at"]
            record["_resolve_comment"] = overrides[num].get("comment", "")

    # 🔥 Detect deleted tickets HERE
    detect_deleted_tickets(records)

    # Save snapshot to DB (background thread so response is not delayed)
    import threading
    threading.Thread(
        target=db.save_snapshot,
        args=(records, _excel_filename),
        daemon=True,
    ).start()

    # Return via JSONResponse to bypass FastAPI/Pydantic serialization entirely.
    return JSONResponse(content={
        "success": True,
        "count": len(records),
        "data": records
    })


# =============================
# KPI API
# =============================
@app.get("/api/kpis")
async def get_kpis() -> Dict[str, Any]:
    df, _ = read_excel()

    if df.empty:
        return {"success": True, "kpis": {}}

    def chart(series):
        counts = series.value_counts(dropna=False).to_dict()
        return {
            "data": {str(k): int(v) for k, v in counts.items()},
            "chart_data": [{"label": str(k), "value": int(v)} for k, v in counts.items()]
        }

    sla_col = COLUMNS["sla"]


    breached = df[sla_col].astype(str).str.contains("breach", case=False, na=False).sum()
    at_risk = df[sla_col].astype(str).str.contains("risk", case=False, na=False).sum()

    return {
        "success": True,
        "kpis": {
            "total_tickets": len(df),
            "by_priority": chart(df[COLUMNS["priority"]]),
            "by_sla": chart(df[sla_col]),
            "by_team": chart(df[COLUMNS["team"]]),
            "by_assignment_group": chart(df[COLUMNS["assignment_group"]]),
            "breached_count": int(breached),
            "at_risk_count": int(at_risk)
        }
    }


# =============================
# Aging Buckets
# =============================
@app.get("/api/aging-buckets")
async def aging_buckets():
    df, _ = read_excel()

    

    if COLUMNS["hours"] not in df.columns:
        return {"success": True, "chart_data": []}

    hours = pd.to_numeric(df[COLUMNS["hours"]], errors="coerce")

    buckets = {
        "0-4 hrs": ((hours >= 0) & (hours < 4)).sum(),
        "4-8 hrs": ((hours >= 4) & (hours < 8)).sum(),
        "8-24 hrs": ((hours >= 8) & (hours < 24)).sum(),
        ">24 hrs": (hours >= 24).sum()
    }

    return {
        "success": True,
        "chart_data": [{"label": k, "value": int(v)} for k, v in buckets.items()]
    }


# =============================
# Top Oldest Tickets
# =============================
@app.get("/api/top-oldest-tickets")
async def top_oldest():
    df, _ = read_excel()

    if COLUMNS["hours"] not in df.columns:
        return {"success": True, "data": []}

    df["__hours"] = pd.to_numeric(df[COLUMNS["hours"]], errors="coerce")
    df = df.sort_values("__hours", ascending=False)

    return {
        "success": True,
        "data": df.head(10).drop(columns="__hours").to_dict("records")
    }


# =============================
# SLA Breach Trend
# =============================
@app.get("/api/created-trend")
async def created_trend():
    df, _ = read_excel()

    if COLUMNS["created"] not in df.columns or COLUMNS["sla"] not in df.columns:
        return {"success": True, "chart_data": []}

    # Filter for breached tickets only
    sla_col = COLUMNS["sla"]
    breached_df = df[df[sla_col].astype(str).str.contains("breach", case=False, na=False)].copy()
    
    if breached_df.empty:
        return {"success": True, "chart_data": []}

    created = pd.to_datetime(breached_df[COLUMNS["created"]], errors="coerce").dt.date
    counts = created.value_counts().sort_index()

    return {
        "success": True,
        "chart_data": [{"date": str(k), "count": int(v)} for k, v in counts.items()]
    }


# =============================
# File Upload API
# =============================
@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    """
    Upload an Excel (.xlsx / .xls) or CSV file to use as the ticket data source.
    The file is stored in-memory for the lifetime of the server process.
    Subsequent calls to /api/tickets (and all other data APIs) will use this file.
    """
    global _uploaded_file

    fname = file.filename or ""
    if not any(fname.lower().endswith(ext) for ext in (".xlsx", ".xls", ".csv")):
        raise HTTPException(status_code=400, detail="Only .xlsx, .xls or .csv files are accepted.")

    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")

    _uploaded_file = {"bytes": contents, "filename": fname}
    logger.info("File uploaded: %s (%d bytes)", fname, len(contents))
    return {"success": True, "filename": fname, "size_bytes": len(contents)}


@app.delete("/api/upload")
async def clear_uploaded_file():
    """Remove the in-memory uploaded file so the app falls back to SharePoint."""
    global _uploaded_file
    _uploaded_file = None
    return {"success": True, "message": "Uploaded file cleared. Falling back to SharePoint data source."}


@app.get("/api/upload/status")
async def upload_status():
    """Return whether an uploaded file is currently active."""
    if _uploaded_file:
        return {"active": True, "filename": _uploaded_file["filename"], "size_bytes": len(_uploaded_file["bytes"])}
    return {"active": False}


# =============================
# Dashboard UI
# =============================
@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request):
    return templates.TemplateResponse("dashboard.html", {"request": request})


# =============================
# API Info
# =============================
@app.get("/api")
async def api_info():
    return {
        "message": "SLA Monitoring API",
        "endpoints": [
            "/health",
            "/api/tickets",
            "/api/kpis",
            "/api/aging-buckets",
            "/api/top-oldest-tickets",
            "/api/created-trend"
        ]
    }


# =============================
# Run
# =============================
# =============================
# Recently Closed Tickets API
# =============================
@app.get("/api/recently-closed")
async def recently_closed():
    try:
        if not os.path.exists(CLOSED_FILE):
            return {"success": True, "data": []}

        with open(CLOSED_FILE, "r") as f:
            raw = json.load(f)

        # Normalize structure
        if isinstance(raw, dict):
            data = list(raw.values())
        elif isinstance(raw, list):
            data = raw
        else:
            data = []

        # Filter ONLY last 24 hours at response time (do not modify the stored JSON)
        now = datetime.now()
        cutoff = now - timedelta(hours=24)

        def _parse_closed_at(item):
            # Support common field names; file may keep older/legacy keys
            ts = item.get("closed_at") or item.get("closedAt") or item.get("closed_at_ts") or ""
            if not ts:
                return None
            s = str(ts).strip()
            try:
                # Handle ISO strings (optionally ending with Z)
                if s.endswith("Z"):
                    s = s[:-1]
                return datetime.fromisoformat(s)
            except Exception:
                pass
            try:
                # Handles detect_deleted_tickets format: "YYYY-mm-dd HH:MM:SS"
                return datetime.strptime(s, "%Y-%m-%d %H:%M:%S")
            except Exception:
                return None

        recent = []
        for item in data:
            if not isinstance(item, dict):
                continue
            dt = _parse_closed_at(item)
            if dt is None:
                continue
            if dt >= cutoff:
                recent.append((dt, item))

        # Sort by closed time (most recent first)
        recent.sort(key=lambda t: t[0], reverse=True)
        data = [item for _, item in recent]

        # 🔒 SANITIZE values (CRITICAL FIX)
        clean_data = []
        for item in data:
            clean_item = {}
            for k, v in item.items():
                if v is None:
                    clean_item[k] = ""
                elif isinstance(v, float):
                    if pd.isna(v) or v == float("inf") or v == float("-inf"):
                        clean_item[k] = ""
                    else:
                        clean_item[k] = v
                else:
                    clean_item[k] = v
            clean_data.append(clean_item)

        # Keep output ordering consistent (closed_at desc) after sanitize
        clean_data = sorted(clean_data, key=lambda x: x.get("closed_at", ""), reverse=True)

        return {
            "success": True,
            "data": clean_data[:10]
        }

    except Exception as e:
        print("RECENTLY CLOSED ERROR:", e)
        raise HTTPException(status_code=500, detail=str(e))

# =============================
# Ticket History (Snapshots)
# =============================
@app.get("/api/history")
async def get_history():
    """List all snapshots, most recent first."""
    snapshots = db.list_snapshots()
    return JSONResponse(content={"success": True, "data": snapshots})


@app.get("/api/history/{snapshot_id}")
async def get_history_snapshot(snapshot_id: int):
    """Return all ticket records for a given snapshot, plus metadata."""
    result = db.get_snapshot_tickets(snapshot_id)
    if result is None:
        raise HTTPException(status_code=404, detail="Snapshot not found")
    tickets = result.pop("tickets", [])
    return JSONResponse(content={
        "success": True,
        "count": len(tickets),
        "data": tickets,
        **result  # includes id, fetched_at, file_name, ticket_count
    })


# =============================
# Ticket State Override (Resolve)
# =============================
class ResolveRequest(BaseModel):
    comment: str = ""

@app.post("/api/tickets/{ticket_number}/resolve")
async def resolve_ticket(ticket_number: str, body: ResolveRequest = ResolveRequest()):
    """Mark a ticket as Resolved via a manual override stored in DB."""
    ok = db.set_override(ticket_number, state="Resolved", comment=body.comment)
    if not ok:
        raise HTTPException(status_code=500, detail="Database unavailable or write failed")
    return {"success": True, "ticket_number": ticket_number, "state": "Resolved"}


@app.delete("/api/tickets/{ticket_number}/resolve")
async def unresolve_ticket(ticket_number: str):
    """Remove a manual Resolved override (reverts to Excel state on next fetch)."""
    ok = db.remove_override(ticket_number)
    if not ok:
        raise HTTPException(status_code=500, detail="Database unavailable or write failed")
    return {"success": True, "ticket_number": ticket_number}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8001)
    
