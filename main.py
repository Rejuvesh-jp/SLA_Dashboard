import pandas as pd
import json
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
    with open(path, "r") as f:
        return json.load(f)

def save_json(path, data):
    with open(path, "w") as f:
        json.dump(data, f, indent=2)



from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse, HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from starlette.templating import Jinja2Templates
import pandas as pd
from typing import Dict, Any

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
async def auth_callback(code: str | None = None, state: str | None = None, error: str | None = None):
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


COLUMNS = {
    "number": "Number",
    "priority": "Priority",
    "state": "State",
    "created": "Created",
    "hours": "Hours Outstanding",
    "sla": "SLA",
    "team": "Team",
    "assignment_group": "Assignment group"
}

# =============================
# Excel reader (NO CACHE)
# =============================
def read_excel() -> pd.DataFrame:
    try:
        # Always read the first sheet (sheet name may be anything).
        # Data source: SharePoint via Microsoft Graph (read-only).
        from graph_sharepoint_excel import fetch_latest_excel_first_sheet_as_dataframe

        df = fetch_latest_excel_first_sheet_as_dataframe()
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
            df[COLUMNS["created"]] = pd.to_datetime(
                df[COLUMNS["created"]], errors="coerce"
            ).dt.strftime("%Y-%m-%d %H:%M:%S")

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
            
            # 2. ELSE IF Email exists in EMAIL_TEAM_MAP
            email = row.get("Email")
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

        return df

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
async def get_tickets() -> Dict[str, Any]:
    df = read_excel()
    records = df.to_dict("records")

    # 🔥 Detect deleted tickets HERE
    detect_deleted_tickets(records)

    return {
        "success": True,
        "count": len(df),
        "data": records
    }


# =============================
# KPI API
# =============================
@app.get("/api/kpis")
async def get_kpis() -> Dict[str, Any]:
    df = read_excel()

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
    df = read_excel()

    

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
    df = read_excel()

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
    df = read_excel()

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
# Dashboard UI
# =============================
@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request):
    # If not signed in, redirect to Microsoft login.
    from graph_auth import hasDelegatedGraphToken

    if not hasDelegatedGraphToken():
        return RedirectResponse("/auth/login")
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

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
