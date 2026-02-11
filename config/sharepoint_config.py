"""
SharePoint / Microsoft Graph configuration.

Note: Values here are non-secret identifiers/paths. Secrets remain in environment variables
used by `graph_auth.py`.
"""

# Fixed configuration (per requirements)
SHAREPOINT_SITE_DOMAIN = "titancompltd-my.sharepoint.com"
SHAREPOINT_SITE_NAME = "personal/rejuveshj_titan_co_in"
# OneDrive folder path (relative to /me/drive/root:)
SHAREPOINT_FOLDER_PATH = "Sla_data"
EXCEL_FILE_NAME = "sla_pending(1).xlsx"

