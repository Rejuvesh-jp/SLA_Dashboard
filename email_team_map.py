"""
Compatibility shim.

The project stores the mapping in `config/email_team_map.py`, but some modules expect:
    from email_team_map import EMAIL_TEAM_MAP
"""

from config.email_team_map import EMAIL_TEAM_MAP  # re-export

