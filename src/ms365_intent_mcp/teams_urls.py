"""Teams browser hosts shared by link parsers and invite formatting.

These are URL-recognition hosts, not authenticated API destinations.
"""

import re
from urllib.parse import urlparse

TEAMS_HOSTS = ("teams.microsoft.com", "teams.cloud.microsoft")
TEAMS_URL_PREFIX = r"https?://(?:" + "|".join(re.escape(host) for host in TEAMS_HOSTS) + r")"


def is_teams_url(url: str) -> bool:
    """Recognize an HTTP(S) URL on either exact Teams browser host."""
    parsed = urlparse(url)
    return parsed.scheme in {"http", "https"} and parsed.hostname in TEAMS_HOSTS
