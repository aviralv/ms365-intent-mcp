#!/usr/bin/env python3
"""Device code flow authentication for ms365-intent-mcp.

Thin wrapper around `ms365_intent_mcp.auth_cli.main`. Prefer the installed
CLI: `ms365-intent-mcp auth`. This script exists for `uv run python
scripts/authenticate.py` muscle memory and for pre-install bootstrap.
"""

import sys

from ms365_intent_mcp.auth_cli import main

if __name__ == "__main__":
    sys.exit(main())
