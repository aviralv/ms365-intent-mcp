#!/usr/bin/env python3
"""Device code flow authentication for ms365-intent-mcp."""

import argparse
import json
import os
import sys
import tempfile
import webbrowser
from pathlib import Path

TOKEN_PATH = Path.home() / ".config" / "ms365-intent-mcp" / "token.json"
CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
DEFAULT_TENANT = "common"
SCOPES = [
    "Calendars.ReadWrite",
    "Mail.Read",
    "Mail.ReadWrite",
    "Tasks.Read",
    "Files.Read",
    "Contacts.Read",
    "User.Read",
    "Chat.ReadWrite",
    "ChatMessage.Send",
    "Channel.ReadBasic.All",
    "Team.ReadBasic.All",
    "Sites.Read.All",
]


def main():
    parser = argparse.ArgumentParser(description="Authenticate for ms365-intent-mcp")
    parser.add_argument(
        "--tenant",
        default=os.environ.get("MS365_INTENT_TENANT_ID", DEFAULT_TENANT),
        help="Azure AD tenant ID (default: %(default)s)",
    )
    args = parser.parse_args()

    try:
        import msal
    except ImportError:
        print("Error: msal not installed. Run: uv pip install ms365-intent-mcp[auth]")
        sys.exit(1)

    authority = f"https://login.microsoftonline.com/{args.tenant}"
    app = msal.PublicClientApplication(CLIENT_ID, authority=authority)
    flow = app.initiate_device_flow(scopes=SCOPES)

    if "user_code" not in flow:
        print(f"Error: {flow.get('error_description', 'Unknown error')}")
        sys.exit(1)

    user_code = flow["user_code"]
    print(f"\nDevice code: {user_code}")
    print("Opening browser...")

    auth_url = f"https://microsoft.com/devicelogin?otc={user_code}"
    webbrowser.open(auth_url)
    print("Waiting for authentication (complete in browser)...")

    result = app.acquire_token_by_device_flow(flow)

    if "error" in result:
        print(f"\nError: {result.get('error_description', 'Authentication failed')}")
        sys.exit(1)

    TOKEN_PATH.parent.mkdir(parents=True, exist_ok=True)
    token_data = {
        "access_token": result["access_token"],
        "refresh_token": result.get("refresh_token"),
        "expires_in": result.get("expires_in", 3600),
        "scope": result.get("scope", ""),
        "token_type": result.get("token_type", "Bearer"),
    }
    # Atomic write: temp file in the same directory + os.replace, so a
    # concurrent server-side refresh can't tear this write.
    fd, tmp_path = tempfile.mkstemp(prefix=".token.", suffix=".tmp", dir=TOKEN_PATH.parent)
    try:
        with os.fdopen(fd, "w") as f:
            json.dump(token_data, f, indent=2)
        os.chmod(tmp_path, 0o600)
        os.replace(tmp_path, TOKEN_PATH)
    except Exception:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise

    print(f"\n✓ Authentication successful!")
    print(f"✓ Token saved to: {TOKEN_PATH}")


if __name__ == "__main__":
    main()
