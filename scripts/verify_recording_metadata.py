#!/usr/bin/env python3
"""End-to-end verification for issue #8.

Drives the same two paths a user would exercise:

  1. compose_meeting(identifier)                     → rendered markdown with
                                                       Recording block (vroom_url,
                                                       drive_id, etc.) if the
                                                       meeting has a recording.
  2. compose_resolve(url=<chat_thread URL>)          → chat-thread markdown with
                                                       call entries carrying the
                                                       same drive metadata as
                                                       sub-bullets.

Success = the Recording block shows vroom_url; and the chat-thread call entry
shows the same fields under the [recording] link.

Usage:
  uv run python scripts/verify_recording_metadata.py "<subject substring>"
  uv run python scripts/verify_recording_metadata.py "Refinement"
"""

from __future__ import annotations

import asyncio
import re
import sys
import urllib.parse
from datetime import datetime, timedelta, timezone

from ms365_intent_mcp.auth import TokenManager
from ms365_intent_mcp.composers.meeting import compose_meeting
from ms365_intent_mcp.composers.resolve import compose_resolve
from ms365_intent_mcp.config import Config
from ms365_intent_mcp.graph import GraphClient
from ms365_intent_mcp.permissions import PermissionRegistry


def _line(char: str = "─", n: int = 72) -> None:
    print(char * n)


async def _find_target_thread(client: GraphClient, subject_substring: str) -> tuple[str, str] | None:
    """Scan past 14 days for a meeting matching subject; return (thread_id, subject)
    of the first one with recording events on its chat."""
    now = datetime.now(timezone.utc)
    start = (now - timedelta(days=14)).strftime("%Y-%m-%dT%H:%M:%SZ")
    end = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    result = await client.get("/me/calendarView", params={
        "startDateTime": start, "endDateTime": end,
        "$top": "100", "$orderby": "start/dateTime desc",
    })
    events = (result or {}).get("value", [])
    lower = subject_substring.lower()
    for e in events:
        if lower not in (e.get("subject") or "").lower():
            continue
        join_url = urllib.parse.unquote((e.get("onlineMeeting") or {}).get("joinUrl", ""))
        m = re.search(r"meetup-join/(19:[^/]+@thread\.v2)", join_url)
        if not m:
            continue
        thread_id = m.group(1)
        try:
            msgs = await client.get(f"/chats/{thread_id}/messages", params={"$top": "50"})
        except Exception:
            continue
        messages = (msgs or {}).get("value", [])
        if any("callRecording" in ((m.get("eventDetail") or {}).get("@odata.type", "")) for m in messages):
            return thread_id, e.get("subject") or ""
    return None


async def run(subject: str) -> None:
    config = Config()
    auth = TokenManager(config)
    auth.ensure_authenticated()

    async with GraphClient(config.graph_base_url, auth.get_access_token) as client:
        permissions = PermissionRegistry.from_token_provider(auth.peek_access_token)

        _line("═")
        print(f"STEP 1 — meeting(identifier={subject!r})")
        _line()
        meeting_out = await compose_meeting(
            client=client,
            permissions=permissions,
            identifier=subject,
            timezone=config.default_timezone,
        )
        print(meeting_out)

        _line("═")
        print("STEP 2 — resolve(<chat thread URL>)")
        _line()
        target = await _find_target_thread(client, subject)
        if not target:
            print("  ❌ No meeting with recording found for subject match.")
            return
        thread_id, matched_subject = target
        # Construct a chat URL using the thread ID.
        chat_url = f"https://teams.microsoft.com/l/chat/{urllib.parse.quote(thread_id, safe='')}"
        print(f"  Matched: {matched_subject}")
        print(f"  Chat URL: {chat_url[:80]}…")
        print()
        resolve_out = await compose_resolve(
            client=client,
            permissions=permissions,
            url=chat_url,
        )
        print(resolve_out)

        _line("═")
        print("VERIFICATION")
        _line()
        meeting_has_recording = "**Recording:**" in meeting_out
        meeting_has_vroom = "vroom_url:" in meeting_out
        chat_has_call = "📞" in resolve_out
        chat_has_vroom = "vroom_url:" in resolve_out
        print(f"  meeting() shows Recording block:      {'✅' if meeting_has_recording else '❌'}")
        print(f"  meeting() Recording has vroom_url:    {'✅' if meeting_has_vroom else '❌'}")
        print(f"  resolve() chat thread has call entry: {'✅' if chat_has_call else '❌'}")
        print(f"  resolve() call has vroom_url sub-bullet: {'✅' if chat_has_vroom else '❌'}")
        if meeting_has_recording and meeting_has_vroom and chat_has_call and chat_has_vroom:
            print("\n  ✅ All checks pass — recording metadata is exposed end-to-end")
        else:
            print("\n  ⚠️  Some checks failed — inspect output above")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    query = " ".join(sys.argv[1:])
    asyncio.run(run(query))
