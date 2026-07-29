from ms365_intent_mcp.windows_tz import WINDOWS_TO_IANA, windows_to_iana


def test_common_windows_names_map_to_iana():
    assert windows_to_iana("W. Europe Standard Time") == "Europe/Berlin"
    assert windows_to_iana("UTC") == "Etc/UTC"
    assert windows_to_iana("Eastern Standard Time") == "America/New_York"
    assert windows_to_iana("Pacific Standard Time") == "America/Los_Angeles"


def test_unmapped_name_returns_none():
    assert windows_to_iana("Nonexistent Standard Time") is None


def test_map_is_nonempty_dict():
    assert isinstance(WINDOWS_TO_IANA, dict)
    assert len(WINDOWS_TO_IANA) > 100
