"""Dependency-audit outcome classification contracts."""

import json

from scripts.audit_dependencies import interpret_audit


def _payload(vulns: list[dict[str, object]] | None = None) -> str:
    return json.dumps(
        {
            "dependencies": [
                {
                    "name": "example",
                    "version": "1.0",
                    "vulns": vulns or [],
                }
            ]
        }
    )


def test_clean_dependency_audit_returns_zero(capsys) -> None:
    assert interpret_audit(0, _payload(), "") == 0
    assert "passed" in capsys.readouterr().out


def test_vulnerability_is_distinct_from_tool_failure(capsys) -> None:
    assert interpret_audit(1, _payload([{"id": "CVE-2026-0001"}]), "") == 2
    assert "CVE-2026-0001" in capsys.readouterr().out


def test_unavailable_feed_or_malformed_output_returns_one(capsys) -> None:
    assert interpret_audit(1, "not-json", "network unavailable") == 1
    assert "unavailable" in capsys.readouterr().err
