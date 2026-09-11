"""Public-tree privacy policy contracts."""

from scripts.audit_public_tree import _PROHIBITED_PUBLIC_PATTERNS


def test_private_workload_codename_is_rejected_in_content_and_paths() -> None:
    private_codename = bytes.fromhex("54584f").decode("ascii")
    pattern = _PROHIBITED_PUBLIC_PATTERNS["private workload codename"]

    assert pattern.search(f"private {private_codename} benchmark")
    assert pattern.search(f"scripts/{private_codename.lower()}_acceptance.py")
    assert pattern.search("representative large-workbook benchmark") is None
