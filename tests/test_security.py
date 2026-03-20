import subprocess
import json
import sys
import pytest


def test_bandit_scan_no_high_severity():
    """Run bandit on the source directory and assert there are no HIGH severity issues.
    This test requires 'bandit' to be installed in the environment (add to dev requirements).
    """
    try:
        completed = subprocess.run(
            [sys.executable, "-m", "bandit", "-r", "src", "-f", "json"],
            capture_output=True,
            text=True,
            check=True,
        )
    except subprocess.CalledProcessError as e:
        # bandit returns exit code 1 when issues are found, but we still want to parse its output
        completed = e

    try:
        data = json.loads(completed.stdout)
    except json.JSONDecodeError:
        pytest.skip("Bandit output could not be decoded; is bandit installed?\n" + completed.stdout)

    results = data.get("results", [])
    high_issues = [issue for issue in results if issue.get("issue_severity") == "HIGH"]
    assert not high_issues, f"Found high severity security issues:\n{high_issues}"


def test_pip_audit_no_critical():
    """Run pip-audit to ensure no critical vulnerabilities in installed packages."""
    # if pip-audit isn't installed we should skip rather than fail
    try:
        import pip_audit  # just to check availability
    except ImportError:
        pytest.skip("pip-audit not installed")

    try:
        completed = subprocess.run(
            [sys.executable, "-m", "pip_audit", "--fail-on", "critical"],
            capture_output=True,
            text=True,
            check=True,
        )
    except subprocess.CalledProcessError as e:
        # pip-audit returned a non-zero exit code (vulnerabilities found)
        pytest.fail(f"pip-audit detected vulnerabilities:\n{e.stdout}\n{e.stderr}")