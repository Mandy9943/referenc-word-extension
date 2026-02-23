#!/usr/bin/env python3
"""Reset Desktop add-in logs file to a clean header."""

from pathlib import Path

LOG_FILE_PATH = Path.home() / "Desktop" / "add-in-logs.md"
LOG_HEADER = """# Add-in Pipeline Logs

Auto-generated run telemetry for DOCX/PPTX automation.
This file logs success/failure runs, timings, request behavior, and improvement hints.

"""


def main() -> int:
    LOG_FILE_PATH.parent.mkdir(parents=True, exist_ok=True)
    LOG_FILE_PATH.write_text(LOG_HEADER, encoding="utf-8")
    print(f"Reset logs: {LOG_FILE_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
