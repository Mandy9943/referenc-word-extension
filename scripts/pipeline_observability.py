#!/usr/bin/env python3
"""
Shared run-observability utilities for DOCX/PPTX pipelines.

Writes a human-readable markdown log on Desktop:
  ~/Desktop/add-in-logs.md
"""

from __future__ import annotations

import json
import os
import socket
import traceback
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

LOG_FILE_PATH = Path.home() / "Desktop" / "add-in-logs.md"
LOG_HEADER = """# Add-in Pipeline Logs

Auto-generated run telemetry for DOCX/PPTX automation.
This file logs success/failure runs, timings, request behavior, and improvement hints.

"""


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def local_now_pretty() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S %z")


def _ensure_log_header(path: Path) -> None:
    if path.exists():
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(LOG_HEADER, encoding="utf-8")


def _safe_float(value: Any, default: float = 0.0) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    return default


def _error_analysis(error_message: str) -> Dict[str, List[str]]:
    lower = (error_message or "").lower()
    causes: List[str] = []
    fixes: List[str] = []

    if "paraphrase button not responding" in lower or "e_selector_miss" in lower:
        causes.append("QuillBot UI event did not transition into a detectable running state.")
        fixes.append("Restart accounts (`/restart`), then retry with smaller chunk size.")
        fixes.append("Verify QuillBot tab is focused and no modal/overlay blocks button events.")

    if "tripped until" in lower or "e_throttled" in lower:
        causes.append("Account health guard tripped after repeated failures/rate-limit behavior.")
        fixes.append("Wait for tripped cooldown to expire or restart specific account workers.")
        fixes.append("Reduce per-request words or temporary concurrency to lower throttle pressure.")

    if "response count mismatch" in lower:
        causes.append("Delimiter/segmentation mismatch between sent paragraphs and returned parts.")
        fixes.append("Keep paragraphs shorter and avoid unusually long single paragraphs.")
        fixes.append("Use recovery split fallback (already enabled); inspect if mismatch persists.")

    if "no eligible" in lower:
        causes.append("Pipeline filters excluded all candidate paragraphs for that step.")
        fixes.append("Check headings, references-only docs, and minimum-word thresholds.")

    if "could not detect a reference section" in lower:
        causes.append("Reference section pattern detection failed on current document layout.")
        fixes.append("Ensure references are grouped in a clear section at the end of the document.")

    if "invalid" in lower and ("docx" in lower or "pptx" in lower):
        causes.append("Input package appears malformed or missing required XML parts.")
        fixes.append("Open and resave source file in Office, then rerun pipeline.")

    if "input file does not exist" in lower or "no .docx file found on desktop" in lower or "no .pptx file found on desktop" in lower:
        causes.append("Input path resolution failed (missing file or no eligible Desktop source file).")
        fixes.append("Place exactly one source file on Desktop, or pass an explicit absolute file path.")

    if "found " in lower and "files on desktop" in lower:
        causes.append("Auto-pick guard found multiple Desktop files and intentionally stopped.")
        fixes.append("Keep only one source file on Desktop or pass explicit input path.")

    if not causes:
        causes.append("Unclassified runtime error.")
        fixes.append("Check the raw error stack and latest server logs for precise failure point.")

    return {"causes": causes, "fixes": fixes}


def _performance_suggestions(run: Dict[str, Any]) -> List[str]:
    suggestions: List[str] = []
    status = str(run.get("status", "")).lower()
    mode = str(run.get("mode", "dual"))
    total_words = int(run.get("step2_total_words") or 0)
    total_seconds = _safe_float(run.get("total_duration_seconds"))
    step2_seconds = _safe_float(run.get("step_durations", {}).get("step2_paraphrase"))
    request_summaries = run.get("step2_request_summaries") or []
    step2_initial_failures = int(run.get("step2_initial_chunk_failures") or 0)
    step2_recoveries = int(run.get("step2_recovery_attempts") or 0)

    if status == "success" and total_words > 0 and total_seconds > 0:
        words_per_sec = total_words / max(total_seconds, 0.001)
        if mode == "dual" and words_per_sec < 8.0:
            suggestions.append(
                "Throughput is low for SIMPLE+SHORT. Keep all 3 accounts healthy and reduce oversized chunks."
            )
        if mode == "standard" and words_per_sec < 12.0:
            suggestions.append(
                "Throughput is low for STANDARD. Consider increasing words/request if failure rate stays low."
            )

    if step2_seconds > 0 and total_seconds > 0 and (step2_seconds / total_seconds) > 0.85:
        suggestions.append(
            "Paraphrase step dominates runtime. Focus optimization on request sizing and account health."
        )

    if isinstance(request_summaries, list) and len(request_summaries) >= 4:
        avg_words = sum(int(item.get("words", 0)) for item in request_summaries) / max(len(request_summaries), 1)
        if avg_words < 450:
            suggestions.append(
                "Average words/request is small; overhead is high. Increase max words/request cautiously."
            )

    if isinstance(request_summaries, list) and request_summaries:
        estimated_total = sum(_safe_float(item.get("estimated_seconds")) for item in request_summaries)
        if step2_seconds > 0 and estimated_total > 0:
            estimate_ratio = step2_seconds / max(estimated_total, 0.001)
            if estimate_ratio > 1.4:
                suggestions.append(
                    "Real step2 time is much higher than estimate; likely retries, throttling, or click/start failures."
                )
            elif estimate_ratio < 0.65:
                suggestions.append(
                    "Real step2 time is much lower than estimate; scheduler capacity assumptions may be too conservative."
                )

        trimmed_count = sum(1 for item in request_summaries if bool(item.get("trimmed_for_single_account")))
        if trimmed_count > 0:
            suggestions.append(
                f"Single-account guardrail trimmed {trimmed_count} request(s); account health was constrained during run."
            )

        account_mismatch_count = 0
        imbalance_count = 0
        for item in request_summaries:
            accounts = [str(a) for a in (item.get("accounts") or [])]
            chunk_words = item.get("chunk_words_by_account") or {}
            if isinstance(chunk_words, dict) and chunk_words:
                chunk_keys = [str(key) for key in chunk_words.keys()]
                if accounts and set(accounts) != set(chunk_keys):
                    account_mismatch_count += 1

                totals = [max(0, int(value)) for value in chunk_words.values() if isinstance(value, (int, float))]
                if len(totals) > 1 and sum(totals) > 0:
                    largest_share = max(totals) / sum(totals)
                    if largest_share > 0.75:
                        imbalance_count += 1

        if account_mismatch_count > 0:
            suggestions.append(
                f"Detected account-selection mismatch in {account_mismatch_count} request(s); verify chunk keys match selected accounts."
            )

        if imbalance_count > 0:
            suggestions.append(
                f"Detected heavy account imbalance in {imbalance_count} request(s); rebalance chunks by words to improve parallelism."
            )

    if step2_initial_failures > 0:
        suggestions.append(
            f"Detected {step2_initial_failures} initial chunk failure(s); verify QuillBot UI stability and selector health."
        )

    if step2_recoveries > 0:
        suggestions.append(
            f"Detected {step2_recoveries} recovery fallback invocation(s); monitor delimiter and chunk boundary quality."
        )

    if not suggestions:
        suggestions.append("No immediate optimization flags from this run.")

    return suggestions


def append_pipeline_log(run: Dict[str, Any]) -> None:
    """
    Append one run entry to Desktop markdown logs.
    Expects a dict with fields such as:
      workflow, status, mode, input_path, output_path, total_duration_seconds,
      step_durations, step1_stats, step2_stats, step3_stats, step2_request_summaries,
      error_message, error_step, traceback
    """
    try:
        _ensure_log_header(LOG_FILE_PATH)
        status = str(run.get("status", "unknown")).upper()
        workflow = str(run.get("workflow", "unknown")).upper()
        mode = str(run.get("mode", "unknown"))
        started = str(run.get("started_at_local", local_now_pretty()))
        run_id = str(run.get("run_id", f"{workflow.lower()}-{int(datetime.now().timestamp())}"))
        input_path = str(run.get("input_path", ""))
        output_path = str(run.get("output_path", ""))
        total_seconds = _safe_float(run.get("total_duration_seconds"))
        step_durations = run.get("step_durations", {})
        request_summaries = run.get("step2_request_summaries") or []
        error_message = str(run.get("error_message", "") or "")
        error_step = str(run.get("error_step", "") or "")
        host = run.get("host") or socket.gethostname()
        user = run.get("user") or os.getenv("USER", "")

        error_enrichment = _error_analysis(error_message) if error_message else {"causes": [], "fixes": []}
        perf_suggestions = _performance_suggestions(run)

        lines: List[str] = []
        lines.append(f"## {started} | {workflow} | {status}")
        lines.append("")
        lines.append(f"- `run_id`: `{run_id}`")
        lines.append(f"- `host`: `{host}`")
        lines.append(f"- `user`: `{user}`")
        lines.append(f"- `mode`: `{mode}`")
        lines.append(f"- `input`: `{input_path}`")
        lines.append(f"- `output`: `{output_path}`")
        lines.append(f"- `total_duration_s`: `{total_seconds:.2f}`")
        if run.get("dry_run") is True:
            lines.append("- `dry_run`: `true`")
        lines.append("")

        lines.append("### Step Timings")
        if isinstance(step_durations, dict) and step_durations:
            for key, value in step_durations.items():
                lines.append(f"- `{key}`: `{_safe_float(value):.2f}s`")
        else:
            lines.append("- (no step timing data)")
        lines.append("")

        lines.append("### Step Metrics")
        for section_key in ("step1_stats", "step2_stats", "step3_stats"):
            section = run.get(section_key)
            if isinstance(section, dict) and section:
                lines.append(f"- `{section_key}`: `{json.dumps(section, ensure_ascii=False)}`")
        lines.append("")

        lines.append("### Request Breakdown")
        if isinstance(request_summaries, list) and request_summaries:
            for item in request_summaries:
                idx = item.get("request_index")
                words = item.get("words")
                paragraphs = item.get("paragraphs")
                accounts = item.get("accounts", [])
                acc_count = item.get("account_count")
                est = item.get("estimated_seconds")
                cap = item.get("effective_capacity")
                trimmed = item.get("trimmed_for_single_account", False)
                chunk_words = item.get("chunk_words_by_account", {})
                lines.append(
                    f"- `req {idx}` paragraphs={paragraphs}, words={words}, accounts={acc_count} {accounts}, "
                    f"est={est:.1f}s, capacity={cap:.0f}, trimmed_single_account={trimmed}, chunks={chunk_words}"
                )
        else:
            lines.append("- (no request-level data)")
        lines.append("")

        if error_message:
            lines.append("### Error")
            lines.append(f"- `step`: `{error_step or 'unknown'}`")
            lines.append(f"- `message`: `{error_message}`")
            lines.append("")
            lines.append("### Likely Causes")
            for cause in error_enrichment["causes"]:
                lines.append(f"- {cause}")
            lines.append("")
            lines.append("### Suggested Fixes")
            for fix in error_enrichment["fixes"]:
                lines.append(f"- {fix}")
            lines.append("")

        lines.append("### Improvement Suggestions")
        for suggestion in perf_suggestions:
            lines.append(f"- {suggestion}")
        lines.append("")

        if run.get("traceback"):
            lines.append("<details><summary>Traceback</summary>")
            lines.append("")
            lines.append("```text")
            lines.append(str(run["traceback"]))
            lines.append("```")
            lines.append("")
            lines.append("</details>")
            lines.append("")

        lines.append("<details><summary>Raw Run Payload</summary>")
        lines.append("")
        lines.append("```json")
        lines.append(json.dumps(run, ensure_ascii=False, indent=2))
        lines.append("```")
        lines.append("")
        lines.append("</details>")
        lines.append("")
        lines.append("---")
        lines.append("")

        with LOG_FILE_PATH.open("a", encoding="utf-8") as handle:
            handle.write("\n".join(lines))
    except Exception:
        # Logging must never break the pipeline.
        try:
            fallback = traceback.format_exc()
            with (Path.home() / "Desktop" / "add-in-logs-fallback.txt").open(
                "a", encoding="utf-8"
            ) as handle:
                handle.write(
                    f"[{utc_now_iso()}] Failed to append pipeline log:\n{fallback}\n\n"
                )
        except Exception:
            pass
