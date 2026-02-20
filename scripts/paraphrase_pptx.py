#!/usr/bin/env python3
"""
Bulk paraphrase PowerPoint slide text + speaker notes via the QuillBot batch API.

This script works on .pptx files directly (zip + XML), so it can process an entire
deck in one run and write paraphrased text back to the exact slide/notes paragraphs.
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from collections import deque
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple

API_URL_DEFAULT = "https://analizeai.com/paraphrase-batch"
PARAPHRASE_DELIMITER = "qbpdelim123"
ACCOUNT_KEYS = ("acc1", "acc2", "acc3")
ZERO_WIDTH_RE = re.compile(r"[\u200B-\u200D\uFEFF]")

SLIDE_XML_RE = re.compile(r"^ppt/slides/slide(\d+)\.xml$")
NOTES_XML_RE = re.compile(r"^ppt/notesSlides/notesSlide(\d+)\.xml$")

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}

# Keep these in sync with taskpane word/powerpoint logic.
TRAILING_PUNCTUATION = r"[-:;,.!?-]*"
NUMBERING_PREFIX = r"(?:(?:\d+(?:\.\d+)*|[IVX]+)\.?\s*)?"
REFERENCE_HEADER_PATTERNS = [
    re.compile(
        rf"^\s*{NUMBERING_PREFIX}references?(?:\s+list)?(?:\s+section)?\s*{TRAILING_PUNCTUATION}\s*$",
        re.IGNORECASE,
    ),
    re.compile(
        rf"^\s*{NUMBERING_PREFIX}reference\s+list\s*{TRAILING_PUNCTUATION}\s*$",
        re.IGNORECASE,
    ),
    re.compile(
        rf"^\s*{NUMBERING_PREFIX}references\s+list\s*{TRAILING_PUNCTUATION}\s*$",
        re.IGNORECASE,
    ),
    re.compile(
        rf"^\s*{NUMBERING_PREFIX}bibliograph(?:y|ies)\s*{TRAILING_PUNCTUATION}\s*$",
        re.IGNORECASE,
    ),
    re.compile(
        rf"^\s*{NUMBERING_PREFIX}list\s+of\s+references\s*{TRAILING_PUNCTUATION}\s*$",
        re.IGNORECASE,
    ),
]


@dataclass
class Candidate:
    archive_path: str
    kind: str  # "slide" | "notes"
    file_number: int
    paragraph_index: int
    text_nodes: List[ET.Element]
    original_text: str
    word_count: int
    paraphrased_text: str | None = None


def sanitize_text(text: str) -> str:
    return ZERO_WIDTH_RE.sub("", text or "").strip()


def count_words(text: str) -> int:
    trimmed = sanitize_text(text)
    if not trimmed:
        return 0
    return len([token for token in trimmed.split() if token])


def matches_reference_header(text: str) -> bool:
    trimmed = sanitize_text(text)
    if not trimmed:
        return False
    if any(pattern.match(trimmed) for pattern in REFERENCE_HEADER_PATTERNS):
        return True
    first_line = trimmed.splitlines()[0].strip() if "\n" in trimmed else trimmed
    return any(pattern.match(first_line) for pattern in REFERENCE_HEADER_PATTERNS)


def choose_account_count(total_words: int) -> int:
    if total_words > 1500:
        return 3
    if total_words >= 500:
        return 2
    return 1


def parse_paraphrase_parts(text: str, expected_count: int) -> List[str]:
    parts = [p.strip() for p in re.split(re.escape(PARAPHRASE_DELIMITER), text, flags=re.IGNORECASE) if p.strip()]

    if len(parts) < expected_count:
        recovered: List[str] = []
        for part in parts:
            if "\n\n" in part:
                recovered.extend([p.strip() for p in re.split(r"\n\n+", part) if p.strip()])
            else:
                recovered.append(part)
        if len(recovered) == expected_count:
            return recovered
        if abs(len(recovered) - expected_count) < abs(len(parts) - expected_count):
            return recovered

    return parts


def build_payload_text(items: Sequence[Candidate]) -> str:
    chunks: List[str] = []
    for item in items:
        chunks.append(PARAPHRASE_DELIMITER)
        chunks.append(item.original_text)
    return "\n\n".join(chunks)


def split_into_account_chunks(items: Sequence[Candidate], account_count: int) -> List[Tuple[str, List[Candidate]]]:
    chunk_size = math.ceil(len(items) / account_count)
    chunks: List[Tuple[str, List[Candidate]]] = []
    for i in range(account_count):
        start = i * chunk_size
        end = min(start + chunk_size, len(items))
        if start >= len(items):
            break
        chunk = list(items[start:end])
        if chunk:
            chunks.append((ACCOUNT_KEYS[i], chunk))
    return chunks


def post_batch_request(api_url: str, payload: Dict[str, str], timeout_seconds: int) -> Dict[str, object]:
    body = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(api_url, data=body, method="POST")
    req.add_header("Content-Type", "application/json")

    try:
        with urllib.request.urlopen(req, timeout=timeout_seconds) as response:
            response_body = response.read().decode("utf-8")
            if response.status < 200 or response.status >= 300:
                raise RuntimeError(f"Batch API returned HTTP {response.status}: {response_body[:400]}")
            return json.loads(response_body)
    except urllib.error.HTTPError as err:
        details = (err.read().decode("utf-8", errors="replace") if hasattr(err, "read") else "")[:400]
        raise RuntimeError(f"Batch API returned HTTP {err.code}: {details}") from err
    except urllib.error.URLError as err:
        raise RuntimeError(f"Failed to reach batch API: {err}") from err


def sorted_target_xml_paths(
    names: Iterable[str], include_slides: bool, include_notes: bool
) -> List[Tuple[str, str, int]]:
    targets: List[Tuple[str, str, int]] = []
    for name in names:
        slide_match = SLIDE_XML_RE.match(name)
        if slide_match and include_slides:
            targets.append((name, "slide", int(slide_match.group(1))))
            continue
        notes_match = NOTES_XML_RE.match(name)
        if notes_match and include_notes:
            targets.append((name, "notes", int(notes_match.group(1))))

    # Keep deterministic order: slide1, notes1, slide2, notes2, ...
    def sort_key(item: Tuple[str, str, int]) -> Tuple[int, int]:
        path, kind, number = item
        kind_rank = 0 if kind == "slide" else 1
        return number, kind_rank

    return sorted(targets, key=sort_key)


def extract_candidates(
    input_zip: zipfile.ZipFile, mode: str, include_slides: bool, include_notes: bool
) -> Tuple[Dict[str, ET.ElementTree], List[Candidate]]:
    xml_docs: Dict[str, ET.ElementTree] = {}
    candidates: List[Candidate] = []

    for archive_path, kind, file_number in sorted_target_xml_paths(
        input_zip.namelist(), include_slides, include_notes
    ):
        raw_xml = input_zip.read(archive_path)
        root = ET.fromstring(raw_xml)
        tree = ET.ElementTree(root)
        xml_docs[archive_path] = tree

        # PowerPoint slide/notes text bodies are usually p:txBody with a:p children.
        # Keep a fallback for a:txBody variants from non-standard generators.
        paragraphs = root.findall(".//p:txBody/a:p", NS)
        if not paragraphs:
            paragraphs = root.findall(".//a:txBody/a:p", NS)
        first_non_empty_seen = False
        for paragraph_index, paragraph in enumerate(paragraphs):
            text_nodes = paragraph.findall(".//a:t", NS)
            if not text_nodes:
                continue

            text = sanitize_text("".join(node.text or "" for node in text_nodes))
            if not text:
                continue

            words = count_words(text)
            is_title = False
            if kind == "slide" and not first_non_empty_seen:
                is_title = words < 10
            first_non_empty_seen = True

            if matches_reference_header(text):
                continue

            if mode == "dual":
                if is_title or words < 11 or text.endswith(":"):
                    continue
            else:
                if words < 15:
                    continue

            candidates.append(
                Candidate(
                    archive_path=archive_path,
                    kind=kind,
                    file_number=file_number,
                    paragraph_index=paragraph_index,
                    text_nodes=text_nodes,
                    original_text=text,
                    word_count=words,
                )
            )

    return xml_docs, candidates


def take_request_batch(
    remaining: deque[Candidate], max_items_per_request: int, max_words_per_request: int
) -> List[Candidate]:
    batch: List[Candidate] = []
    total_words = 0

    while remaining and len(batch) < max_items_per_request:
        next_item = remaining[0]
        if batch and (total_words + next_item.word_count > max_words_per_request):
            break
        batch.append(remaining.popleft())
        total_words += next_item.word_count

    if not batch and remaining:
        batch.append(remaining.popleft())

    return batch


def paraphrase_candidates(
    candidates: List[Candidate],
    mode: str,
    api_url: str,
    timeout_seconds: int,
    max_items_per_request: int,
    max_words_per_request: int,
) -> None:
    remaining = deque(candidates)
    request_index = 0

    while remaining:
        request_index += 1
        batch = take_request_batch(remaining, max_items_per_request, max_words_per_request)
        batch_word_count = sum(item.word_count for item in batch)
        account_count = choose_account_count(batch_word_count)
        account_chunks = split_into_account_chunks(batch, account_count)

        payload: Dict[str, str] = {"mode": mode}
        for account_key, account_items in account_chunks:
            payload[account_key] = build_payload_text(account_items)

        print(
            f"[request {request_index}] sending {len(batch)} paragraphs "
            f"({batch_word_count} words) across {len(account_chunks)} account(s)"
        )
        response = post_batch_request(api_url, payload, timeout_seconds)

        for account_key, account_items in account_chunks:
            account_result = response.get(account_key)
            if not isinstance(account_result, dict):
                raise RuntimeError(f"Missing response for account {account_key} in request {request_index}")

            if mode == "dual":
                paraphrased = account_result.get("secondMode")
            else:
                paraphrased = account_result.get("result")

            if not paraphrased:
                error_message = account_result.get("error", "missing paraphrased output")
                raise RuntimeError(f"Account {account_key} failed in request {request_index}: {error_message}")

            parts = parse_paraphrase_parts(str(paraphrased), len(account_items))
            if len(parts) != len(account_items):
                raise RuntimeError(
                    f"Response count mismatch in request {request_index} account {account_key}: "
                    f"expected {len(account_items)}, got {len(parts)}"
                )

            for candidate, new_text in zip(account_items, parts):
                candidate.paraphrased_text = new_text.strip()

            if account_result.get("fallbackUsed"):
                print(
                    f"[request {request_index}] warning: {account_key} used fallback "
                    f"{account_result['fallbackUsed']}"
                )


def apply_candidate_updates(candidates: Sequence[Candidate]) -> int:
    updated = 0
    for candidate in candidates:
        if candidate.paraphrased_text is None:
            continue
        if not candidate.text_nodes:
            continue
        candidate.text_nodes[0].text = candidate.paraphrased_text
        for node in candidate.text_nodes[1:]:
            node.text = ""
        updated += 1
    return updated


def write_output_pptx(
    input_path: Path, output_path: Path, xml_docs: Dict[str, ET.ElementTree]
) -> None:
    with zipfile.ZipFile(input_path, "r") as source_zip, zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as out_zip:
        for info in source_zip.infolist():
            if info.filename in xml_docs:
                root = xml_docs[info.filename].getroot()
                xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                out_zip.writestr(info, xml_bytes)
            else:
                out_zip.writestr(info, source_zip.read(info.filename))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Paraphrase a PPTX in bulk (slides + speaker notes) and write output back to a new PPTX."
        )
    )
    parser.add_argument(
        "input",
        nargs="?",
        type=Path,
        help="Input .pptx file path. If omitted, auto-detects a single Desktop .pptx.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output .pptx file path (default: pr <input-name>.pptx)",
    )
    parser.add_argument(
        "--mode",
        choices=("dual", "standard"),
        default="dual",
        help="Paraphrase mode to request from batch API",
    )
    parser.add_argument(
        "--api-url",
        default=API_URL_DEFAULT,
        help=f"Batch API URL (default: {API_URL_DEFAULT})",
    )
    parser.add_argument(
        "--timeout-seconds",
        type=int,
        default=180,
        help="HTTP timeout for each batch request",
    )
    parser.add_argument(
        "--max-items-per-request",
        type=int,
        default=120,
        help="Max paragraphs sent in one API request",
    )
    parser.add_argument(
        "--max-words-per-request",
        type=int,
        default=2400,
        help="Approximate max total words sent in one API request",
    )
    parser.add_argument(
        "--no-slides",
        action="store_true",
        help="Do not paraphrase slide text boxes",
    )
    parser.add_argument(
        "--no-notes",
        action="store_true",
        help="Do not paraphrase speaker notes",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Only inspect and report eligible paragraph counts; do not call API or write output",
    )
    return parser.parse_args()


def resolve_input_path(provided_input: Path | None) -> Path:
    if provided_input is not None:
        return provided_input

    desktop = Path.home() / "Desktop"
    if not desktop.exists() or not desktop.is_dir():
        raise RuntimeError(f"Desktop folder not found: {desktop}")

    desktop_pptx = [
        path
        for path in desktop.iterdir()
        if path.is_file()
        and path.suffix.lower() == ".pptx"
        and not path.name.startswith("~$")
        and not path.name.lower().endswith(".paraphrased.pptx")
    ]

    existing_names = {path.name.lower() for path in desktop_pptx}
    candidates: List[Path] = []
    for path in desktop_pptx:
        lower_name = path.name.lower()
        if lower_name.startswith("pr "):
            original_name = path.name[3:]
            # Treat "pr <name>.pptx" as generated output only when "<name>.pptx"
            # exists on Desktop in the same moment.
            if original_name.lower() in existing_names:
                continue
        candidates.append(path)

    candidates = sorted(candidates)

    if len(candidates) == 1:
        print(f"Auto-selected Desktop file: {candidates[0]}")
        return candidates[0]

    if len(candidates) == 0:
        raise RuntimeError(
            f"No .pptx file found on Desktop ({desktop}). Put one file there or pass the path explicitly."
        )

    preview = ", ".join(path.name for path in candidates[:5])
    suffix = "..." if len(candidates) > 5 else ""
    raise RuntimeError(
        f"Found {len(candidates)} .pptx files on Desktop ({preview}{suffix}). "
        "Keep only one file there or pass the path explicitly."
    )


def main() -> int:
    args = parse_args()
    try:
        input_path = resolve_input_path(args.input)
    except RuntimeError as error:
        print(str(error), file=sys.stderr)
        return 1

    if not input_path.exists():
        print(f"Input file does not exist: {input_path}", file=sys.stderr)
        return 1
    if input_path.suffix.lower() != ".pptx":
        print(f"Input must be a .pptx file: {input_path}", file=sys.stderr)
        return 1

    include_slides = not args.no_slides
    include_notes = not args.no_notes
    if not include_slides and not include_notes:
        print("Nothing to do: both slides and notes are disabled.", file=sys.stderr)
        return 1

    output_path = (
        args.output
        if args.output
        else input_path.with_name(f"pr {input_path.name}")
    )

    with zipfile.ZipFile(input_path, "r") as input_zip:
        xml_docs, candidates = extract_candidates(
            input_zip=input_zip,
            mode=args.mode,
            include_slides=include_slides,
            include_notes=include_notes,
        )

    slide_count = sum(1 for c in candidates if c.kind == "slide")
    notes_count = sum(1 for c in candidates if c.kind == "notes")
    total_words = sum(c.word_count for c in candidates)
    print(
        f"Eligible paragraphs: total={len(candidates)}, slide={slide_count}, notes={notes_count}, words={total_words}"
    )

    if not candidates:
        print("No eligible text found. Exiting without output changes.")
        return 0

    if args.dry_run:
        print("Dry run complete.")
        return 0

    paraphrase_candidates(
        candidates=candidates,
        mode=args.mode,
        api_url=args.api_url,
        timeout_seconds=args.timeout_seconds,
        max_items_per_request=args.max_items_per_request,
        max_words_per_request=args.max_words_per_request,
    )

    updated = apply_candidate_updates(candidates)
    write_output_pptx(input_path, output_path, xml_docs)
    print(f"Done. Updated {updated} paragraphs. Output: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
