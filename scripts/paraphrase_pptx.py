#!/usr/bin/env python3
"""
One-shot PPTX pipeline for:
1) cleaning in-text citations + weird artifacts (while preserving headings/subtitles),
2) paraphrasing eligible slide + speaker-notes paragraphs via batch API,
3) inserting refreshed in-text references into slide + notes body text,
4) writing output as: pr <original-name>.pptx
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
import traceback
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from collections import deque
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple

from pipeline_observability import append_pipeline_log, local_now_pretty, utc_now_iso

API_URL_DEFAULT = "https://analizeai.com/paraphrase-batch"
PARAPHRASE_DELIMITER = "qbpdelim123"
ACCOUNT_KEYS = ("acc1", "acc2", "acc3")
MODE_DEFAULT_BUDGET = {"dual": 520.0, "standard": 950.0, "ludicrous": 600.0}
MODE_TARGET_SECONDS = {"dual": 18.0, "standard": 9.0, "ludicrous": 16.0}
MODE_MIN_WORDS_PER_ACCOUNT = {"dual": 280.0, "standard": 500.0, "ludicrous": 300.0}
MODE_MAX_WORDS_PER_ACCOUNT = {"dual": 760, "standard": 1400, "ludicrous": 900}
MODE_COORDINATION_PENALTY_SECONDS = {"dual": 1.2, "standard": 0.7, "ludicrous": 1.0}
MODE_RATE_SUFFIX = {"dual": "Dual", "standard": "Standard", "ludicrous": "Ludicrous"}

ZERO_WIDTH_RE = re.compile(r"[\u200B-\u200D\uFEFF]")
XML_INVALID_CHAR_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
SLIDE_XML_RE = re.compile(r"^ppt/slides/slide(\d+)\.xml$")
NOTES_XML_RE = re.compile(r"^ppt/notesSlides/notesSlide(\d+)\.xml$")

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}
XSI_TYPE_ATTR = "{http://www.w3.org/2001/XMLSchema-instance}type"
METADATA_FIXED_TIMESTAMP = "2000-01-01T00:00:00Z"
PPTX_METADATA_XML_PATHS = {
    "docProps/core.xml",
    "docProps/app.xml",
    "docProps/custom.xml",
    "ppt/commentAuthors.xml",
}

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
    re.compile(
        rf"^\s*{NUMBERING_PREFIX}works\s+cited\s*{TRAILING_PUNCTUATION}\s*$",
        re.IGNORECASE,
    ),
    re.compile(
        rf"^\s*{NUMBERING_PREFIX}sources?\s*{TRAILING_PUNCTUATION}\s*$",
        re.IGNORECASE,
    ),
]

YEAR_RE = re.compile(r"\b(?:19|20)\d{2}[a-z]?\b")
URL_OR_DOI_RE = re.compile(r"\b(?:https?://|www\.|doi:\s*|10\.\d{4,9}/)\S+", re.IGNORECASE)
REFERENCE_CUE_RE = re.compile(
    r"\b(?:available at|retrieved from|accessed|doi|journal|vol\.?|no\.?|pp\.?|edition|ed\.)\b",
    re.IGNORECASE,
)
AUTHOR_RE = re.compile(r"(?:^|[\s;])(?:[A-Z][a-z]+,\s*(?:[A-Z]\.|[A-Z][a-z]+))")
LIST_PREFIX_RE = re.compile(r"^\s*(?:\d{1,3}[.)\]]|[-•])\s+")

CITATION_PATTERNS = [
    re.compile(r"\[(?:[^\]]+)[,\s]\s?\d{4}[a-z]?\]"),
    re.compile(r"\((?:[^,()]+(,\s[^,()]+)*(?:,\sand\s[^,()]+)?)[,\s]\s?\d{4}[a-z]?\)"),
    re.compile(r"\((?:[^,()]+)[,\s]\s?\d{4}[a-z]?\)"),
    re.compile(r"\((?:[^()]+\sand\s[^,()]+)[,\s]\s?\d{4}[a-z]?\)"),
    re.compile(r"\((?:[^()]+)\set\sal\.?[,\s]\s?\d{4}[a-z]?\)"),
    re.compile(r"\((?:[^,()]+(,\s[^,()]+)*)[,\s]\s?\d{4}[a-z]?\)"),
]
WEIRD_NUMBER_PATTERN = re.compile(r"[【\[]\d+.*?[†+t].*?[】\]]\S*")
EXISTING_CITATION_RE = re.compile(r"\(\s*[^)]*?\d{4}[a-z]?\s*\)")
LINK_PATTERNS = [
    re.compile(r"\bhttps?://[^\s)\]}]+", re.IGNORECASE),
    re.compile(r"\bwww\.[^\s)\]}]+", re.IGNORECASE),
    re.compile(r"\b(?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,}(?:/[^\s)\]}]*)?", re.IGNORECASE),
]
XMLNS_DECLARATION_RE = re.compile(r"""xmlns(?::([A-Za-z_][\w.\-]*))?\s*=\s*(['"])(.*?)\2""")

ParagraphKey = Tuple[str, int]


@dataclass
class PptParagraph:
    archive_path: str
    kind: str  # "slide" | "notes"
    file_number: int
    paragraph_index: int
    global_index: int
    text_nodes: List[ET.Element]
    text: str
    word_count: int
    is_first_non_empty: bool


@dataclass
class ReferenceDetection:
    mode: str
    reference_keys: Set[ParagraphKey]
    inferred_kind: Optional[str] = None
    inferred_file_number: Optional[int] = None


@dataclass
class Step1Stats:
    cleaned_paragraphs: int
    removed_citations: int
    removed_links: int
    removed_weird_numbers: int


@dataclass
class Step2Stats:
    paraphrased_paragraphs: int
    request_count: int
    total_words: int
    mode: str
    request_summaries: List[Dict[str, Any]]
    initial_chunk_failures: int
    recovery_attempts: int


@dataclass
class Step3Stats:
    detection_mode: str
    reference_count: int
    inserted_citations: int
    inserted_slide_paragraphs: int
    inserted_notes_paragraphs: int


def sanitize_text(text: str) -> str:
    return ZERO_WIDTH_RE.sub("", text or "").strip()


def sanitize_xml_text(text: str) -> str:
    return XML_INVALID_CHAR_RE.sub("", text or "")


def local_name(tag_name: str) -> str:
    if "}" in tag_name:
        return tag_name.rsplit("}", 1)[1]
    return tag_name


def register_metadata_namespaces() -> None:
    ET.register_namespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
    ET.register_namespace("dc", "http://purl.org/dc/elements/1.1/")
    ET.register_namespace("dcterms", "http://purl.org/dc/terms/")
    ET.register_namespace("dcmitype", "http://purl.org/dc/dcmitype/")
    ET.register_namespace("xsi", "http://www.w3.org/2001/XMLSchema-instance")
    ET.register_namespace("ep", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties")
    ET.register_namespace("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")
    ET.register_namespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main")


def inject_missing_namespace_declarations(xml_text: str, declarations: Dict[str, str]) -> str:
    if not declarations:
        return xml_text

    root_match = re.search(r"<([A-Za-z_][\w.:-]*)([^>]*)>", xml_text)
    if not root_match:
        return xml_text

    root_tag = root_match.group(0)
    attrs = root_match.group(2) or ""
    existing: Dict[Optional[str], str] = {}
    for match in XMLNS_DECLARATION_RE.finditer(attrs):
        prefix = match.group(1) if match.group(1) else None
        existing[prefix] = match.group(3)

    additions: List[str] = []
    for prefix, uri in declarations.items():
        if not prefix:
            continue
        if existing.get(prefix) == uri:
            continue
        if uri in existing.values():
            continue
        additions.append(f' xmlns:{prefix}="{uri}"')

    if not additions:
        return xml_text

    updated_root_tag = root_tag[:-1] + "".join(additions) + ">"
    return xml_text[: root_match.start()] + updated_root_tag + xml_text[root_match.end() :]


def normalize_dcterms_xsi_type_prefix(xml_text: str) -> str:
    qname_match = re.search(r"""xsi:type=(["'])dcterms:W3CDTF\1""", xml_text)
    if not qname_match or "xmlns:dcterms=" in xml_text:
        return xml_text

    root_match = re.search(r"<([A-Za-z_][\w.:-]*)([^>]*)>", xml_text)
    if not root_match:
        return xml_text

    attrs = root_match.group(2) or ""
    mapped_prefix: Optional[str] = None
    for match in XMLNS_DECLARATION_RE.finditer(attrs):
        prefix = match.group(1)
        uri = match.group(3)
        if prefix and uri == "http://purl.org/dc/terms/":
            mapped_prefix = prefix
            break

    if mapped_prefix:
        quote = qname_match.group(1)
        return re.sub(
            r"""xsi:type=(["'])dcterms:W3CDTF\1""",
            f'xsi:type={quote}{mapped_prefix}:W3CDTF{quote}',
            xml_text,
            count=1,
        )

    updated_root_tag = root_match.group(0)[:-1] + ' xmlns:dcterms="http://purl.org/dc/terms/">'
    injected = xml_text[: root_match.start()] + updated_root_tag + xml_text[root_match.end() :]
    return injected


def scrub_pptx_metadata_xml(path_name: str, xml_bytes: bytes) -> bytes:
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return xml_bytes

    changed = False

    if path_name == "docProps/core.xml":
        clear_fields = {
            "creator",
            "lastModifiedBy",
            "keywords",
            "description",
            "subject",
            "category",
            "contentStatus",
            "identifier",
            "language",
            "title",
        }
        for elem in root.iter():
            name = local_name(elem.tag)
            if name in clear_fields:
                if elem.text:
                    changed = True
                elem.text = ""
            elif name == "revision":
                if (elem.text or "") != "1":
                    changed = True
                elem.text = "1"
            elif name in {"created", "modified", "lastPrinted"}:
                if (elem.text or "") != METADATA_FIXED_TIMESTAMP:
                    changed = True
                elem.text = METADATA_FIXED_TIMESTAMP
                if name in {"created", "modified"}:
                    if elem.get(XSI_TYPE_ATTR) != "dcterms:W3CDTF":
                        changed = True
                    elem.set(XSI_TYPE_ATTR, "dcterms:W3CDTF")

    elif path_name == "docProps/app.xml":
        for elem in root.iter():
            name = local_name(elem.tag)
            if name in {"Company", "Manager", "LastAuthor", "HyperlinkBase", "Template"}:
                if elem.text:
                    changed = True
                elem.text = ""
            elif name == "TotalTime":
                if (elem.text or "") != "0":
                    changed = True
                elem.text = "0"

    elif path_name == "docProps/custom.xml":
        children = list(root)
        if children:
            changed = True
            for child in children:
                root.remove(child)

    elif path_name == "ppt/commentAuthors.xml":
        for elem in root.iter():
            for attr_name, attr_value in list(elem.attrib.items()):
                attr_local = local_name(attr_name)
                if attr_local in {"name", "initials"}:
                    if attr_value:
                        changed = True
                    elem.set(attr_name, "")

    if not changed:
        if path_name == "docProps/core.xml":
            normalized = normalize_dcterms_xsi_type_prefix(
                xml_bytes.decode("utf-8", errors="replace")
            )
            if normalized.encode("utf-8") != xml_bytes:
                return normalized.encode("utf-8")
        return xml_bytes

    register_metadata_namespaces()
    xml_text = ET.tostring(root, encoding="utf-8", xml_declaration=True).decode("utf-8")
    xml_text = normalize_dcterms_xsi_type_prefix(xml_text)
    return xml_text.encode("utf-8")


def normalize_space(text: str) -> str:
    return re.sub(r"\s{2,}", " ", text).strip()


def count_words(text: str) -> int:
    trimmed = sanitize_text(text)
    if not trimmed:
        return 0
    return len([token for token in trimmed.split() if token])


def paragraph_key(paragraph: PptParagraph) -> ParagraphKey:
    return (paragraph.archive_path, paragraph.paragraph_index)


def matches_reference_header(text: str) -> bool:
    trimmed = sanitize_text(text)
    if not trimmed:
        return False
    if any(pattern.match(trimmed) for pattern in REFERENCE_HEADER_PATTERNS):
        return True
    first_line = trimmed.splitlines()[0].strip() if "\n" in trimmed else trimmed
    return any(pattern.match(first_line) for pattern in REFERENCE_HEADER_PATTERNS)


def is_reference_like_line(raw_line: str) -> bool:
    line = normalize_space(raw_line)
    if not line:
        return False
    if count_words(line) < 4:
        return False

    has_year = bool(YEAR_RE.search(line))
    has_url_or_doi = bool(URL_OR_DOI_RE.search(line))
    has_cue = bool(REFERENCE_CUE_RE.search(line))
    has_author = bool(AUTHOR_RE.search(line))
    has_list_prefix = bool(LIST_PREFIX_RE.search(line))

    if has_url_or_doi and (has_year or has_author or has_list_prefix):
        return True
    if has_year and (has_author or has_cue or has_list_prefix):
        return True
    return False


def split_reference_candidate_lines(text: str) -> List[str]:
    if not text:
        return []
    prepared = text.replace("\r", "\n")
    prepared = re.sub(r"(?<!\d)(\d{1,3}[.)]\s*)", r"\n\1", prepared)
    lines = [sanitize_text(line) for line in re.split(r"\n+", prepared) if sanitize_text(line)]
    if lines:
        return lines
    return [sanitize_text(text)] if sanitize_text(text) else []


def count_reference_like_lines(text: str) -> int:
    lines = split_reference_candidate_lines(text)
    return sum(1 for line in lines if is_reference_like_line(line))


def infer_reference_file_number(paragraphs: Sequence[PptParagraph], kind: str) -> int:
    grouped: Dict[int, List[str]] = {}
    for paragraph in paragraphs:
        if paragraph.kind != kind:
            continue
        grouped.setdefault(paragraph.file_number, []).append(paragraph.text)

    best_match: Optional[Tuple[int, int]] = None
    for file_number, texts in grouped.items():
        reference_line_count = 0
        dense_paragraph_count = 0
        url_or_doi_paragraph_count = 0

        for text in texts:
            line_count = count_reference_like_lines(text)
            reference_line_count += line_count
            if line_count >= 2:
                dense_paragraph_count += 1
            if URL_OR_DOI_RE.search(text):
                url_or_doi_paragraph_count += 1

        looks_like_reference_block = (
            reference_line_count >= 4
            or dense_paragraph_count >= 2
            or (reference_line_count >= 3 and url_or_doi_paragraph_count >= 1)
        )
        if not looks_like_reference_block:
            continue

        score = reference_line_count * 3 + dense_paragraph_count * 5 + url_or_doi_paragraph_count * 2
        if (
            best_match is None
            or score > best_match[1]
            or (score == best_match[1] and file_number > best_match[0])
        ):
            best_match = (file_number, score)

    return best_match[0] if best_match else -1


def detect_reference_section(paragraphs: Sequence[PptParagraph]) -> Optional[ReferenceDetection]:
    for idx in range(len(paragraphs) - 1, -1, -1):
        paragraph = paragraphs[idx]
        if not matches_reference_header(paragraph.text):
            continue

        reference_keys: Set[ParagraphKey] = {paragraph_key(paragraph)}
        for candidate in paragraphs:
            if candidate.archive_path == paragraph.archive_path and candidate.paragraph_index < paragraph.paragraph_index:
                reference_keys.add(paragraph_key(candidate))
        for tail in paragraphs[idx + 1 :]:
            reference_keys.add(paragraph_key(tail))

        return ReferenceDetection(mode="header", reference_keys=reference_keys)

    for kind, mode_label in (("slide", "inferred-slide"), ("notes", "inferred-notes")):
        inferred_file_number = infer_reference_file_number(paragraphs, kind)
        if inferred_file_number == -1:
            continue

        reference_keys = {
            paragraph_key(paragraph)
            for paragraph in paragraphs
            if paragraph.kind == kind and paragraph.file_number == inferred_file_number
        }
        if reference_keys:
            return ReferenceDetection(
                mode=mode_label,
                reference_keys=reference_keys,
                inferred_kind=kind,
                inferred_file_number=inferred_file_number,
            )

    return None


def is_heading_or_subtitle(paragraph: PptParagraph) -> bool:
    text = paragraph.text
    if not text:
        return False

    if paragraph.kind == "slide" and paragraph.is_first_non_empty and paragraph.word_count <= 14:
        return True

    if text.endswith(":"):
        return True

    has_terminal_punctuation = bool(re.search(r"[.!?]$", text))
    is_shortish = 0 < paragraph.word_count <= 12
    words = [w for w in text.split() if w]
    if not words:
        return False

    capitalized = sum(1 for w in words if re.match(r"^[A-Z]", w))
    is_title_case = len(words) > 1 and (capitalized / len(words)) > 0.6

    return (not has_terminal_punctuation) and is_shortish and is_title_case


def split_sentences(text: str) -> List[str]:
    matches = re.findall(r"[^.!?]+(?:[.!?]+[\"')\]]*)?|[^.!?]+$", text)
    return matches if matches else [text]


def append_citation_at_sentence_end(sentence: str, citation: str) -> str:
    trimmed = sentence.rstrip()
    match = re.search(r"([.?!][\"')\]]*)$", trimmed)
    if not match:
        sep = "" if trimmed.endswith(" ") else " "
        return trimmed + sep + citation
    punctuation = match.group(1)
    core = trimmed[: -len(punctuation)]
    sep = "" if core.endswith(" ") else " "
    return f"{core}{sep}{citation}{punctuation}"


def select_sentence_index_for_citation(text: str) -> int:
    sentences = split_sentences(text)
    candidate_indexes: List[int] = []

    for i, sentence in enumerate(sentences):
        stripped = sentence.strip()
        if not stripped:
            continue
        if i == 0 and len(sentences) > 1:
            continue
        if count_words(stripped) < 8:
            continue
        if EXISTING_CITATION_RE.search(stripped):
            continue

        lower = stripped.lower()
        if (
            lower.startswith("in conclusion")
            or lower.startswith("to conclude")
            or lower.startswith("overall,")
            or lower.startswith("to sum up")
        ):
            continue

        candidate_indexes.append(i)

    if candidate_indexes:
        return candidate_indexes[-1]

    if not EXISTING_CITATION_RE.search(text) and count_words(text) >= 8:
        return max(0, len(sentences) - 1)

    return -1


def inject_citation_into_paragraph(text: str, citation: str) -> Tuple[str, bool]:
    sentences = split_sentences(text)
    sentence_index = select_sentence_index_for_citation(text)
    if sentence_index == -1:
        return text, False

    updated_sentences = list(sentences)
    updated_sentences[sentence_index] = append_citation_at_sentence_end(updated_sentences[sentence_index], citation)
    reconstructed = "".join(updated_sentences)
    return reconstructed, reconstructed != text


def remove_citations_links_and_weird_tokens(text: str) -> Tuple[str, int, int, int]:
    updated = text
    removed_citations = 0
    removed_links = 0
    removed_weird = 0

    for pattern in CITATION_PATTERNS:
        matches = list(pattern.finditer(updated))
        if matches:
            removed_citations += len(matches)
            updated = pattern.sub("", updated)

    for pattern in LINK_PATTERNS:
        matches = list(pattern.finditer(updated))
        if matches:
            removed_links += len(matches)
            updated = pattern.sub("", updated)

    weird_matches = list(WEIRD_NUMBER_PATTERN.finditer(updated))
    if weird_matches:
        removed_weird += len(weird_matches)
        updated = WEIRD_NUMBER_PATTERN.sub("", updated)

    updated = re.sub(r"\(\s*\)", "", updated)
    updated = re.sub(r"\s+\.", ".", updated)
    updated = re.sub(r"\s+([,;:])", r"\1", updated)
    updated = re.sub(r"\s{2,}", " ", updated).strip()

    return updated, removed_citations, removed_links, removed_weird


def clamp(value: float, min_value: float, max_value: float) -> float:
    return max(min_value, min(max_value, value))


def status_url_from_api(api_url: str) -> str:
    marker = "/paraphrase-batch"
    if api_url.endswith(marker):
        return api_url[: -len(marker)] + "/health"
    if api_url.endswith("/"):
        return api_url + "health"
    return api_url + "/health"


def fetch_health_snapshot(api_url: str, timeout_seconds: int) -> Optional[Dict[str, Any]]:
    url = status_url_from_api(api_url)
    request = urllib.request.Request(url, method="GET")
    timeout = max(2, min(timeout_seconds, 10))

    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            payload = response.read().decode("utf-8", errors="replace")
            if response.status < 200 or response.status >= 300:
                return None
            data = json.loads(payload)
            return data if isinstance(data, dict) else None
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError):
        return None


def get_available_accounts(snapshot: Optional[Dict[str, Any]]) -> List[str]:
    if not isinstance(snapshot, dict):
        return list(ACCOUNT_KEYS)

    ready_accounts: List[str] = []
    for account_key in ACCOUNT_KEYS:
        account_data = snapshot.get(account_key)
        status = ""
        if isinstance(account_data, dict):
            status = str(account_data.get("status", "")).lower()
        if not status or status in ("ready", "ok"):
            ready_accounts.append(account_key)

    if not ready_accounts:
        return [ACCOUNT_KEYS[0]]

    scheduler = snapshot.get("scheduler")
    scheduler_accounts = scheduler.get("accounts") if isinstance(scheduler, dict) else None
    if not isinstance(scheduler_accounts, dict):
        return ready_accounts

    non_tripped: List[str] = []
    for account_key in ready_accounts:
        account_stats = scheduler_accounts.get(account_key)
        health = str(account_stats.get("health", "")).lower() if isinstance(account_stats, dict) else ""
        if health != "tripped":
            non_tripped.append(account_key)

    if non_tripped:
        return non_tripped
    return [ready_accounts[0]]


def get_raw_budget(snapshot: Optional[Dict[str, Any]], account_key: str, mode: str) -> float:
    if not isinstance(snapshot, dict):
        return MODE_DEFAULT_BUDGET[mode]

    scheduler = snapshot.get("scheduler")
    if not isinstance(scheduler, dict):
        return MODE_DEFAULT_BUDGET[mode]

    budgets = scheduler.get("recommendedBudgets")
    if not isinstance(budgets, dict):
        return MODE_DEFAULT_BUDGET[mode]

    per_account = budgets.get("perAccount")
    if isinstance(per_account, dict):
        per_account_budget = per_account.get(account_key)
        if isinstance(per_account_budget, dict):
            value = per_account_budget.get(mode)
            if isinstance(value, (int, float)) and value > 0:
                return float(value)

    global_value = budgets.get(mode)
    if isinstance(global_value, (int, float)) and global_value > 0:
        return float(global_value)
    return MODE_DEFAULT_BUDGET[mode]


def get_reliability_factor(snapshot: Optional[Dict[str, Any]], account_key: str, mode: str) -> float:
    if not isinstance(snapshot, dict):
        return 1.0

    scheduler = snapshot.get("scheduler")
    if not isinstance(scheduler, dict):
        return 1.0

    accounts = scheduler.get("accounts")
    if not isinstance(accounts, dict):
        return 1.0

    account_stats = accounts.get(account_key)
    if not isinstance(account_stats, dict):
        return 1.0

    suffix = MODE_RATE_SUFFIX[mode]
    success_rate = account_stats.get(f"successRate{suffix}")
    retry_rate = account_stats.get(f"retryRate{suffix}")
    timeout_rate = account_stats.get(f"timeoutRate{suffix}")

    success = float(success_rate) if isinstance(success_rate, (int, float)) else 1.0
    retry = float(retry_rate) if isinstance(retry_rate, (int, float)) else 0.0
    timeout = float(timeout_rate) if isinstance(timeout_rate, (int, float)) else 0.0

    rate_factor = clamp(success - retry * 0.45 - timeout * 0.8, 0.4, 1.05)

    health = str(account_stats.get("health", "")).lower()
    health_factor = 1.0
    if health == "degraded":
        health_factor = 0.8
    elif health == "tripped":
        health_factor = 0.35

    return clamp(rate_factor * health_factor, 0.35, 1.05)


def get_system_penalty_seconds(snapshot: Optional[Dict[str, Any]], account_count: int) -> float:
    if account_count <= 1 or not isinstance(snapshot, dict):
        return 0.0

    scheduler = snapshot.get("scheduler")
    if not isinstance(scheduler, dict):
        return 0.0

    rolling = scheduler.get("rolling")
    if not isinstance(rolling, dict):
        return 0.0

    success_ratio_value = rolling.get("successRatio")
    fallback_rate_value = rolling.get("fallbackRate")
    success_ratio = float(success_ratio_value) if isinstance(success_ratio_value, (int, float)) else 1.0
    fallback_rate = float(fallback_rate_value) if isinstance(fallback_rate_value, (int, float)) else 0.0
    success_ratio = clamp(success_ratio, 0.0, 1.0)
    fallback_rate = clamp(fallback_rate, 0.0, 1.0)

    return (fallback_rate * 4.0 + (1.0 - success_ratio) * 6.0) * (account_count - 1)


def choose_account_plan(total_words: int, mode: str, snapshot: Optional[Dict[str, Any]]) -> Tuple[int, float, List[str], float]:
    if total_words <= 0:
        return 1, 0.0, [ACCOUNT_KEYS[0]], MODE_DEFAULT_BUDGET[mode]

    available_accounts = get_available_accounts(snapshot)
    profiles: List[Tuple[str, float]] = []
    for account_key in available_accounts:
        raw_budget = get_raw_budget(snapshot, account_key, mode)
        reliability = get_reliability_factor(snapshot, account_key, mode)
        effective_budget = max(120.0, raw_budget * reliability)
        profiles.append((account_key, effective_budget))

    profiles.sort(key=lambda item: item[1], reverse=True)
    if not profiles:
        return 1, (total_words / MODE_DEFAULT_BUDGET[mode]) * MODE_TARGET_SECONDS[mode], [ACCOUNT_KEYS[0]], MODE_DEFAULT_BUDGET[mode]

    best_count = 1
    best_estimated = float("inf")
    best_accounts = [profiles[0][0]]
    best_capacity = profiles[0][1]
    min_words_per_account = MODE_MIN_WORDS_PER_ACCOUNT[mode]
    max_count_by_words = max(1, int(total_words // min_words_per_account))
    max_candidate_count = min(len(profiles), max_count_by_words)

    for count in range(1, max_candidate_count + 1):
        selected = profiles[:count]
        capacity = max(100.0, sum(item[1] for item in selected))
        estimated = (total_words / capacity) * MODE_TARGET_SECONDS[mode]
        estimated += (count - 1) * MODE_COORDINATION_PENALTY_SECONDS[mode]
        estimated += get_system_penalty_seconds(snapshot, count)

        if count == 1:
            best_count = count
            best_estimated = estimated
            best_accounts = [item[0] for item in selected]
            best_capacity = capacity
            continue

        if estimated + 0.9 < best_estimated:
            best_count = count
            best_estimated = estimated
            best_accounts = [item[0] for item in selected]
            best_capacity = capacity

    return best_count, best_estimated, best_accounts, best_capacity


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


def build_payload_text(items: Sequence[PptParagraph]) -> str:
    chunks: List[str] = []
    for item in items:
        chunks.append(PARAPHRASE_DELIMITER)
        chunks.append(item.text)
    return "\n\n".join(chunks)


def split_into_account_chunks(
    items: Sequence[PptParagraph],
    selected_accounts: Sequence[str],
) -> List[Tuple[str, List[PptParagraph]]]:
    accounts = [key for key in selected_accounts if key in ACCOUNT_KEYS]
    if not accounts:
        accounts = [ACCOUNT_KEYS[0]]

    # Balance by words so one account does not get most of the heavy paragraphs.
    buckets: Dict[str, List[PptParagraph]] = {account_key: [] for account_key in accounts}
    bucket_words: Dict[str, int] = {account_key: 0 for account_key in accounts}

    for item in items:
        target = min(accounts, key=lambda key: (bucket_words[key], len(buckets[key])))
        buckets[target].append(item)
        bucket_words[target] += item.word_count

    chunks: List[Tuple[str, List[PptParagraph]]] = []
    for account_key in accounts:
        section = buckets[account_key]
        if section:
            chunks.append((account_key, section))
    return chunks


def take_request_batch(
    remaining: deque[PptParagraph],
    max_items_per_request: int,
    max_words_per_request: int,
) -> List[PptParagraph]:
    batch: List[PptParagraph] = []
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


def extract_account_output(
    response: Dict[str, object], account_key: str, mode: str, request_label: str
) -> Tuple[str, Optional[str]]:
    account_result = response.get(account_key)
    if not isinstance(account_result, dict):
        raise RuntimeError(f"Missing response for account {account_key} in {request_label}")

    paraphrased = account_result.get("secondMode") if mode == "dual" else account_result.get("result")
    if not paraphrased:
        error_message = account_result.get("error", "missing paraphrased output")
        raise RuntimeError(f"Account {account_key} failed in {request_label}: {error_message}")

    fallback_used = account_result.get("fallbackUsed")
    if fallback_used:
        print(f"[{request_label}] warning: {account_key} used fallback {fallback_used}")

    return str(paraphrased), (str(fallback_used) if fallback_used else None)


def recovery_account_order(preferred_account_key: str) -> List[str]:
    ordered = [preferred_account_key] if preferred_account_key in ACCOUNT_KEYS else []
    for key in ACCOUNT_KEYS:
        if key not in ordered:
            ordered.append(key)
    return ordered


def request_chunk_output_with_fallback(
    *,
    preferred_account_key: str,
    account_items: Sequence[PptParagraph],
    mode: str,
    api_url: str,
    timeout_seconds: int,
    request_label: str,
) -> str:
    payload_text = build_payload_text(account_items)
    errors: List[str] = []

    for key in recovery_account_order(preferred_account_key):
        try:
            payload = {"mode": mode, key: payload_text}
            response = post_batch_request(api_url, payload, timeout_seconds)
            output, _fallback = extract_account_output(response, key, mode, request_label)
            return output
        except Exception as error:  # pylint: disable=broad-except
            errors.append(f"{key}: {error}")
            print(f"[{request_label}] warning: recovery via {key} failed ({error})")

    raise RuntimeError(
        f"Recovery failed across all accounts for {request_label}: " + " | ".join(errors)
    )


def recover_account_chunk_parts(
    *,
    account_key: str,
    account_items: Sequence[PptParagraph],
    mode: str,
    api_url: str,
    timeout_seconds: int,
    request_label: str,
    depth: int = 0,
) -> List[str]:
    if not account_items:
        return []

    paraphrased_text = request_chunk_output_with_fallback(
        preferred_account_key=account_key,
        account_items=account_items,
        mode=mode,
        api_url=api_url,
        timeout_seconds=timeout_seconds,
        request_label=f"{request_label}:retry-d{depth}",
    )
    parts = parse_paraphrase_parts(paraphrased_text, len(account_items))

    if len(parts) == len(account_items):
        return parts

    if len(account_items) == 1:
        single = paraphrased_text.strip()
        if not single:
            raise RuntimeError(f"Recovery failed for single paragraph in {request_label}")
        return [single]

    if depth >= 6:
        recovered: List[str] = []
        for idx, item in enumerate(account_items):
            one = recover_account_chunk_parts(
                account_key=account_key,
                account_items=[item],
                mode=mode,
                api_url=api_url,
                timeout_seconds=timeout_seconds,
                request_label=f"{request_label}:single-{idx}",
                depth=depth + 1,
            )
            recovered.extend(one)
        return recovered

    midpoint = len(account_items) // 2
    if midpoint <= 0 or midpoint >= len(account_items):
        raise RuntimeError(
            f"Recovery split failed in {request_label}: expected {len(account_items)}, got {len(parts)}"
        )

    print(
        f"[{request_label}] warning: delimiter mismatch for {account_key} "
        f"(expected {len(account_items)}, got {len(parts)}); retrying in smaller batches"
    )
    left = recover_account_chunk_parts(
        account_key=account_key,
        account_items=account_items[:midpoint],
        mode=mode,
        api_url=api_url,
        timeout_seconds=timeout_seconds,
        request_label=f"{request_label}:left",
        depth=depth + 1,
    )
    right = recover_account_chunk_parts(
        account_key=account_key,
        account_items=account_items[midpoint:],
        mode=mode,
        api_url=api_url,
        timeout_seconds=timeout_seconds,
        request_label=f"{request_label}:right",
        depth=depth + 1,
    )
    return left + right


def extract_paragraphs_from_xml(root: ET.Element) -> List[ET.Element]:
    paragraphs = root.findall(".//p:txBody/a:p", NS)
    if paragraphs:
        return paragraphs
    return root.findall(".//a:txBody/a:p", NS)


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

    def sort_key(item: Tuple[str, str, int]) -> Tuple[int, int]:
        _, kind, number = item
        kind_rank = 0 if kind == "slide" else 1
        return number, kind_rank

    return sorted(targets, key=sort_key)


def collect_xml_docs_and_paragraphs(
    input_zip: zipfile.ZipFile,
    include_slides: bool,
    include_notes: bool,
) -> Tuple[Dict[str, ET.ElementTree], List[PptParagraph]]:
    xml_docs: Dict[str, ET.ElementTree] = {}
    paragraphs: List[PptParagraph] = []
    global_index = 0

    for archive_path, kind, file_number in sorted_target_xml_paths(input_zip.namelist(), include_slides, include_notes):
        root = ET.fromstring(input_zip.read(archive_path))
        tree = ET.ElementTree(root)
        xml_docs[archive_path] = tree

        paragraph_elements = extract_paragraphs_from_xml(root)
        first_non_empty_seen = False
        for paragraph_index, paragraph_element in enumerate(paragraph_elements):
            text_nodes = paragraph_element.findall(".//a:t", NS)
            if not text_nodes:
                continue

            text = sanitize_text("".join(node.text or "" for node in text_nodes))
            if not text:
                continue

            paragraph = PptParagraph(
                archive_path=archive_path,
                kind=kind,
                file_number=file_number,
                paragraph_index=paragraph_index,
                global_index=global_index,
                text_nodes=text_nodes,
                text=text,
                word_count=count_words(text),
                is_first_non_empty=not first_non_empty_seen,
            )
            paragraphs.append(paragraph)
            first_non_empty_seen = True
            global_index += 1

    return xml_docs, paragraphs


def set_paragraph_text(paragraph: PptParagraph, new_text: str) -> None:
    if not paragraph.text_nodes:
        return

    safe_text = sanitize_xml_text(new_text)
    paragraph.text_nodes[0].text = safe_text
    for node in paragraph.text_nodes[1:]:
        node.text = ""

    paragraph.text = sanitize_text(safe_text)
    paragraph.word_count = count_words(paragraph.text)


def run_step_1_clean(paragraphs: Sequence[PptParagraph], reference_keys: Set[ParagraphKey]) -> Step1Stats:
    cleaned_paragraphs = 0
    removed_citations = 0
    removed_links = 0
    removed_weird_numbers = 0

    for paragraph in paragraphs:
        if not paragraph.text:
            continue
        if paragraph_key(paragraph) in reference_keys:
            continue
        if is_heading_or_subtitle(paragraph):
            continue

        new_text, citation_count, link_count, weird_count = remove_citations_links_and_weird_tokens(paragraph.text)
        removed_citations += citation_count
        removed_links += link_count
        removed_weird_numbers += weird_count
        if new_text != paragraph.text:
            set_paragraph_text(paragraph, new_text)
            cleaned_paragraphs += 1

    return Step1Stats(
        cleaned_paragraphs=cleaned_paragraphs,
        removed_citations=removed_citations,
        removed_links=removed_links,
        removed_weird_numbers=removed_weird_numbers,
    )


def build_step_2_eligible_paragraphs(
    paragraphs: Sequence[PptParagraph],
    mode: str,
    reference_keys: Set[ParagraphKey],
) -> List[PptParagraph]:
    eligible: List[PptParagraph] = []
    min_words = 11 if mode == "dual" else 15

    for paragraph in paragraphs:
        if not paragraph.text:
            continue
        if paragraph_key(paragraph) in reference_keys:
            continue
        if is_heading_or_subtitle(paragraph):
            continue
        if paragraph.word_count < min_words:
            continue
        if paragraph.text.endswith(":"):
            continue
        eligible.append(paragraph)

    return eligible


def run_step_2_paraphrase(
    paragraphs: Sequence[PptParagraph],
    mode: str,
    reference_keys: Set[ParagraphKey],
    api_url: str,
    timeout_seconds: int,
    max_items_per_request: int,
    max_words_per_request: int,
) -> Step2Stats:
    eligible = build_step_2_eligible_paragraphs(paragraphs, mode, reference_keys)
    if not eligible:
        raise RuntimeError("No eligible slide/notes paragraphs found for paraphrasing.")

    remaining = deque(eligible)
    request_count = 0
    updated_count = 0
    total_words = sum(item.word_count for item in eligible)
    request_summaries: List[Dict[str, Any]] = []
    initial_chunk_failures = 0
    recovery_attempts = 0
    scheduler_snapshot = fetch_health_snapshot(api_url, timeout_seconds)

    while remaining:
        request_count += 1
        batch = take_request_batch(remaining, max_items_per_request, max_words_per_request)
        batch_word_count = sum(item.word_count for item in batch)
        _account_count, estimated_seconds, selected_accounts, effective_capacity = choose_account_plan(
            batch_word_count,
            mode,
            scheduler_snapshot,
        )
        trimmed_for_single_account = False

        # Guardrail: when only one account is currently usable, keep request size
        # within per-account bounds to avoid long click/retry failure storms.
        if len(selected_accounts) == 1 and len(batch) > 1:
            max_single_account_words = MODE_MAX_WORDS_PER_ACCOUNT.get(mode, max_words_per_request)
            if batch_word_count > max_single_account_words:
                trimmed_for_single_account = True
                original_count = len(batch)
                while len(batch) > 1 and batch_word_count > max_single_account_words:
                    moved = batch.pop()
                    batch_word_count -= moved.word_count
                    remaining.appendleft(moved)

                _account_count, estimated_seconds, selected_accounts, effective_capacity = choose_account_plan(
                    batch_word_count,
                    mode,
                    scheduler_snapshot,
                )
                print(
                    f"[2/4] request {request_count}: single-account guard trimmed "
                    f"{original_count}->{len(batch)} paragraphs ({batch_word_count} words)"
                )

        account_chunks = split_into_account_chunks(batch, selected_accounts)
        active_accounts = [account_key for account_key, _ in account_chunks]
        chunk_words_by_account = {
            account_key: sum(item.word_count for item in account_items)
            for account_key, account_items in account_chunks
        }
        request_summaries.append(
            {
                "request_index": request_count,
                "paragraphs": len(batch),
                "words": batch_word_count,
                "account_count": len(account_chunks),
                "accounts": active_accounts,
                "effective_capacity": effective_capacity,
                "estimated_seconds": estimated_seconds,
                "trimmed_for_single_account": trimmed_for_single_account,
                "chunk_words_by_account": chunk_words_by_account,
            }
        )

        payload: Dict[str, str] = {"mode": mode}
        for account_key, account_items in account_chunks:
            payload[account_key] = build_payload_text(account_items)

        print(
            f"[2/4] request {request_count}: {len(batch)} paragraphs, "
            f"{batch_word_count} words, {len(account_chunks)} account(s)"
            f" | est={estimated_seconds:.1f}s | accounts={'/'.join(active_accounts)} | capacity={effective_capacity:.0f}"
        )
        response = post_batch_request(api_url, payload, timeout_seconds)

        for account_key, account_items in account_chunks:
            request_label = f"request {request_count} {account_key}"
            try:
                paraphrased, _fallback_used = extract_account_output(
                    response, account_key, mode, f"request {request_count}"
                )
                parts = parse_paraphrase_parts(paraphrased, len(account_items))
            except Exception as error:  # pylint: disable=broad-except
                initial_chunk_failures += 1
                print(f"[{request_label}] warning: initial chunk failed ({error}); retrying with recovery")
                parts = []

            if len(parts) != len(account_items):
                recovery_attempts += 1
                parts = recover_account_chunk_parts(
                    account_key=account_key,
                    account_items=account_items,
                    mode=mode,
                    api_url=api_url,
                    timeout_seconds=timeout_seconds,
                    request_label=request_label,
                )

            if len(parts) != len(account_items):
                raise RuntimeError(
                    f"Response count mismatch in {request_label}: expected {len(account_items)}, "
                    f"got {len(parts)} after recovery"
                )

            for paragraph, new_text in zip(account_items, parts):
                set_paragraph_text(paragraph, new_text.strip())
                updated_count += 1

    return Step2Stats(
        paraphrased_paragraphs=updated_count,
        request_count=request_count,
        total_words=total_words,
        mode=mode,
        request_summaries=request_summaries,
        initial_chunk_failures=initial_chunk_failures,
        recovery_attempts=recovery_attempts,
    )


def extract_reference_entries(paragraphs: Sequence[PptParagraph], reference_keys: Set[ParagraphKey]) -> List[str]:
    section_paragraphs = [p for p in sorted(paragraphs, key=lambda p: p.global_index) if paragraph_key(p) in reference_keys and p.text]
    if not section_paragraphs:
        return []

    combined = "\n".join(paragraph.text for paragraph in section_paragraphs)
    lines = split_reference_candidate_lines(combined)

    entries: List[str] = []
    for line in lines:
        if matches_reference_header(line):
            continue
        if is_reference_like_line(line):
            entries.append(line)
        elif entries:
            entries[-1] = normalize_space(entries[-1] + " " + line)

    if not entries:
        for paragraph in section_paragraphs:
            if count_reference_like_lines(paragraph.text) > 0:
                entries.append(paragraph.text)

    deduped: List[str] = []
    seen = set()
    for entry in entries:
        cleaned = normalize_space(entry)
        if not cleaned:
            continue
        key = cleaned.lower()
        if key in seen:
            continue
        seen.add(key)
        deduped.append(cleaned)

    return deduped


def build_citation_from_reference(reference: str, index: int) -> str:
    cleaned = re.sub(r"^\s*(?:\d{1,3}[.)\]]|[-•])\s*", "", reference).strip(" .;:")
    if not cleaned:
        return f"(Source {index + 1})"

    year_match = YEAR_RE.search(cleaned)
    year = year_match.group(0) if year_match else None

    prefix = cleaned[: year_match.start()] if year_match else cleaned
    prefix = prefix.strip(" ,.;:-()[]")
    author = ""
    if "," in prefix:
        author = prefix.split(",", 1)[0].strip()
    if not author:
        words = [w.strip(" ,.;:-()[]") for w in prefix.split() if w.strip(" ,.;:-()[]")]
        author = " ".join(words[:4])

    if not author:
        author = f"Source {index + 1}"

    if year:
        return f"({author}, {year})"
    return f"({author})"


def build_step_3_eligible_paragraphs(
    paragraphs: Sequence[PptParagraph],
    reference_keys: Set[ParagraphKey],
) -> List[PptParagraph]:
    candidates: List[PptParagraph] = []
    for paragraph in paragraphs:
        if not paragraph.text:
            continue
        if paragraph_key(paragraph) in reference_keys:
            continue
        if is_heading_or_subtitle(paragraph):
            continue
        if paragraph.word_count < 11:
            continue
        if paragraph.text.endswith(":"):
            continue
        candidates.append(paragraph)
    return candidates


def run_step_3_add_references(paragraphs: Sequence[PptParagraph]) -> Step3Stats:
    detection = detect_reference_section(paragraphs)
    if detection is None:
        raise RuntimeError(
            "Could not detect a reference section in PPTX. "
            "No in-text references were added."
        )

    references = extract_reference_entries(paragraphs, detection.reference_keys)
    if not references:
        raise RuntimeError(
            "Reference section detected but no valid references could be parsed. "
            "No in-text references were added."
        )

    citations = [build_citation_from_reference(ref, idx) for idx, ref in enumerate(references)]
    if not citations:
        raise RuntimeError("Unable to build citation labels from detected references.")

    candidates = build_step_3_eligible_paragraphs(paragraphs, detection.reference_keys)
    if not candidates:
        raise RuntimeError("No eligible slide/notes paragraphs found for inserting references.")

    notes_candidates = [paragraph for paragraph in candidates if paragraph.kind == "notes"]
    target_count = min(len(candidates), max(1, len(citations)))
    inserted_total = 0
    inserted_slide = 0
    inserted_notes = 0
    citation_cursor = 0
    inserted_keys: Set[ParagraphKey] = set()

    # Ensure notes get at least one citation when eligible notes paragraphs exist.
    if notes_candidates:
        note_inserted = False
        for paragraph in notes_candidates:
            citation = citations[citation_cursor % len(citations)]
            updated_text, changed = inject_citation_into_paragraph(paragraph.text, citation)
            if not changed:
                continue
            set_paragraph_text(paragraph, updated_text)
            inserted_total += 1
            inserted_notes += 1
            citation_cursor += 1
            inserted_keys.add(paragraph_key(paragraph))
            note_inserted = True
            break

        if not note_inserted:
            raise RuntimeError(
                "Speaker notes were detected but no in-text citation could be inserted into eligible notes paragraphs."
            )

    for paragraph in candidates:
        if inserted_total >= target_count:
            break
        if paragraph_key(paragraph) in inserted_keys:
            continue

        citation = citations[citation_cursor % len(citations)]
        updated_text, changed = inject_citation_into_paragraph(paragraph.text, citation)
        if not changed:
            continue

        set_paragraph_text(paragraph, updated_text)
        inserted_total += 1
        citation_cursor += 1
        if paragraph.kind == "slide":
            inserted_slide += 1
        else:
            inserted_notes += 1

    if inserted_total == 0:
        raise RuntimeError(
            "Reference section was detected, but no citations could be inserted into slide/notes paragraphs."
        )
    if notes_candidates and inserted_notes == 0:
        raise RuntimeError(
            "Reference section was detected, but no citations were inserted into speaker notes paragraphs."
        )

    return Step3Stats(
        detection_mode=detection.mode,
        reference_count=len(references),
        inserted_citations=inserted_total,
        inserted_slide_paragraphs=inserted_slide,
        inserted_notes_paragraphs=inserted_notes,
    )


def write_output_pptx(input_path: Path, output_path: Path, xml_docs: Dict[str, ET.ElementTree]) -> None:
    with zipfile.ZipFile(input_path, "r") as source_zip, zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as out_zip:
        for info in source_zip.infolist():
            if info.filename in xml_docs:
                root = xml_docs[info.filename].getroot()
                xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                out_zip.writestr(info, xml_bytes)
            elif info.filename in PPTX_METADATA_XML_PATHS:
                original_bytes = source_zip.read(info.filename)
                scrubbed_bytes = scrub_pptx_metadata_xml(info.filename, original_bytes)
                out_zip.writestr(info, scrubbed_bytes)
            else:
                out_zip.writestr(info, source_zip.read(info.filename))


def parse_mode_and_input(
    positional: Sequence[str],
    explicit_mode: Optional[str],
) -> Tuple[str, Optional[Path]]:
    mode = explicit_mode or "dual"
    tokens = list(positional)

    if tokens:
        first = tokens[0].strip().lower()
        if first in {"standard", "std"}:
            mode = "standard"
            tokens = tokens[1:]
        elif first in {"simple", "simple+short", "simple-short", "dual", "fast"}:
            mode = "dual"
            tokens = tokens[1:]

    input_path: Optional[Path] = None
    if tokens:
        input_path = Path(tokens[0])
        tokens = tokens[1:]

    if tokens:
        raise RuntimeError(
            "Too many positional arguments. Use: `npm run ppt`, `npm run ppt standard`, "
            "or pass one input path."
        )

    return mode, input_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "PPTX one-shot pipeline: clean -> paraphrase -> add references -> output `pr <name>.pptx`."
        )
    )
    parser.add_argument(
        "positional",
        nargs="*",
        help="Optional mode/input. Examples: `standard`, `/path/file.pptx`, `standard /path/file.pptx`",
    )
    parser.add_argument(
        "--mode",
        choices=("dual", "standard"),
        help="Explicit paraphrase mode (overrides positional mode token).",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output PPTX path (default: `pr <input-name>.pptx`).",
    )
    parser.add_argument(
        "--api-url",
        default=API_URL_DEFAULT,
        help=f"Batch API URL (default: {API_URL_DEFAULT}).",
    )
    parser.add_argument(
        "--timeout-seconds",
        type=int,
        default=180,
        help="HTTP timeout for each paraphrase batch request.",
    )
    parser.add_argument(
        "--max-items-per-request",
        type=int,
        default=120,
        help="Max paragraphs sent per paraphrase API request.",
    )
    parser.add_argument(
        "--max-words-per-request",
        type=int,
        default=2400,
        help="Approximate max total words sent per paraphrase API request.",
    )
    parser.add_argument(
        "--no-slides",
        action="store_true",
        help="Do not process slide text.",
    )
    parser.add_argument(
        "--no-notes",
        action="store_true",
        help="Do not process speaker notes.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Run detection/eligibility checks without calling paraphrase API or writing output.",
    )
    return parser.parse_args()


def resolve_input_path(provided_input: Optional[Path]) -> Path:
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
    run_started_perf = time.perf_counter()
    run_started_local = local_now_pretty()
    run_id = f"pptx-{int(time.time())}"
    step_durations: Dict[str, float] = {}
    step1_payload: Dict[str, Any] = {}
    step2_payload: Dict[str, Any] = {}
    step3_payload: Dict[str, Any] = {}
    request_summaries: List[Dict[str, Any]] = []
    input_path: Optional[Path] = None
    output_path: Optional[Path] = None
    mode = "dual"
    error_step = ""
    error_message = ""
    error_traceback = ""
    paragraph_count = 0
    slide_count = 0
    notes_count = 0
    total_words = 0
    reference_detection_mode = "missing"
    reference_paragraph_count = 0

    def finalize(status: str, exit_code: int) -> int:
        total_duration = time.perf_counter() - run_started_perf
        step2_total_words = 0
        if isinstance(step2_payload, dict) and step2_payload:
            step2_total_words = int(step2_payload.get("total_words", 0) or 0)
            if not request_summaries and isinstance(step2_payload.get("request_summaries"), list):
                request_summaries.extend(step2_payload.get("request_summaries", []))
        append_pipeline_log(
            {
                "workflow": "pptx",
                "run_id": run_id,
                "status": status,
                "mode": mode,
                "started_at_local": run_started_local,
                "started_at_utc": utc_now_iso(),
                "dry_run": bool(args.dry_run),
                "input_path": str(input_path) if input_path else "",
                "output_path": str(output_path) if output_path else "",
                "include_slides": not args.no_slides,
                "include_notes": not args.no_notes,
                "paragraph_count": paragraph_count,
                "slide_paragraph_count": slide_count,
                "notes_paragraph_count": notes_count,
                "total_words_collected": total_words,
                "reference_detection_mode": reference_detection_mode,
                "reference_paragraph_count": reference_paragraph_count,
                "total_duration_seconds": total_duration,
                "step_durations": step_durations,
                "step1_stats": step1_payload,
                "step2_stats": step2_payload,
                "step3_stats": step3_payload,
                "step2_request_summaries": request_summaries,
                "step2_total_words": step2_total_words,
                "step2_initial_chunk_failures": int(step2_payload.get("initial_chunk_failures", 0) or 0),
                "step2_recovery_attempts": int(step2_payload.get("recovery_attempts", 0) or 0),
                "error_step": error_step,
                "error_message": error_message,
                "traceback": error_traceback,
            }
        )
        return exit_code

    try:
        mode, provided_input = parse_mode_and_input(args.positional, args.mode)
        input_path = resolve_input_path(provided_input)
    except RuntimeError as error:
        error_step = "bootstrap"
        error_message = str(error)
        print(error_message, file=sys.stderr)
        return finalize("failed", 1)

    if not input_path.exists():
        error_step = "bootstrap"
        error_message = f"Input file does not exist: {input_path}"
        print(error_message, file=sys.stderr)
        return finalize("failed", 1)
    if input_path.suffix.lower() != ".pptx":
        error_step = "bootstrap"
        error_message = f"Input must be a .pptx file: {input_path}"
        print(error_message, file=sys.stderr)
        return finalize("failed", 1)

    include_slides = not args.no_slides
    include_notes = not args.no_notes
    if not include_slides and not include_notes:
        error_step = "bootstrap"
        error_message = "Nothing to do: both slides and notes are disabled."
        print(error_message, file=sys.stderr)
        return finalize("failed", 1)

    output_path = args.output if args.output else input_path.with_name(f"pr {input_path.name}")

    try:
        with zipfile.ZipFile(input_path, "r") as input_zip:
            xml_docs, paragraphs = collect_xml_docs_and_paragraphs(
                input_zip=input_zip,
                include_slides=include_slides,
                include_notes=include_notes,
            )
    except Exception as error:  # pylint: disable=broad-except
        error_step = "bootstrap"
        error_message = f"Failed to read PPTX: {error}"
        error_traceback = traceback.format_exc()
        print(error_message, file=sys.stderr)
        return finalize("failed", 1)

    if not paragraphs:
        error_step = "bootstrap"
        error_message = "No eligible text containers found in the presentation XML."
        print(error_message, file=sys.stderr)
        return finalize("failed", 1)

    paragraph_count = len(paragraphs)
    slide_count = sum(1 for paragraph in paragraphs if paragraph.kind == "slide")
    notes_count = sum(1 for paragraph in paragraphs if paragraph.kind == "notes")
    total_words = sum(paragraph.word_count for paragraph in paragraphs)
    print(
        f"Collected paragraphs: total={len(paragraphs)}, slide={slide_count}, "
        f"notes={notes_count}, words={total_words}"
    )

    initial_detection = detect_reference_section(paragraphs)
    reference_keys = initial_detection.reference_keys if initial_detection else set()
    reference_paragraph_count = len(reference_keys)
    if initial_detection is None:
        reference_detection_mode = "missing"
        print(
            "Reference detection (initial): missing. Step 3 will fail-safe if references cannot be detected.",
            file=sys.stderr,
        )
    else:
        reference_detection_mode = initial_detection.mode
        print(
            f"Reference detection (initial): mode={initial_detection.mode}, "
            f"reference_paragraphs={len(initial_detection.reference_keys)}"
        )

    print("[1/4] Clean (weird numbers + links + in-text citations, keep headings/subtitles) ...", flush=True)
    try:
        step_started = time.perf_counter()
        step1 = run_step_1_clean(paragraphs, reference_keys)
        step_durations["step1_clean"] = round(time.perf_counter() - step_started, 3)
        step1_payload = asdict(step1)
        print(
            "[1/4] OK"
            f" | updated_paragraphs={step1.cleaned_paragraphs}"
            f", removed_citations={step1.removed_citations}"
            f", removed_links={step1.removed_links}"
            f", removed_weird_numbers={step1.removed_weird_numbers}"
        )
    except Exception as error:  # pylint: disable=broad-except
        error_step = "step1_clean"
        error_message = str(error)
        error_traceback = traceback.format_exc()
        print(f"[1/4] FAILED: {error}", file=sys.stderr)
        return finalize("failed", 1)

    # Fail fast before expensive paraphrase if we cannot reliably add references later.
    try:
        refreshed_detection = detect_reference_section(paragraphs)
        if refreshed_detection is None:
            raise RuntimeError("Could not detect a reference section for citation insertion.")
        preflight_references = extract_reference_entries(paragraphs, refreshed_detection.reference_keys)
        if not preflight_references:
            raise RuntimeError("Reference section found, but no usable references were parsed.")
        reference_keys = refreshed_detection.reference_keys
        reference_detection_mode = refreshed_detection.mode
        reference_paragraph_count = len(reference_keys)
        print(
            "[preflight] Reference section ready"
            f" | detection={reference_detection_mode}"
            f", reference_paragraphs={reference_paragraph_count}"
            f", parsed_references={len(preflight_references)}"
        )
    except Exception as error:  # pylint: disable=broad-except
        error_step = "step3_precheck"
        error_message = str(error)
        error_traceback = traceback.format_exc()
        print(
            "[preflight] FAILED: "
            f"{error}\n"
            "No output file was written. Review references formatting before sending this presentation.",
            file=sys.stderr,
        )
        return finalize("failed", 1)

    print(f"[2/4] Paraphrase ({'SIMPLE+SHORT' if mode == 'dual' else 'STANDARD'}) ...", flush=True)
    try:
        step_started = time.perf_counter()
        if args.dry_run:
            eligible = build_step_2_eligible_paragraphs(paragraphs, mode, reference_keys)
            if not eligible:
                raise RuntimeError("No eligible slide/notes paragraphs found for paraphrasing.")
            step2_payload = {
                "paraphrased_paragraphs": 0,
                "request_count": 0,
                "total_words": 0,
                "mode": mode,
                "request_summaries": [],
                "initial_chunk_failures": 0,
                "recovery_attempts": 0,
                "dry_run_eligible_paragraphs": len(eligible),
            }
            print(f"[2/4] DRY-RUN OK | eligible_paragraphs={len(eligible)}, mode={mode}")
        else:
            step2 = run_step_2_paraphrase(
                paragraphs=paragraphs,
                mode=mode,
                reference_keys=reference_keys,
                api_url=args.api_url,
                timeout_seconds=args.timeout_seconds,
                max_items_per_request=args.max_items_per_request,
                max_words_per_request=args.max_words_per_request,
            )
            step2_payload = asdict(step2)
            request_summaries = step2.request_summaries
            print(
                "[2/4] OK"
                f" | paraphrased_paragraphs={step2.paraphrased_paragraphs}"
                f", requests={step2.request_count}"
                f", words={step2.total_words}"
                f", mode={step2.mode}"
            )
        step_durations["step2_paraphrase"] = round(time.perf_counter() - step_started, 3)
    except Exception as error:  # pylint: disable=broad-except
        error_step = "step2_paraphrase"
        error_message = str(error)
        error_traceback = traceback.format_exc()
        print(f"[2/4] FAILED: {error}", file=sys.stderr)
        return finalize("failed", 1)

    print("[3/4] Add new in-text references (slides + speaker notes) ...", flush=True)
    try:
        step_started = time.perf_counter()
        if args.dry_run:
            detection = detect_reference_section(paragraphs)
            if detection is None:
                raise RuntimeError("Could not detect a reference section for citation insertion.")
            references = extract_reference_entries(paragraphs, detection.reference_keys)
            if not references:
                raise RuntimeError("Reference section found, but no usable references were parsed.")
            candidates = build_step_3_eligible_paragraphs(paragraphs, detection.reference_keys)
            if not candidates:
                raise RuntimeError("No eligible slide/notes paragraphs found for inserting references.")
            step3_payload = {
                "detection_mode": detection.mode,
                "reference_count": len(references),
                "inserted_citations": 0,
                "inserted_slide_paragraphs": 0,
                "inserted_notes_paragraphs": 0,
                "dry_run_eligible_targets": len(candidates),
            }
            print(
                "[3/4] DRY-RUN OK"
                f" | detection={detection.mode}"
                f", references={len(references)}"
                f", eligible_targets={len(candidates)}"
            )
        else:
            step3 = run_step_3_add_references(paragraphs)
            step3_payload = asdict(step3)
            print(
                "[3/4] OK"
                f" | detection={step3.detection_mode}"
                f", parsed_references={step3.reference_count}"
                f", inserted_citations={step3.inserted_citations}"
                f", inserted_slide={step3.inserted_slide_paragraphs}"
                f", inserted_notes={step3.inserted_notes_paragraphs}"
            )
        step_durations["step3_references"] = round(time.perf_counter() - step_started, 3)
    except Exception as error:  # pylint: disable=broad-except
        error_step = "step3_references"
        error_message = str(error)
        error_traceback = traceback.format_exc()
        print(
            "[3/4] FAILED: "
            f"{error}\n"
            "No output file was written. Review references formatting before sending this presentation.",
            file=sys.stderr,
        )
        return finalize("failed", 1)

    print("[4/4] Write output PPTX ...", flush=True)
    try:
        step_started = time.perf_counter()
        if args.dry_run:
            print("[4/4] DRY-RUN OK | no file written")
        else:
            write_output_pptx(input_path, output_path, xml_docs)
            print(f"[4/4] OK | output={output_path}")
        step_durations["step4_write_output"] = round(time.perf_counter() - step_started, 3)
    except Exception as error:  # pylint: disable=broad-except
        error_step = "step4_write_output"
        error_message = str(error)
        error_traceback = traceback.format_exc()
        print(f"[4/4] FAILED: {error}", file=sys.stderr)
        return finalize("failed", 1)

    if args.dry_run:
        print("Pipeline dry-run completed successfully.")
    else:
        print("Pipeline completed successfully.")
    return finalize("success", 0)


if __name__ == "__main__":
    raise SystemExit(main())
