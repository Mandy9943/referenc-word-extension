#!/usr/bin/env python3
"""
One-shot DOCX pipeline for:
1) cleaning in-text citations + weird artifacts (while preserving headings/subtitles),
2) paraphrasing body paragraphs via batch API,
3) inserting refreshed in-text references,
4) writing output as: pr <original-name>.docx
"""

from __future__ import annotations

import argparse
import json
import math
import random
import re
import sys
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

API_URL_DEFAULT = "https://analizeai.com/paraphrase-batch"
PARAPHRASE_DELIMITER = "qbpdelim123"
ACCOUNT_KEYS = ("acc1", "acc2", "acc3")
MODE_DEFAULT_BUDGET = {"dual": 520.0, "standard": 950.0, "ludicrous": 600.0}
MODE_TARGET_SECONDS = {"dual": 18.0, "standard": 9.0, "ludicrous": 16.0}
MODE_MIN_WORDS_PER_ACCOUNT = {"dual": 280.0, "standard": 500.0, "ludicrous": 300.0}
MODE_MAX_WORDS_PER_ACCOUNT = {"dual": 760, "standard": 1400, "ludicrous": 900}
MODE_COORDINATION_PENALTY_SECONDS = {"dual": 1.2, "standard": 0.7, "ludicrous": 1.0}
MODE_RATE_SUFFIX = {"dual": "Dual", "standard": "Standard", "ludicrous": "Ludicrous"}

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
XML_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"
XSI_TYPE_ATTR = "{http://www.w3.org/2001/XMLSchema-instance}type"
METADATA_FIXED_TIMESTAMP = "2000-01-01T00:00:00Z"
DOCX_METADATA_XML_PATHS = {
    "docProps/core.xml",
    "docProps/app.xml",
    "docProps/custom.xml",
    "word/comments.xml",
    "word/people.xml",
}

ZERO_WIDTH_RE = re.compile(r"[\u200B-\u200D\uFEFF]")
XML_INVALID_CHAR_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
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
    re.compile(
        rf"^\s*{NUMBERING_PREFIX}literature\s+cited\s*{TRAILING_PUNCTUATION}\s*$",
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
LIST_PREFIX_RE = re.compile(r"^\s*(?:\[\d{1,3}\]|\d{1,3}[.)\]]|[-•])\s+")

TOC_HEADER_PATTERNS = [
    re.compile(rf"^\s*{NUMBERING_PREFIX}table\s+of\s+contents?\s*{TRAILING_PUNCTUATION}\s*$", re.IGNORECASE),
    re.compile(rf"^\s*{NUMBERING_PREFIX}contents?\s*{TRAILING_PUNCTUATION}\s*$", re.IGNORECASE),
    re.compile(rf"^\s*{NUMBERING_PREFIX}toc\s*{TRAILING_PUNCTUATION}\s*$", re.IGNORECASE),
]

CONCLUSION_HEADER_PATTERNS = [
    re.compile(rf"^\s*{NUMBERING_PREFIX}conclusions?(?:\s+section)?\s*{TRAILING_PUNCTUATION}\s*$", re.IGNORECASE),
    re.compile(rf"^\s*{NUMBERING_PREFIX}concluding\s+remarks\s*{TRAILING_PUNCTUATION}\s*$", re.IGNORECASE),
    re.compile(rf"^\s*{NUMBERING_PREFIX}final\s+thoughts\s*{TRAILING_PUNCTUATION}\s*$", re.IGNORECASE),
    re.compile(rf"^\s*{NUMBERING_PREFIX}summary(?:\s+and\s+future\s+work)?\s*{TRAILING_PUNCTUATION}\s*$", re.IGNORECASE),
    re.compile(rf"^\s*{NUMBERING_PREFIX}closing\s+remarks\s*{TRAILING_PUNCTUATION}\s*$", re.IGNORECASE),
    re.compile(
        rf"^\s*{NUMBERING_PREFIX}conclusions?\s+and\s+recommendations\s*{TRAILING_PUNCTUATION}\s*$",
        re.IGNORECASE,
    ),
]

CITATION_PATTERNS = [
    re.compile(r"\[(?:[^\]]+)[,\s]\s?\d{4}[a-z]?\]"),
    re.compile(r"\((?:[^,()]+(,\s[^,()]+)*(?:,\sand\s[^,()]+)?)[,\s]\s?\d{4}[a-z]?\)"),
    re.compile(r"\((?:[^,()]+)[,\s]\s?\d{4}[a-z]?\)"),
    re.compile(r"\((?:[^()]+\sand\s[^,()]+)[,\s]\s?\d{4}[a-z]?\)"),
    re.compile(r"\((?:[^()]+)\set\sal\.?[,\s]\s?\d{4}[a-z]?\)"),
    re.compile(r"\((?:[^,()]+(,\s[^,()]+)*)[,\s]\s?\d{4}[a-z]?\)"),
]
WEIRD_NUMBER_PATTERNS = [
    re.compile(r"[【\[]\s*\d{9,}[^\]】]*?[†‡]?\s*[Ll]\d{1,4}(?:\s*[-–—]\s*[Ll]?\d{1,4})?[^\]】]*?[】\]]"),
    re.compile(r"[【\[]\s*\d{9,}\s*[】\]]"),
]
EXISTING_CITATION_RE = re.compile(r"\(\s*[^)]*?\d{4}[a-z]?\s*\)")
XMLNS_DECLARATION_RE = re.compile(r"""xmlns(?::([A-Za-z_][\w.\-]*))?\s*=\s*(['"])(.*?)\2""")


@dataclass
class Paragraph:
    index: int
    element: ET.Element
    text_nodes: List[ET.Element]
    text: str
    style_id: str
    align: str
    word_count: int


@dataclass
class ParaphraseItem:
    paragraph_index: int
    text: str
    word_count: int


@dataclass
class Step1Stats:
    cleaned_paragraphs: int
    removed_citations: int
    removed_weird_numbers: int


@dataclass
class Step2Stats:
    paraphrased_paragraphs: int
    request_count: int
    total_words: int
    mode: str


@dataclass
class Step3Stats:
    detection_mode: str
    reference_start_index: int
    reference_count: int
    inserted_citations: int


def qn(tag: str) -> str:
    prefix, local = tag.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


def sanitize_text(text: str) -> str:
    return ZERO_WIDTH_RE.sub("", text or "").strip()


def sanitize_xml_text(text: str) -> str:
    # Word can report "unreadable content" when control chars leak into document.xml.
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
    ET.register_namespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    ET.register_namespace("w15", "http://schemas.microsoft.com/office/word/2012/wordml")


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


def scrub_docx_metadata_xml(path_name: str, xml_bytes: bytes) -> bytes:
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

    elif path_name == "word/comments.xml":
        for elem in root.iter():
            for attr_name, attr_value in list(elem.attrib.items()):
                attr_local = local_name(attr_name)
                if attr_local in {"author", "initials"}:
                    if attr_value:
                        changed = True
                    elem.set(attr_name, "")
                elif attr_local == "date":
                    if attr_value != METADATA_FIXED_TIMESTAMP:
                        changed = True
                    elem.set(attr_name, METADATA_FIXED_TIMESTAMP)

    elif path_name == "word/people.xml":
        scrub_attrs = {"author", "name", "initials", "presenceInfo", "providerId", "userId"}
        for elem in root.iter():
            if elem.text and elem.text.strip():
                changed = True
                elem.text = ""
            for attr_name, attr_value in list(elem.attrib.items()):
                attr_local = local_name(attr_name)
                if attr_local in scrub_attrs:
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


def looks_like_toc_line(text: str) -> bool:
    trimmed = text.strip()
    if not trimmed:
        return False
    if re.search(r"\.{5,}.*\d+\s*$", trimmed):
        return True
    if re.search(r"\.{3,}\s*\d+\s*$", trimmed):
        return True
    if "\t" in trimmed:
        return True
    if re.search(r"\s+\.{2,}\s*\d+\s*$", trimmed):
        return True
    return False


def matches_reference_header(text: str) -> bool:
    trimmed = sanitize_text(text)
    if not trimmed:
        return False
    if any(pattern.match(trimmed) for pattern in REFERENCE_HEADER_PATTERNS):
        return True
    first_line = trimmed.splitlines()[0].strip() if "\n" in trimmed else trimmed
    return any(pattern.match(first_line) for pattern in REFERENCE_HEADER_PATTERNS)


def matches_conclusion_header(text: str) -> bool:
    trimmed = sanitize_text(text)
    if not trimmed:
        return False
    if any(pattern.match(trimmed) for pattern in CONCLUSION_HEADER_PATTERNS):
        return True
    first_line = trimmed.splitlines()[0].strip() if "\n" in trimmed else trimmed
    return any(pattern.match(first_line) for pattern in CONCLUSION_HEADER_PATTERNS)


def is_heading_or_subtitle(paragraph: Paragraph) -> bool:
    style = paragraph.style_id.lower()
    if style:
        if re.search(r"(^|[-_])heading[1-9]?$", style):
            return True
        if "title" in style or "subtitle" in style:
            return True
        if style == "tocheading":
            return True
        if style.startswith("heading"):
            return True

    text = paragraph.text
    if not text:
        return False

    has_terminal_punctuation = bool(re.search(r"[.!?]$", text))
    is_shortish = 0 < paragraph.word_count <= 15
    is_centered = paragraph.align in {"center", "centered"}

    words = [w for w in text.split() if w]
    if not words:
        return False
    capitalized = sum(1 for w in words if re.match(r"^[A-Z]", w))
    is_title_case = len(words) > 1 and (capitalized / len(words)) > 0.6

    return (not has_terminal_punctuation) and is_shortish and (is_centered or is_title_case)


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


def infer_reference_start_index(paragraphs: Sequence[Paragraph]) -> int:
    if not paragraphs:
        return -1

    n = len(paragraphs)
    tail_start = max(0, int(n * 0.45))
    scores = [count_reference_like_lines(p.text) for p in paragraphs]

    # Single dense paragraph with many references (common when references are pasted as one block).
    dense_candidates = [i for i in range(tail_start, n) if scores[i] >= 4]
    if dense_candidates:
        return dense_candidates[0]

    scored_indices = [i for i in range(tail_start, n) if scores[i] >= 1]
    if len(scored_indices) < 3:
        return -1

    clusters: List[List[int]] = []
    cluster = [scored_indices[0]]
    for idx in scored_indices[1:]:
        if idx - cluster[-1] <= 2:
            cluster.append(idx)
        else:
            clusters.append(cluster)
            cluster = [idx]
    clusters.append(cluster)

    best_cluster: Optional[List[int]] = None
    best_score = -1
    for c in clusters:
        score = sum(scores[i] for i in c)
        if score > best_score:
            best_score = score
            best_cluster = c
        elif score == best_score and best_cluster and c[-1] > best_cluster[-1]:
            best_cluster = c

    if not best_cluster:
        return -1

    if best_score >= 4 or len(best_cluster) >= 3:
        return best_cluster[0]

    return -1


def detect_reference_section(paragraphs: Sequence[Paragraph]) -> Tuple[int, str]:
    for i in range(len(paragraphs) - 1, -1, -1):
        if matches_reference_header(paragraphs[i].text):
            return i, "header"

    inferred = infer_reference_start_index(paragraphs)
    if inferred != -1:
        return inferred, "inferred"

    return -1, "missing"


def find_conclusion_range(paragraphs: Sequence[Paragraph], reference_start_index: int) -> Tuple[int, int]:
    search_end = reference_start_index if reference_start_index != -1 else len(paragraphs)
    conclusion_heading_index = -1

    for i in range(search_end - 1, -1, -1):
        if matches_conclusion_header(paragraphs[i].text):
            conclusion_heading_index = i
            break

    if conclusion_heading_index == -1:
        return -1, -1

    conclusion_end_index = search_end
    for i in range(conclusion_heading_index + 1, search_end):
        if matches_conclusion_header(paragraphs[i].text):
            continue
        if is_heading_or_subtitle(paragraphs[i]):
            conclusion_end_index = i
            break

    return conclusion_heading_index, conclusion_end_index


def extract_namespace_declarations(xml_text: str) -> Dict[str, str]:
    root_match = re.search(r"<([A-Za-z_][\w:.\-]*)([^>]*)>", xml_text)
    if not root_match:
        return {}

    attrs = root_match.group(2)
    declarations: Dict[str, str] = {}
    for match in XMLNS_DECLARATION_RE.finditer(attrs):
        prefix = match.group(1) or ""
        uri = match.group(3)
        declarations[prefix] = uri
    return declarations


def inject_missing_namespace_declarations(xml_text: str, declarations: Dict[str, str]) -> str:
    if not declarations:
        return xml_text

    root_match = re.search(r"<([A-Za-z_][\w:.\-]*)([^>]*)>", xml_text)
    if not root_match:
        return xml_text

    root_tag = root_match.group(0)
    additions: List[str] = []
    for prefix, uri in declarations.items():
        if prefix:
            attr_pattern = rf"""\sxmlns:{re.escape(prefix)}\s*="""
        else:
            attr_pattern = r"""\sxmlns\s*="""
        if re.search(attr_pattern, root_tag):
            continue
        if prefix:
            additions.append(f' xmlns:{prefix}="{uri}"')
        else:
            additions.append(f' xmlns="{uri}"')

    if not additions:
        return xml_text

    updated_root_tag = root_tag[:-1] + "".join(additions) + ">"
    return xml_text[: root_match.start()] + updated_root_tag + xml_text[root_match.end() :]


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


def build_batch_payload_text(items: Sequence[ParaphraseItem]) -> str:
    chunks: List[str] = []
    for item in items:
        chunks.append(PARAPHRASE_DELIMITER)
        chunks.append(item.text)
    return "\n\n".join(chunks)


def split_into_account_chunks(
    items: Sequence[ParaphraseItem],
    selected_accounts: Sequence[str],
) -> List[Tuple[str, List[ParaphraseItem]]]:
    accounts = [key for key in selected_accounts if key in ACCOUNT_KEYS]
    if not accounts:
        accounts = [ACCOUNT_KEYS[0]]

    # Balance by words (not item count) so one account does not get a huge
    # paragraph while others stay nearly idle.
    buckets: Dict[str, List[ParaphraseItem]] = {account_key: [] for account_key in accounts}
    bucket_words: Dict[str, int] = {account_key: 0 for account_key in accounts}

    for item in items:
        target = min(accounts, key=lambda key: (bucket_words[key], len(buckets[key])))
        buckets[target].append(item)
        bucket_words[target] += item.word_count

    chunks: List[Tuple[str, List[ParaphraseItem]]] = []
    for account_key in accounts:
        section = buckets[account_key]
        if section:
            chunks.append((account_key, section))
    return chunks


def take_request_batch(
    items: Sequence[ParaphraseItem],
    start_index: int,
    max_items_per_request: int,
    max_words_per_request: int,
) -> Tuple[List[ParaphraseItem], int, int]:
    batch: List[ParaphraseItem] = []
    total_words = 0
    idx = start_index

    while idx < len(items) and len(batch) < max_items_per_request:
        next_item = items[idx]
        if batch and (total_words + next_item.word_count > max_words_per_request):
            break
        batch.append(next_item)
        total_words += next_item.word_count
        idx += 1

    if not batch and idx < len(items):
        batch.append(items[idx])
        total_words = items[idx].word_count
        idx += 1

    return batch, idx, total_words


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
    response: Dict[str, object],
    account_key: str,
    mode: str,
    request_label: str,
) -> str:
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

    return str(paraphrased)


def recovery_account_order(preferred_account_key: str) -> List[str]:
    ordered = [preferred_account_key] if preferred_account_key in ACCOUNT_KEYS else []
    for key in ACCOUNT_KEYS:
        if key not in ordered:
            ordered.append(key)
    return ordered


def request_chunk_output_with_fallback(
    *,
    preferred_account_key: str,
    account_items: Sequence[ParaphraseItem],
    mode: str,
    api_url: str,
    timeout_seconds: int,
    request_label: str,
) -> str:
    payload_text = build_batch_payload_text(account_items)
    errors: List[str] = []

    for key in recovery_account_order(preferred_account_key):
        try:
            payload = {"mode": mode, key: payload_text}
            response = post_batch_request(api_url, payload, timeout_seconds)
            return extract_account_output(response, key, mode, request_label)
        except Exception as error:  # pylint: disable=broad-except
            errors.append(f"{key}: {error}")
            print(f"[{request_label}] warning: recovery via {key} failed ({error})")

    raise RuntimeError(
        f"Recovery failed across all accounts for {request_label}: " + " | ".join(errors)
    )


def recover_account_chunk_parts(
    *,
    account_key: str,
    account_items: Sequence[ParaphraseItem],
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


def get_paragraph_style_id(paragraph: ET.Element) -> str:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        return ""
    style = ppr.find("w:pStyle", NS)
    if style is None:
        return ""
    return (style.get(qn("w:val")) or "").strip()


def get_paragraph_alignment(paragraph: ET.Element) -> str:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        return ""
    jc = ppr.find("w:jc", NS)
    if jc is None:
        return ""
    return (jc.get(qn("w:val")) or "").strip().lower()


def extract_paragraph_text_nodes(paragraph: ET.Element) -> List[ET.Element]:
    return paragraph.findall(".//w:t", NS)


def paragraph_text_from_nodes(text_nodes: Sequence[ET.Element]) -> str:
    return sanitize_text("".join(node.text or "" for node in text_nodes))


def collect_paragraphs(document_root: ET.Element) -> List[Paragraph]:
    body = document_root.find("w:body", NS)
    if body is None:
        return []

    paragraph_elements = body.findall(".//w:p", NS)
    paragraphs: List[Paragraph] = []
    for index, element in enumerate(paragraph_elements):
        text_nodes = extract_paragraph_text_nodes(element)
        text = paragraph_text_from_nodes(text_nodes)
        paragraphs.append(
            Paragraph(
                index=index,
                element=element,
                text_nodes=text_nodes,
                text=text,
                style_id=get_paragraph_style_id(element),
                align=get_paragraph_alignment(element),
                word_count=count_words(text),
            )
        )
    return paragraphs


def set_text_nodes_value(text_nodes: Sequence[ET.Element], text_value: str) -> None:
    if not text_nodes:
        return

    text_nodes[0].text = text_value
    if text_value.startswith(" ") or text_value.endswith(" "):
        text_nodes[0].set(XML_SPACE_ATTR, "preserve")
    elif XML_SPACE_ATTR in text_nodes[0].attrib:
        del text_nodes[0].attrib[XML_SPACE_ATTR]

    for node in text_nodes[1:]:
        node.text = ""
        if XML_SPACE_ATTR in node.attrib:
            del node.attrib[XML_SPACE_ATTR]


def collect_text_nodes_by_break(paragraph_element: ET.Element) -> List[List[ET.Element]]:
    segments: List[List[ET.Element]] = [[]]
    br_tag = qn("w:br")
    cr_tag = qn("w:cr")
    t_tag = qn("w:t")

    for elem in paragraph_element.iter():
        if elem.tag == br_tag or elem.tag == cr_tag:
            segments.append([])
            continue
        if elem.tag == t_tag:
            segments[-1].append(elem)

    return [segment for segment in segments if segment]


def split_text_by_target_word_counts(text: str, target_word_counts: Sequence[int]) -> List[str]:
    segment_count = len(target_word_counts)
    if segment_count == 0:
        return []
    if segment_count == 1:
        return [text]

    words = [word for word in text.split() if word]
    if not words:
        return [""] * segment_count

    total_words = len(words)
    weights = [max(1, int(count)) for count in target_word_counts]
    total_weight = sum(weights)

    boundaries: List[int] = []
    cumulative_weight = 0
    previous = 0
    for idx in range(segment_count - 1):
        cumulative_weight += weights[idx]
        suggested = round(total_words * cumulative_weight / total_weight)
        min_allowed = previous + 1
        max_allowed = total_words - (segment_count - idx - 1)
        boundary = max(min_allowed, min(suggested, max_allowed))
        boundaries.append(boundary)
        previous = boundary

    parts: List[str] = []
    start = 0
    for boundary in boundaries:
        parts.append(" ".join(words[start:boundary]).strip())
        start = boundary
    parts.append(" ".join(words[start:]).strip())
    return parts


def set_paragraph_text(paragraph: Paragraph, new_text: str) -> None:
    safe_text = sanitize_xml_text(new_text)
    text_nodes = paragraph.text_nodes
    if not text_nodes:
        run = ET.Element(qn("w:r"))
        node = ET.Element(qn("w:t"))
        run.append(node)
        paragraph.element.append(run)
        text_nodes = [node]
        paragraph.text_nodes = text_nodes

    nodes_by_break_segment = collect_text_nodes_by_break(paragraph.element)
    if len(nodes_by_break_segment) > 1:
        original_segment_word_counts = [
            count_words("".join(node.text or "" for node in segment_nodes))
            for segment_nodes in nodes_by_break_segment
        ]
        segment_texts = split_text_by_target_word_counts(safe_text, original_segment_word_counts)
        for segment_nodes, segment_text in zip(nodes_by_break_segment, segment_texts):
            set_text_nodes_value(segment_nodes, segment_text)
    else:
        set_text_nodes_value(text_nodes, safe_text)

    paragraph.text = sanitize_text(safe_text)
    paragraph.word_count = count_words(paragraph.text)


def remove_citations_and_weird_tokens(text: str) -> Tuple[str, int, int]:
    updated = text
    removed_citations = 0
    removed_weird = 0

    for pattern in CITATION_PATTERNS:
        matches = list(pattern.finditer(updated))
        if matches:
            removed_citations += len(matches)
            updated = pattern.sub("", updated)

    for pattern in WEIRD_NUMBER_PATTERNS:
        weird_matches = list(pattern.finditer(updated))
        if weird_matches:
            removed_weird += len(weird_matches)
            updated = pattern.sub("", updated)

    updated = re.sub(r"\s+\.", ".", updated)
    updated = re.sub(r"\s+([,;:])", r"\1", updated)
    updated = re.sub(r"\s{2,}", " ", updated).strip()

    return updated, removed_citations, removed_weird


def run_step_1_clean(paragraphs: Sequence[Paragraph], reference_start_index: int) -> Step1Stats:
    cleaned_paragraphs = 0
    removed_citations = 0
    removed_weird_numbers = 0

    for p in paragraphs:
        if not p.text:
            continue
        if reference_start_index != -1 and p.index >= reference_start_index:
            continue
        if any(pattern.match(p.text) for pattern in TOC_HEADER_PATTERNS):
            continue
        if looks_like_toc_line(p.text):
            continue
        if is_heading_or_subtitle(p):
            continue

        new_text, citations_count, weird_count = remove_citations_and_weird_tokens(p.text)
        removed_citations += citations_count
        removed_weird_numbers += weird_count
        if new_text != p.text:
            set_paragraph_text(p, new_text)
            cleaned_paragraphs += 1

    return Step1Stats(
        cleaned_paragraphs=cleaned_paragraphs,
        removed_citations=removed_citations,
        removed_weird_numbers=removed_weird_numbers,
    )


def run_step_2_paraphrase(
    paragraphs: Sequence[Paragraph],
    mode: str,
    api_url: str,
    timeout_seconds: int,
    max_items_per_request: int,
    max_words_per_request: int,
    reference_start_index: int,
) -> Step2Stats:
    items: List[ParaphraseItem] = []
    for p in paragraphs:
        if not p.text:
            continue
        if reference_start_index != -1 and p.index >= reference_start_index:
            continue
        if is_heading_or_subtitle(p):
            continue
        if looks_like_toc_line(p.text):
            continue
        if p.word_count < 15:
            continue
        items.append(ParaphraseItem(paragraph_index=p.index, text=p.text, word_count=p.word_count))

    if not items:
        raise RuntimeError("No eligible body paragraphs found for paraphrasing.")

    request_count = 0
    updated_count = 0
    cursor = 0
    total_words = sum(item.word_count for item in items)
    paragraph_by_index = {p.index: p for p in paragraphs}
    scheduler_snapshot: Optional[Dict[str, Any]] = None

    while cursor < len(items):
        request_count += 1
        latest_snapshot = fetch_health_snapshot(api_url, timeout_seconds)
        if isinstance(latest_snapshot, dict):
            scheduler_snapshot = latest_snapshot

        batch, cursor, batch_words = take_request_batch(items, cursor, max_items_per_request, max_words_per_request)
        _account_count, estimated_seconds, selected_accounts, effective_capacity = choose_account_plan(
            batch_words,
            mode,
            scheduler_snapshot,
        )

        # Guardrail: when only one account is currently usable, keep request size
        # within per-account bounds to avoid long click/retry failure storms.
        if len(selected_accounts) == 1 and len(batch) > 1:
            max_single_account_words = MODE_MAX_WORDS_PER_ACCOUNT.get(mode, max_words_per_request)
            if batch_words > max_single_account_words:
                original_count = len(batch)
                while len(batch) > 1 and batch_words > max_single_account_words:
                    moved = batch.pop()
                    batch_words -= moved.word_count
                    cursor -= 1

                _account_count, estimated_seconds, selected_accounts, effective_capacity = choose_account_plan(
                    batch_words,
                    mode,
                    scheduler_snapshot,
                )
                print(
                    f"[2/4] request {request_count}: single-account guard trimmed "
                    f"{original_count}->{len(batch)} paragraphs ({batch_words} words)"
                )

        account_chunks = split_into_account_chunks(batch, selected_accounts)

        payload: Dict[str, str] = {"mode": mode}
        for account_key, account_items in account_chunks:
            payload[account_key] = build_batch_payload_text(account_items)

        print(
            f"[2/4] request {request_count}: {len(batch)} paragraphs, {batch_words} words, {len(account_chunks)} account(s)"
            f" | est={estimated_seconds:.1f}s | accounts={'/'.join(selected_accounts)} | capacity={effective_capacity:.0f}"
        )
        response = post_batch_request(api_url, payload, timeout_seconds)

        for account_key, account_items in account_chunks:
            request_label = f"request {request_count} {account_key}"
            try:
                paraphrased = extract_account_output(response, account_key, mode, f"request {request_count}")
                parts = parse_paraphrase_parts(paraphrased, len(account_items))
            except Exception as error:  # pylint: disable=broad-except
                print(f"[{request_label}] warning: initial chunk failed ({error}); retrying with recovery")
                parts = []

            if len(parts) != len(account_items):
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
                    f"Response count mismatch in {request_label}: "
                    f"expected {len(account_items)}, got {len(parts)} after recovery"
                )

            for item, new_text in zip(account_items, parts):
                paragraph = paragraph_by_index.get(item.paragraph_index)
                if paragraph is None:
                    continue
                single_paragraph_text = re.sub(r"\s*\n+\s*", " ", str(new_text)).strip()
                set_paragraph_text(paragraph, single_paragraph_text)
                updated_count += 1

    return Step2Stats(
        paraphrased_paragraphs=updated_count,
        request_count=request_count,
        total_words=total_words,
        mode=mode,
    )


def extract_reference_entries(paragraphs: Sequence[Paragraph], reference_start_index: int) -> List[str]:
    if reference_start_index < 0:
        return []

    section_texts = [p.text for p in paragraphs if p.index >= reference_start_index and p.text]
    if not section_texts:
        return []

    combined = "\n".join(section_texts)
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
        for paragraph in paragraphs:
            if paragraph.index < reference_start_index:
                continue
            if count_reference_like_lines(paragraph.text) > 0:
                entries.append(paragraph.text)

    # Deduplicate while preserving order.
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
    cleaned = re.sub(r"^\s*(?:\[\d{1,3}\]|\d{1,3}[.)\]]|[-•])\s*", "", reference).strip(" .;:")
    if not cleaned:
        return f"(Source {index + 1})"

    year_match = YEAR_RE.search(cleaned)
    year = year_match.group(0) if year_match else None

    prefix = cleaned[: year_match.start()] if year_match else cleaned
    # Keep only the leading author/source segment; drop title/details fragments.
    for marker in ("“", "\"", "’", "'"):
        marker_index = prefix.find(marker)
        if marker_index > 0:
            prefix = prefix[:marker_index]
            break
    if "." in prefix:
        prefix = prefix.split(".", 1)[0]
    prefix = normalize_space(prefix).strip(" ,.;:-()[]")

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
        if lower.startswith("in conclusion") or lower.startswith("to conclude") or lower.startswith("overall,") or lower.startswith("to sum up"):
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


def run_step_3_add_references(paragraphs: Sequence[Paragraph]) -> Step3Stats:
    reference_start_index, detection_mode = detect_reference_section(paragraphs)
    if reference_start_index == -1:
        raise RuntimeError(
            "Could not detect a reference section. No in-text references were added. "
            "Please verify the document has a references/reference list section."
        )
    conclusion_heading_index, conclusion_end_index = find_conclusion_range(paragraphs, reference_start_index)

    references = extract_reference_entries(paragraphs, reference_start_index)
    if not references:
        raise RuntimeError(
            "Reference section detected but no valid references could be parsed. "
            "No in-text references were added."
        )

    citations = [build_citation_from_reference(ref, idx) for idx, ref in enumerate(references)]
    if not citations:
        raise RuntimeError("Unable to build citation labels from detected references.")

    candidates: List[Paragraph] = []
    first_non_empty_index = next((p.index for p in paragraphs if p.text), -1)
    for p in paragraphs:
        if not p.text:
            continue
        if p.index >= reference_start_index:
            continue
        if (
            conclusion_heading_index != -1
            and p.index > conclusion_heading_index
            and (conclusion_end_index == -1 or p.index < conclusion_end_index)
        ):
            continue
        if p.index == first_non_empty_index:
            continue
        if is_heading_or_subtitle(p):
            continue
        if looks_like_toc_line(p.text):
            continue
        if p.word_count < 11:
            continue
        if p.text.endswith(":"):
            continue
        candidates.append(p)

    if not candidates:
        raise RuntimeError("No eligible body paragraphs found for inserting references.")

    target_count = min(len(candidates), max(1, len(citations)))
    inserted = 0
    citation_cursor = 0

    for paragraph in candidates:
        if inserted >= target_count:
            break
        citation = citations[citation_cursor % len(citations)]
        updated_text, changed = inject_citation_into_paragraph(paragraph.text, citation)
        if not changed:
            continue
        set_paragraph_text(paragraph, updated_text)
        inserted += 1
        citation_cursor += 1

    if inserted == 0:
        raise RuntimeError(
            "Reference section was detected, but no citations could be inserted into body paragraphs."
        )

    return Step3Stats(
        detection_mode=detection_mode,
        reference_start_index=reference_start_index,
        reference_count=len(references),
        inserted_citations=inserted,
    )


def write_output_docx(
    input_path: Path,
    output_path: Path,
    document_root: ET.Element,
    namespace_declarations: Optional[Dict[str, str]] = None,
) -> None:
    with zipfile.ZipFile(input_path, "r") as source_zip, zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as out_zip:
        for info in source_zip.infolist():
            if info.filename == "word/document.xml":
                xml_text = ET.tostring(document_root, encoding="utf-8", xml_declaration=True).decode("utf-8")
                xml_text = inject_missing_namespace_declarations(xml_text, namespace_declarations or {})
                xml_bytes = xml_text.encode("utf-8")
                out_zip.writestr(info, xml_bytes)
            elif info.filename in DOCX_METADATA_XML_PATHS:
                original_bytes = source_zip.read(info.filename)
                scrubbed_bytes = scrub_docx_metadata_xml(info.filename, original_bytes)
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
            "Too many positional arguments. Use: `npm run doc`, `npm run doc standard`, or pass one input path."
        )

    return mode, input_path


def resolve_input_path(provided_input: Optional[Path]) -> Path:
    if provided_input is not None:
        return provided_input

    desktop = Path.home() / "Desktop"
    if not desktop.exists() or not desktop.is_dir():
        raise RuntimeError(f"Desktop folder not found: {desktop}")

    desktop_docx = [
        path
        for path in desktop.iterdir()
        if path.is_file()
        and path.suffix.lower() == ".docx"
        and not path.name.startswith("~$")
    ]

    existing_names = {path.name.lower() for path in desktop_docx}
    candidates: List[Path] = []
    for path in desktop_docx:
        lower_name = path.name.lower()
        if lower_name.startswith("pr "):
            original_name = path.name[3:]
            if original_name.lower() in existing_names:
                continue
        candidates.append(path)

    candidates = sorted(candidates)
    if len(candidates) == 1:
        print(f"Auto-selected Desktop DOCX: {candidates[0]}")
        return candidates[0]
    if len(candidates) == 0:
        raise RuntimeError(
            f"No .docx file found on Desktop ({desktop}). Put one file there or pass the path explicitly."
        )

    preview = ", ".join(path.name for path in candidates[:5])
    suffix = "..." if len(candidates) > 5 else ""
    raise RuntimeError(
        f"Found {len(candidates)} .docx files on Desktop ({preview}{suffix}). "
        "Keep only one file there or pass the path explicitly."
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "DOCX one-shot pipeline: clean -> paraphrase -> add references -> output `pr <name>.docx`."
        )
    )
    parser.add_argument(
        "positional",
        nargs="*",
        help="Optional mode/input. Examples: `standard`, `/path/file.docx`, `standard /path/file.docx`",
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
        help="Output DOCX path (default: `pr <input-name>.docx`).",
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
        default=80,
        help="Max paragraphs sent per paraphrase API request.",
    )
    parser.add_argument(
        "--max-words-per-request",
        type=int,
        default=2200,
        help="Approximate max total words sent per paraphrase API request.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Run detection/eligibility checks without calling paraphrase API or writing output.",
    )
    return parser.parse_args()


def main() -> int:
    random.seed(42)
    args = parse_args()

    try:
        mode, provided_input = parse_mode_and_input(args.positional, args.mode)
        input_path = resolve_input_path(provided_input)
    except RuntimeError as error:
        print(str(error), file=sys.stderr)
        return 1

    if not input_path.exists():
        print(f"Input file does not exist: {input_path}", file=sys.stderr)
        return 1
    if input_path.suffix.lower() != ".docx":
        print(f"Input must be a .docx file: {input_path}", file=sys.stderr)
        return 1

    output_path = args.output if args.output else input_path.with_name(f"pr {input_path.name}")
    namespace_declarations: Dict[str, str] = {}

    try:
        with zipfile.ZipFile(input_path, "r") as input_zip:
            if "word/document.xml" not in input_zip.namelist():
                print("Invalid DOCX: missing word/document.xml", file=sys.stderr)
                return 1
            document_xml_bytes = input_zip.read("word/document.xml")
            namespace_declarations = extract_namespace_declarations(
                document_xml_bytes.decode("utf-8", errors="replace")
            )
            root = ET.fromstring(document_xml_bytes)
    except Exception as error:  # pylint: disable=broad-except
        print(f"Failed to read DOCX: {error}", file=sys.stderr)
        return 1

    paragraphs = collect_paragraphs(root)
    if not paragraphs:
        print("No paragraphs found in the document. Aborting.", file=sys.stderr)
        return 1

    # Initial reference detection for section-aware clean/paraphrase.
    reference_start_index, _ = detect_reference_section(paragraphs)

    print("[1/4] Clean (weird numbers + in-text citations, keep headings/subtitles) ...", flush=True)
    try:
        step1 = run_step_1_clean(paragraphs, reference_start_index)
        print(
            "[1/4] OK"
            f" | updated_paragraphs={step1.cleaned_paragraphs}"
            f", removed_citations={step1.removed_citations}"
            f", removed_weird_numbers={step1.removed_weird_numbers}"
        )
    except Exception as error:  # pylint: disable=broad-except
        print(f"[1/4] FAILED: {error}", file=sys.stderr)
        return 1

    print(f"[2/4] Paraphrase ({'SIMPLE+SHORT' if mode == 'dual' else 'STANDARD'}) ...", flush=True)
    try:
        if args.dry_run:
            eligible_count = sum(
                1
                for p in paragraphs
                if p.text
                and (reference_start_index == -1 or p.index < reference_start_index)
                and not is_heading_or_subtitle(p)
                and not looks_like_toc_line(p.text)
                and p.word_count >= 15
            )
            if eligible_count == 0:
                raise RuntimeError("No eligible body paragraphs found for paraphrasing.")
            print(f"[2/4] DRY-RUN OK | eligible_paragraphs={eligible_count}, mode={mode}")
        else:
            step2 = run_step_2_paraphrase(
                paragraphs=paragraphs,
                mode=mode,
                api_url=args.api_url,
                timeout_seconds=args.timeout_seconds,
                max_items_per_request=args.max_items_per_request,
                max_words_per_request=args.max_words_per_request,
                reference_start_index=reference_start_index,
            )
            print(
                "[2/4] OK"
                f" | paraphrased_paragraphs={step2.paraphrased_paragraphs}"
                f", requests={step2.request_count}"
                f", words={step2.total_words}"
                f", mode={step2.mode}"
            )
    except Exception as error:  # pylint: disable=broad-except
        print(f"[2/4] FAILED: {error}", file=sys.stderr)
        return 1

    print("[3/4] Add new in-text references ...", flush=True)
    try:
        if args.dry_run:
            ref_start, ref_mode = detect_reference_section(paragraphs)
            if ref_start == -1:
                raise RuntimeError("Could not detect a reference section for citation insertion.")
            refs = extract_reference_entries(paragraphs, ref_start)
            if not refs:
                raise RuntimeError("Reference section found, but no usable references were parsed.")
            print(
                "[3/4] DRY-RUN OK"
                f" | detection={ref_mode}"
                f", reference_start={ref_start}"
                f", references={len(refs)}"
            )
        else:
            step3 = run_step_3_add_references(paragraphs)
            print(
                "[3/4] OK"
                f" | detection={step3.detection_mode}"
                f", reference_start={step3.reference_start_index}"
                f", parsed_references={step3.reference_count}"
                f", inserted_citations={step3.inserted_citations}"
            )
    except Exception as error:  # pylint: disable=broad-except
        print(
            "[3/4] FAILED: "
            f"{error}\n"
            "No output file was written. Review references formatting before sending this document.",
            file=sys.stderr,
        )
        return 1

    print("[4/4] Write output DOCX ...", flush=True)
    try:
        if args.dry_run:
            print("[4/4] DRY-RUN OK | no file written")
        else:
            write_output_docx(input_path, output_path, root, namespace_declarations)
            print(f"[4/4] OK | output={output_path}")
    except Exception as error:  # pylint: disable=broad-except
        print(f"[4/4] FAILED: {error}", file=sys.stderr)
        return 1

    if args.dry_run:
        print("Pipeline dry-run completed successfully.")
    else:
        print("Pipeline completed successfully.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
