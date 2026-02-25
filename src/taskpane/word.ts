/* global Word console, fetch, AbortController, setTimeout, clearTimeout */
import { calculateTextChangeMetrics } from "./changeMetrics";
import { getFormattedReferences } from "./gemini";
import { chooseAdaptiveAccountCount, type HealthSnapshot } from "./accountSelection";

// ==================== Batch API Configuration ====================
const BATCH_API_URL = "https://analizeai.com/paraphrase-batch";
const HEALTH_CHECK_URL = "https://analizeai.com/health";
const HEALTH_CHECK_TIMEOUT = 2000; // 2 seconds
const ACCOUNT_KEYS = ["acc1", "acc2", "acc3"] as const;

type AccountKey = (typeof ACCOUNT_KEYS)[number];

// Metrics for tracking text changes during paraphrasing
export interface ChangeMetrics {
  originalWordCount: number;
  newWordCount: number;
  wordsChanged: number;
  wordChangePercent: number;
  reusedWords: number;
  reusePercent: number;
  addedWords: number;
  originalPreview: string;
  newPreview: string;
}

// Result type for paraphrase functions with warnings support
export interface ParaphraseResult {
  message: string;
  warnings: string[];
  metrics?: ChangeMetrics;
}

// Health check response type
interface HealthCheckResponse {
  status?: string;
  ready?: boolean;
  scheduler?: HealthSnapshot["scheduler"];
  acc1?: HealthSnapshot["acc1"];
  acc2?: HealthSnapshot["acc2"];
  acc3?: HealthSnapshot["acc3"];
}

// Batch API response types
interface BatchAccountResult {
  firstMode?: string;
  secondMode?: string;
  result?: string;
  durationMs?: number;
  error?: string;
  fallbackUsed?: string;
}

interface BatchApiResponse {
  acc1?: BatchAccountResult;
  acc2?: BatchAccountResult;
  acc3?: BatchAccountResult;
}

/**
 * Check service health with a 2-second timeout.
 * Returns warnings for accounts that are not ready, but never throws.
 */
async function checkServiceHealth(): Promise<{
  ready: boolean;
  warnings: string[];
  statusData?: HealthCheckResponse;
}> {
  const warnings: string[] = [];

  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), HEALTH_CHECK_TIMEOUT);

    const response = await fetch(HEALTH_CHECK_URL, {
      method: "GET",
      signal: controller.signal,
    });

    clearTimeout(timeoutId);

    if (!response.ok) {
      warnings.push(`Health check returned status ${response.status}`);
      return { ready: false, warnings, statusData: undefined };
    }

    const data: HealthCheckResponse = await response.json();

    // Check each account status
    for (const key of ACCOUNT_KEYS) {
      const accountStatus = data[key]?.status;
      if (accountStatus && accountStatus !== "ready") {
        warnings.push(`Account ${key} is ${accountStatus}`);
      }
    }

    return { ready: data.ready ?? true, warnings, statusData: data };
  } catch (error) {
    if (error.name === "AbortError") {
      warnings.push("Health check timed out (service may be slow)");
    } else {
      warnings.push(`Health check failed: ${error.message}`);
    }
    return { ready: false, warnings, statusData: undefined };
  }
}

/**
 * Collect warnings from batch API response for accounts that had errors or used fallback
 */
function collectBatchWarnings(data: BatchApiResponse, usedAccounts: AccountKey[]): string[] {
  const warnings: string[] = [];

  for (const key of usedAccounts) {
    const result = data[key];
    if (!result) continue;

    if (result.error) {
      if (result.fallbackUsed) {
        warnings.push(`${key} failed (${result.error}), used ${result.fallbackUsed} as fallback`);
      } else {
        warnings.push(`${key} error: ${result.error}`);
      }
    } else if (result.fallbackUsed) {
      warnings.push(`${key} used ${result.fallbackUsed} as fallback`);
    }
  }

  return warnings;
}

// ==================== End Batch API Configuration ====================

const TRAILING_PUNCTUATION = "[-:;,.!?–—]*";
// Allow for "1.", "1.1", "IV.", "2 " prefixes. Changed \s+ to \s* to allow "1.Conclusion"
const NUMBERING_PREFIX = "(?:(?:\\d+(?:\\.\\d+)*|[IVX]+)\\.?\\s*)?";

const REFERENCE_HEADERS = [
  new RegExp(`^\\s*${NUMBERING_PREFIX}references?(?:\\s+list)?(?:\\s+section)?\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}reference\\s+list\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}references\\s+list\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}bibliograph(?:y|ies)\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}list\\s+of\\s+references\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
];

const CONCLUSION_HEADERS = [
  new RegExp(`^\\s*${NUMBERING_PREFIX}conclusions?(?:\\s+section)?\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}concluding\\s+remarks\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}final\\s+thoughts\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}summary(?:\\s+and\\s+future\\s+work)?\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}closing\\s+remarks\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}conclusions?\\s+and\\s+recommendations\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
];

const AIM_OBJECTIVE_HEADERS = [
  new RegExp(`^\\s*${NUMBERING_PREFIX}aims?\\s*(?:and|&)\\s*objectives?\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}aim\\s*&\\s*objectives\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}goal(?:s)?\\s*(?:and|&)\\s*objectives?\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
];

const RESEARCH_QUESTION_HEADERS = [
  new RegExp(`^\\s*${NUMBERING_PREFIX}research\\s+questions?(?:\\s+section)?\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}research\\s+question\\s+\\d+\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
];

const TOC_HEADERS = [
  new RegExp(`^\\s*${NUMBERING_PREFIX}table\\s+of\\s+contents?\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}contents?\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}toc\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
];

const YEAR_RE = /\b(?:19|20)\d{2}[a-z]?\b/i;
const URL_OR_DOI_RE = /\b(?:https?:\/\/|www\.|doi:\s*|10\.\d{4,9}\/)\S+/i;
const REFERENCE_CUE_RE = /\b(?:available at|retrieved from|accessed|doi|journal|vol\.?|no\.?|pp\.?|edition|ed\.)\b/i;
const AUTHOR_RE = /(?:^|[\s;])(?:[A-Z][a-z]+,\s*(?:[A-Z]\.|[A-Z][a-z]+))/;
const ORG_AUTHOR_RE = /^\s*[A-Z][A-Za-z0-9&'’.\- ]{2,80}\s*\((?:19|20)\d{2}[a-z]?\)/;
const LIST_PREFIX_RE = /^\s*(?:\[\d{1,3}\]|\d{1,3}[.)\]]|[-•])\s+/;

type ParagraphMeta = {
  index: number;
  text: string;
  wordCount: number;
  style: string;
  styleBuiltIn: string;
  alignment: string;
};

type SectionRange = {
  name: string;
  headingIndex: number;
  endIndex: number;
};

function buildParagraphMeta(paragraphs: Word.ParagraphCollection): ParagraphMeta[] {
  return paragraphs.items.map((p, index) => {
    const rawText = p.text || "";
    // Remove zero-width spaces (\u200B), zero-width non-joiner (\u200C), zero-width joiner (\u200D), and BOM (\uFEFF)
    const text = rawText.replace(/[\u200B-\u200D\uFEFF]/g, "").trim();
    const words = text.split(/\s+/).filter(Boolean);
    const style = (p.style || "").toString().toLowerCase();
    const styleBuiltIn = (p.styleBuiltIn || "").toString().toLowerCase();
    const alignment = (typeof p.alignment === "string" ? p.alignment : String(p.alignment || "")).toLowerCase();

    return {
      index,
      text,
      wordCount: words.length,
      style,
      styleBuiltIn,
      alignment,
    };
  });
}

function isHeadingOrTitle(meta: ParagraphMeta, debug: boolean = false): boolean {
  const s = meta.style;
  const sb = meta.styleBuiltIn;

  // Only treat Heading 1, 2, 3 as actual headings (titles)
  // Heading 4, 5, 6, etc. may contain body content in some documents
  const isMainHeading =
    /heading\s*[123]$/i.test(s) ||
    /heading[123]$/i.test(sb) ||
    s.includes("title") ||
    s.includes("subtitle") ||
    sb.includes("title") ||
    sb.includes("subtitle");

  if (isMainHeading) {
    if (debug) {
      console.log(
        `isHeadingOrTitle => main heading style: style="${s}", styleBuiltIn="${sb}", text="${meta.text.substring(0, 50)}..."`
      );
    }
    return true;
  }

  const text = meta.text;
  if (!text) return false;

  // Detect numbered subtitles like "2.1. AI in Road Transport" or "1. Introduction"
  // These start with a numbering prefix (digits/roman numerals + dots) followed by text.
  const numberedSubtitleMatch = text.match(/^(\d+(?:\.\d+)*\.?\s+|[IVXivx]+\.?\s+)/);
  if (numberedSubtitleMatch && meta.wordCount <= 15) {
    const afterNumber = text.slice(numberedSubtitleMatch[0].length).trim();
    // If the text after the number is short and doesn't end with sentence-final punctuation,
    // treat it as a heading/subtitle.
    const afterWords = afterNumber.split(/\s+/).filter(Boolean);
    const endsWithSentencePunct = /[.!?]$/.test(afterNumber) && afterWords.length > 3;
    if (!endsWithSentencePunct) {
      if (debug) {
        console.log(
          `isHeadingOrTitle => numbered subtitle: text="${meta.text.substring(0, 50)}..."`
        );
      }
      return true;
    }
  }

  const hasTerminalPunctuation = /[.!?]$/.test(text);
  const isShortish = meta.wordCount > 0 && meta.wordCount <= 15;
  const isCentered = meta.alignment === "centered";

  const words = text.split(/\s+/).filter(Boolean);
  const capitalised = words.filter((w) => /^[A-Z]/.test(w)).length;
  const isTitleCase = words.length > 1 && capitalised / words.length >= 0.6;

  if (!hasTerminalPunctuation && isShortish && (isCentered || isTitleCase)) {
    if (debug) {
      console.log(
        `isHeadingOrTitle => heuristic match: centered=${isCentered}, titleCase=${isTitleCase}, text="${meta.text.substring(0, 50)}..."`
      );
    }
    return true;
  }

  return false;
}

function looksLikeTOC(text: string): boolean {
  const trimmed = text.trim();
  if (!trimmed) return false;
  // Match dot leaders with page numbers (5+ dots followed by number)
  const dotLeaderWithPage = /\.{5,}.*\d+\s*$/.test(trimmed);
  // Match dot leaders with fewer dots (3+) followed by page number
  const shortDotLeader = /\.{3,}\s*\d+\s*$/.test(trimmed);
  // Match tabs (common in TOC entries)
  const hasTab = /\t/.test(trimmed);
  // Match patterns like "Chapter 1 ..... 15" or "Introduction ... 2"
  const spacedDotLeader = /\s+\.{2,}\s*\d+\s*$/.test(trimmed);
  return dotLeaderWithPage || shortDotLeader || hasTab || spacedDotLeader;
}

/**
 * Find the TOC section range in the document.
 * Returns the start and end paragraph indexes of the TOC section.
 */
function findTOCSectionRange(metas: ParagraphMeta[]): { startIndex: number; endIndex: number } {
  let startIndex = -1;

  // Find TOC header
  for (let i = 0; i < metas.length; i++) {
    if (TOC_HEADERS.some((regex) => regex.test(metas[i].text))) {
      startIndex = i;
      break;
    }
  }

  if (startIndex === -1) {
    return { startIndex: -1, endIndex: -1 };
  }

  // Find end of TOC section - look for the last consecutive TOC-like entry
  // or stop at the next major heading
  let endIndex = startIndex;

  for (let i = startIndex + 1; i < metas.length; i++) {
    const meta = metas[i];
    const text = meta.text;

    // If it looks like a TOC entry, extend the end index
    if (looksLikeTOC(text)) {
      endIndex = i;
      continue;
    }

    // Empty paragraphs within TOC are OK
    if (!text) {
      continue;
    }

    // If we hit a heading that's not a TOC header, we've found the end
    if (isHeadingOrTitle(meta) && !TOC_HEADERS.some((regex) => regex.test(text))) {
      break;
    }

    // If we've gone more than 3 paragraphs without a TOC-like entry, stop
    if (i - endIndex > 3) {
      break;
    }
  }

  console.log(`TOC Section detected: start=${startIndex}, end=${endIndex}`);
  return { startIndex, endIndex };
}

export function findReferenceStartIndexFromTexts(paragraphTexts: string[]): number {
  for (let i = paragraphTexts.length - 1; i >= 0; i--) {
    const text = (paragraphTexts[i] || "").replace(/[\u200B-\u200D\uFEFF]/g, "").trim();
    if (REFERENCE_HEADERS.some((regex) => regex.test(text))) {
      return i;
    }
  }

  return inferReferenceStartIndexFromTexts(paragraphTexts);
}

function findReferenceStartIndex(metas: ParagraphMeta[]): number {
  return findReferenceStartIndexFromTexts(metas.map((meta) => meta.text));
}

function normalizeSpace(text: string): string {
  return (text || "").replace(/\s{2,}/g, " ").trim();
}

function splitReferenceCandidateLines(text: string): string[] {
  const prepared = (text || "")
    .replace(/\r/g, "\n")
    .replace(/(?<!\d)(\d{1,3}[.)]\s*)/g, "\n$1");
  const lines = prepared
    .split(/\n+/)
    .map((line) => normalizeSpace(line))
    .filter(Boolean);
  return lines.length > 0 ? lines : normalizeSpace(text) ? [normalizeSpace(text)] : [];
}

function isReferenceLikeLine(rawLine: string): boolean {
  const line = normalizeSpace(rawLine);
  if (!line) return false;
  if (line.split(/\s+/).filter(Boolean).length < 4) return false;

  const hasYear = YEAR_RE.test(line);
  const hasUrlOrDoi = URL_OR_DOI_RE.test(line);
  const hasCue = REFERENCE_CUE_RE.test(line);
  const hasAuthor = AUTHOR_RE.test(line);
  const hasOrgAuthor = ORG_AUTHOR_RE.test(line);
  const hasListPrefix = LIST_PREFIX_RE.test(line);

  if (hasUrlOrDoi && (hasYear || hasAuthor || hasListPrefix)) return true;
  if (hasYear && (hasAuthor || hasOrgAuthor || hasCue || hasListPrefix)) return true;
  return false;
}

function countReferenceLikeLines(text: string): number {
  return splitReferenceCandidateLines(text).filter((line) => isReferenceLikeLine(line)).length;
}

function inferReferenceStartIndexFromTexts(paragraphTexts: string[]): number {
  if (!paragraphTexts || paragraphTexts.length === 0) return -1;

  const total = paragraphTexts.length;
  const tailStart = Math.max(0, Math.floor(total * 0.45));
  const scores = paragraphTexts.map((text) => countReferenceLikeLines(text || ""));

  const denseCandidate = scores.findIndex((score, idx) => idx >= tailStart && score >= 3);
  if (denseCandidate !== -1) return denseCandidate;

  const scoredIndices: number[] = [];
  for (let i = tailStart; i < total; i++) {
    if (scores[i] >= 1) scoredIndices.push(i);
  }
  if (scoredIndices.length < 2) return -1;

  const clusters: number[][] = [];
  let cluster: number[] = [scoredIndices[0]];
  for (let i = 1; i < scoredIndices.length; i++) {
    const idx = scoredIndices[i];
    if (idx - cluster[cluster.length - 1] <= 2) {
      cluster.push(idx);
    } else {
      clusters.push(cluster);
      cluster = [idx];
    }
  }
  clusters.push(cluster);

  let bestCluster: number[] | null = null;
  let bestScore = -1;
  for (const c of clusters) {
    const score = c.reduce((sum, idx) => sum + scores[idx], 0);
    if (score > bestScore) {
      bestScore = score;
      bestCluster = c;
    } else if (score === bestScore && bestCluster && c[c.length - 1] > bestCluster[bestCluster.length - 1]) {
      bestCluster = c;
    }
  }

  if (!bestCluster) return -1;
  if (bestScore >= 3 || bestCluster.length >= 2) return bestCluster[0];
  return -1;
}

function isInReferenceSection(paragraphIndex: number, referenceStartIndex: number): boolean {
  return referenceStartIndex !== -1 && paragraphIndex >= referenceStartIndex;
}

function findSectionRange(
  metas: ParagraphMeta[],
  referenceStartIndex: number,
  headerRegexes: RegExp[],
  sectionName: string
): SectionRange {
  let headingIndex = -1;

  for (let i = referenceStartIndex === -1 ? metas.length - 1 : referenceStartIndex - 1; i >= 0; i--) {
    const txt = metas[i].text;
    if (headerRegexes.some((regex) => regex.test(txt))) {
      headingIndex = i;
      break;
    }
  }

  let endIndex = -1;
  let reason = "End of document (no next section found)";

  if (headingIndex !== -1) {
    for (let i = headingIndex + 1; i < metas.length; i++) {
      if (REFERENCE_HEADERS.some((regex) => regex.test(metas[i].text))) {
        endIndex = i;
        reason = `Found Reference Header at index ${i}: "${metas[i].text}"`;
        break;
      }

      if (isHeadingOrTitle(metas[i]) && !headerRegexes.some((regex) => regex.test(metas[i].text))) {
        endIndex = i;
        reason = `Found next Heading at index ${i}: "${metas[i].text}"`;

        const meta = metas[i];
        const s = meta.style;
        const sb = meta.styleBuiltIn;
        const isStyleMatch =
          s.includes("heading") ||
          s.includes("title") ||
          s.includes("subtitle") ||
          sb.includes("heading") ||
          sb.includes("title") ||
          sb.includes("subtitle");

        const text = meta.text;
        const hasTerminalPunctuation = /[.!?]$/.test(text);
        const isShortish = meta.wordCount > 0 && meta.wordCount <= 15;
        const isCentered = meta.alignment === "centered";
        const words = text.split(/\s+/).filter(Boolean);
        const capitalised = words.filter((w) => /^[A-Z]/.test(w)).length;
        const isTitleCase = words.length > 1 && capitalised / words.length > 0.6;

        console.log(`${sectionName} Heading Detection Details for index ${i}:`, {
          text: meta.text,
          style: meta.style,
          styleBuiltIn: meta.styleBuiltIn,
          alignment: meta.alignment,
          wordCount: meta.wordCount,
          heuristics: {
            isStyleMatch,
            hasTerminalPunctuation,
            isShortish,
            isCentered,
            isTitleCase,
            capitalisedRatio: words.length > 0 ? capitalised / words.length : 0,
          },
        });
        break;
      }
    }

    if (endIndex === -1 && referenceStartIndex !== -1) {
      endIndex = referenceStartIndex;
      reason = `Reached Reference Start Index at ${referenceStartIndex}`;
    }
  }

  console.log(`${sectionName} Detection: Start Index: ${headingIndex}`);
  if (headingIndex !== -1) {
    console.log(`${sectionName} Detection: End Index: ${endIndex}`);
    console.log(`${sectionName} Detection: End Reason: ${reason}`);
  }

  return { name: sectionName, headingIndex, endIndex };
}

interface SentenceSlot {
  paragraphIndex: number;
  sentenceIndex: number;
}

function shuffleInPlace<T>(array: T[]): T[] {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

function getSentenceSlotsForParagraph(paragraphIndex: number, sentenceRanges: Word.RangeCollection): SentenceSlot[] {
  const slots: SentenceSlot[] = [];
  const items = sentenceRanges.items;

  for (let i = 0; i < items.length; i++) {
    const raw = (items[i].text || "").trim();
    if (!raw) continue;

    // Skip first sentence in multi-sentence paragraphs.
    // Allow the last sentence to be a candidate (end of paragraph).
    if (i === 0 && items.length > 1) continue;

    const wordCount = raw.split(/\s+/).filter(Boolean).length;
    if (wordCount < 8) continue; // sentence too short

    // Skip sentences that already have a citation‑ish thing
    if (/\(\s*[^)]*?\d{4}[a-z]?\s*\)/.test(raw)) continue;

    // Optional: skip sentences starting with “In conclusion”, “Overall,” etc.
    const lower = raw.toLowerCase();
    if (
      lower.startsWith("in conclusion") ||
      lower.startsWith("to conclude") ||
      lower.startsWith("overall,") ||
      lower.startsWith("to sum up")
    ) {
      continue;
    }

    slots.push({ paragraphIndex, sentenceIndex: i });
  }

  // Safety rule: at most one sentence per paragraph.
  // This avoids multi-range mutations inside the same paragraph, which can destabilize
  // Word range anchors and occasionally collapse visual paragraph boundaries.
  if (slots.length <= 1) return slots;

  shuffleInPlace(slots);
  return [slots[0]];
}

function appendCitationAtSentenceEnd(sentence: string, citation: string): string {
  const trimmed = sentence.trimEnd();

  const match = trimmed.match(/([.?!]["')\]]*)$/);
  if (!match) {
    // No obvious sentence‑ending punctuation; just append.
    const sep = /\s$/.test(trimmed) ? "" : " ";
    return trimmed + sep + citation;
  }

  const punctuation = match[1];
  const core = trimmed.slice(0, trimmed.length - punctuation.length);
  const sep = /\s$/.test(core) ? "" : " ";

  return `${core}${sep}${citation}${punctuation}`;
}

/**
 * Insert a simple text
 */
export async function insertText(text: string) {
  // Write text to the document.
  try {
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function analyzeDocument(insertEveryOther: boolean = false): Promise<string> {
  try {
    return await Word.run(async (context) => {
      console.log("analyzeDocument => start", { insertEveryOther });
      // Get all paragraphs from the document
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("text,style,styleBuiltIn,alignment");
      await context.sync();

      const metas = buildParagraphMeta(paragraphs);
      console.log("analyzeDocument => paragraph count", metas.length);
      if (metas.length === 0) {
        console.warn("analyzeDocument => abort: empty document");
        return "No content found in the document";
      }

      const referenceStartIndex = findReferenceStartIndex(metas);
      console.log({ referenceStartIndex });
      if (referenceStartIndex === -1) {
        console.warn("analyzeDocument => abort: missing reference header");
        return "No Reference List section found";
      }

      const referenceSection = paragraphs.items
        .slice(referenceStartIndex)
        .map((p) => p.text)
        .join("\n");
      console.log("analyzeDocument => referenceSection length", referenceSection.length);

      let references: string[] = [];
      try {
        const formattedRefs = await getFormattedReferences(referenceSection);
        references = formattedRefs
          .split(/\n\s*\n/)
          .map((ref) => ref.trim())
          .filter(Boolean);
        console.log("analyzeDocument => formatted reference count", references.length);
      } catch (error) {
        console.error("Error in getFormattedReferences:", error);
        throw error;
      }

      if (references.length === 0) {
        console.warn("analyzeDocument => abort: gemini returned no references");
        return "No valid references found in the Reference List section";
      }

      const excludedSections: SectionRange[] = [
        findSectionRange(metas, referenceStartIndex, CONCLUSION_HEADERS, "Conclusion"),
        findSectionRange(metas, referenceStartIndex, AIM_OBJECTIVE_HEADERS, "Aim and Objectives"),
        findSectionRange(metas, referenceStartIndex, RESEARCH_QUESTION_HEADERS, "Research Questions"),
      ];
      console.log(
        "analyzeDocument => section ranges",
        excludedSections.map(({ name, headingIndex, endIndex }) => ({ name, headingIndex, endIndex }))
      );

      // Filter out short paragraphs, TOC lines, and those ending with ":"
      let debugSkipCounts = { reference: 0, excludedSection: 0, heading: 0, short: 0, toc: 0, endsColon: 0 };
      const eligibleIndexes = metas
        .filter((meta) => {
          if (referenceStartIndex !== -1 && meta.index >= referenceStartIndex) {
            debugSkipCounts.reference++;
            return false;
          }

          const insideExcludedSection = excludedSections.some((section) => {
            if (section.headingIndex === -1 || section.endIndex === -1) {
              return false;
            }
            return meta.index > section.headingIndex && meta.index < section.endIndex;
          });

          if (insideExcludedSection) {
            debugSkipCounts.excludedSection++;
            return false;
          }

          if (isHeadingOrTitle(meta, debugSkipCounts.heading < 5)) {
            debugSkipCounts.heading++;
            return false;
          }

          if (meta.wordCount < 11) {
            debugSkipCounts.short++;
            return false;
          }

          if (looksLikeTOC(meta.text)) {
            debugSkipCounts.toc++;
            return false;
          }

          if (meta.text.trim().endsWith(":")) {
            debugSkipCounts.endsColon++;
            return false;
          }

          return true;
        })
        .map((meta) => meta.index);

      console.log("analyzeDocument => skip reasons", debugSkipCounts);

      console.log("analyzeDocument => eligibleIndexes count (before first page filter)", eligibleIndexes.length);

      // Safety: remove first non-empty paragraph even if it passed filters
      const firstNonEmpty = metas.find((meta) => meta.text.length > 0);
      if (firstNonEmpty) {
        const idx = eligibleIndexes.indexOf(firstNonEmpty.index);
        if (idx !== -1) {
          eligibleIndexes.splice(idx, 1);
        }
      }

      // NEW: Exclude all paragraphs on the first page
      // Calculate approximate first page end by cumulative character count
      // Typically a page is ~500-700 words or ~3000-4000 characters
      let firstPageEndIndex = -1;
      let cumulativeChars = 0;
      const FIRST_PAGE_CHAR_THRESHOLD = 1500; // Approximate characters per page

      for (let i = 0; i < metas.length; i++) {
        cumulativeChars += metas[i].text.length;
        if (cumulativeChars > FIRST_PAGE_CHAR_THRESHOLD) {
          firstPageEndIndex = i;
          break;
        }
      }

      // Filter out paragraphs that are likely on the first page
      const firstPageParagraphIndexes = new Set<number>();
      for (const idx of eligibleIndexes) {
        if (firstPageEndIndex !== -1 && idx <= firstPageEndIndex) {
          firstPageParagraphIndexes.add(idx);
        }
      }

      // Remove first page paragraphs from eligible indexes
      const filteredIndexes = eligibleIndexes.filter((idx) => !firstPageParagraphIndexes.has(idx));
      console.log(
        `analyzeDocument => excluded ${firstPageParagraphIndexes.size} paragraphs from first page (up to index ${firstPageEndIndex})`
      );

      if (filteredIndexes.length === 0) {
        console.warn("analyzeDocument => abort: no eligible paragraphs after excluding first page");
        return "No eligible paragraphs found for inserting references (all content on first page)";
      }
      console.log("analyzeDocument => eligible count after first page filter", filteredIndexes.length);

      // If insertEveryOther is true, only use every other paragraph
      let targetIndexes = [...filteredIndexes];
      if (insertEveryOther && targetIndexes.length > 0) {
        // Sort available indexes to ensure we're working with ordered paragraphs
        targetIndexes.sort((a, b) => a - b);
        // Filter to only include every other paragraph
        targetIndexes = targetIndexes.filter((_, i) => i % 2 === 0);
        console.log("analyzeDocument => insertEveryOther applied", targetIndexes.length);
      }

      // NEW: Sentence-level injection
      const allSentenceSlots: SentenceSlot[] = [];
      const sentenceRangeByParagraph: { [index: number]: Word.RangeCollection } = {};

      for (const pIndex of targetIndexes) {
        const paragraph = paragraphs.items[pIndex];
        // Build sentence ranges from paragraph content only (exclude paragraph mark).
        const ranges = paragraph.getRange(Word.RangeLocation.content).getTextRanges([".", "?", "!"], true);
        ranges.load("text");
        sentenceRangeByParagraph[pIndex] = ranges;
      }

      await context.sync();

      for (const pIndex of targetIndexes) {
        const ranges = sentenceRangeByParagraph[pIndex];
        if (!ranges) continue;

        const paragraphSlots = getSentenceSlotsForParagraph(pIndex, ranges);
        allSentenceSlots.push(...paragraphSlots);
      }

      shuffleInPlace(allSentenceSlots);
      console.log("analyzeDocument => sentence slots count", allSentenceSlots.length);

      const usedReferences = new Set<number>();

      for (const slot of allSentenceSlots) {
        const ranges = sentenceRangeByParagraph[slot.paragraphIndex];
        const sentenceRange = ranges.items[slot.sentenceIndex];

        let referenceIndex: number;
        if (usedReferences.size < references.length) {
          const unusedReferences = Array.from(Array(references.length).keys()).filter((i) => !usedReferences.has(i));
          referenceIndex = unusedReferences[Math.floor(Math.random() * unusedReferences.length)];
          usedReferences.add(referenceIndex);
        } else {
          referenceIndex = Math.floor(Math.random() * references.length);
        }

        // Guardrail: never edit ranges that include line/paragraph separators.
        const currentText = sentenceRange.text || "";
        if (/[\r\n\u000B\u000C]/.test(currentText)) {
          console.log("analyzeDocument => skipping unsafe boundary sentence range", {
            paragraphIndex: slot.paragraphIndex,
            sentenceIndex: slot.sentenceIndex,
          });
          continue;
        }
        const citation = references[referenceIndex];
        const newText = appendCitationAtSentenceEnd(currentText, citation);

        console.log("analyzeDocument => inserting reference", {
          paragraphIndex: slot.paragraphIndex,
          sentenceIndex: slot.sentenceIndex,
          referenceIndex,
        });

        // Replace content-only range to avoid touching paragraph boundaries.
        sentenceRange.getRange(Word.RangeLocation.content).insertText(newText, Word.InsertLocation.replace);
      }

      await context.sync();
      console.log("analyzeDocument => completed");
      return "References added successfully";
    });
  } catch (error) {
    console.error("Error in analyzeDocument:", error);
    throw new Error(`Error modifying document: ${error.message}`);
  }
}

/**
 * Removes in-text citations from the document and cleans up any resulting formatting issues.
 * Handles various citation formats including:
 * - (Author, YYYY)
 * - (Author and Author, YYYY)
 * - (Author et al., YYYY)
 * - (Author, Author, Author, and Author, YYYY)
 * - (Author, Author, and Author, YYYY)
 *
 * @returns {Promise<string>} A message indicating the number of citations removed
 */
export async function removeReferences(): Promise<string> {
  try {
    return await Word.run(async (context) => {
      console.log("Starting reference removal process...");

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("text,style,styleBuiltIn,alignment");
      await context.sync();

      console.log(`Found ${paragraphs.items.length} paragraphs to process`);

      // Build paragraph metadata and find TOC section to skip
      const metas = buildParagraphMeta(paragraphs);
      const tocRange = findTOCSectionRange(metas);
      const referenceStartIndex = findReferenceStartIndex(metas);
      console.log(`TOC section range: ${tocRange.startIndex} to ${tocRange.endIndex}`);
      console.log(`Reference section start index: ${referenceStartIndex}`);

      // Updated citation patterns to be more precise
      const citationPatterns = [
        // Square bracket citations: [Harvard DCE, 2025], [Author, 2024], etc.
        /\[(?:[^\]]+)[,\s]\s?\d{4}[a-z]?\]/g,

        // Multiple authors with 'and' and commas - non-greedy match
        /\((?:[^,()]+(,\s[^,()]+)*(?:,\sand\s[^,()]+)?)[,\s]\s?\d{4}[a-z]?\)/g, // Added [a-z]? to handle letter suffixes

        // Standard patterns - non-greedy match
        /\((?:[^,()]+)[,\s]\s?\d{4}[a-z]?\)/g, // Added [a-z]? to handle letter suffixes
        /\((?:[^()]+\sand\s[^,()]+)[,\s]\s?\d{4}[a-z]?\)/g, // Added [a-z]? to handle letter suffixes
        /\((?:[^()]+)\set\sal\.?[,\s]\s?\d{4}[a-z]?\)/g, // Added [a-z]? to handle letter suffixes

        // Additional patterns for edge cases - non-greedy match
        /\((?:[^,()]+(,\s[^,()]+)*)[,\s]\s?\d{4}[a-z]?\)/g, // Added [a-z]? to handle letter suffixes
      ];

      let totalRemoved = 0;

      // Process each paragraph
      for (let i = 0; i < paragraphs.items.length; i++) {
        const meta = metas[i];

        if (isInReferenceSection(i, referenceStartIndex)) {
          console.log(`Skipping reference paragraph ${i + 1}`);
          continue;
        }

        // Skip TOC section paragraphs
        if (tocRange.startIndex !== -1 && i >= tocRange.startIndex && i <= tocRange.endIndex) {
          console.log(`Skipping TOC paragraph ${i + 1}`);
          continue;
        }

        // Skip individual TOC-like entries (fallback for entries outside detected TOC section)
        if (looksLikeTOC(meta.text)) {
          console.log(`Skipping TOC-like entry at paragraph ${i + 1}`);
          continue;
        }

        const paragraph = paragraphs.items[i];
        let text = paragraph.text;
        let hadMatch = false;
        let originalText = text;
        console.log(`\nProcessing paragraph ${i + 1}:`, text);

        // Apply each pattern
        for (const pattern of citationPatterns) {
          const matches = text.match(pattern) || [];
          if (matches.length > 0) {
            console.log(`Found matches with pattern ${pattern}:`, matches);

            // Remove all matches and clean up extra spaces before periods.
            // Use space/tab-only cleanup so line/paragraph separators are never collapsed.
            text = text.replace(pattern, "").replace(/[ \t]+\./g, ".");
            hadMatch = true;
            totalRemoved += matches.length;
          }
        }

        // Only update paragraph if we found and removed citations
        if (hadMatch) {
          // Replace only paragraph content (not paragraph mark) to preserve boundaries.
          paragraph.getRange(Word.RangeLocation.content).insertText(text, Word.InsertLocation.replace);
          console.log("Original:", originalText);
          console.log("Updated:", text);
        }
      }

      await context.sync();

      console.log(`\nCompleted reference removal. Removed ${totalRemoved} citations.`);
      return `Removed ${totalRemoved} citations and cleaned up formatting`;
    });
  } catch (error) {
    console.error("Error in removeReferences:", error);
    const errorMessage = `Error removing references: ${error.message}`;
    console.error(errorMessage);
    return errorMessage;
  }
}

/**
 * Removes hyperlinks from the document, skipping the "References" section.
 * @returns {Promise<string>} A message indicating the number of hyperlinks removed.
 */
export async function removeLinks(deleteAll: boolean = false): Promise<string> {
  try {
    return await Word.run(async (context) => {
      console.log("Starting link removal process...");

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("text,style,styleBuiltIn,alignment");
      await context.sync();

      // Build paragraph metadata and find TOC section to skip
      const metas = buildParagraphMeta(paragraphs);
      const tocRange = findTOCSectionRange(metas);
      const referenceStartIndex = deleteAll ? -1 : findReferenceStartIndex(metas);
      console.log(`TOC section range: ${tocRange.startIndex} to ${tocRange.endIndex}`);
      console.log(`Reference section start index: ${referenceStartIndex}`);

      // Regex to find URL-like text, using word boundaries to correctly handle trailing punctuation.
      const urlRegex = /\b((https?:\/\/)?[\w.-]+(?:\.[\w.-]+)+)\b/g;
      let linksRemovedCount = 0;

      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        const meta = metas[i];

        if (!deleteAll && isInReferenceSection(i, referenceStartIndex)) {
          console.log(`Skipping reference paragraph ${i + 1}`);
          continue;
        }

        // Skip TOC section paragraphs
        if (tocRange.startIndex !== -1 && i >= tocRange.startIndex && i <= tocRange.endIndex) {
          console.log(`Skipping TOC paragraph ${i + 1}`);
          continue;
        }

        // Skip individual TOC-like entries
        if (looksLikeTOC(meta.text)) {
          console.log(`Skipping TOC-like entry at paragraph ${i + 1}`);
          continue;
        }

        const originalText = paragraph.text;
        const matches = originalText.match(urlRegex);

        if (matches) {
          // Replace URL-like text, but skip purely numeric patterns like "2.1" or "3.2.1"
          // that look like section numbering, not URLs.
          let realRemoved = 0;
          const newText = originalText
            .replace(urlRegex, (match) => {
              if (/^[\d.]+$/.test(match)) return match; // preserve section numbers
              realRemoved++;
              return "";
            })
            .replace(/[ \t]+([.,;])/g, "$1");
          linksRemovedCount += realRemoved;
          if (realRemoved > 0) {
            // Replace only paragraph content (not paragraph mark) to preserve boundaries.
            paragraph.getRange(Word.RangeLocation.content).insertText(newText, Word.InsertLocation.replace);
          }
        }
      }

      await context.sync();

      const successMessage = `Removed ${linksRemovedCount} URL-like text snippets.`;
      console.log(successMessage);
      return successMessage;
    });
  } catch (error) {
    console.error("Error in removeLinks:", error);
    throw new Error(`Error removing links: ${error.message}`);
  }
}

/**
 * Removes weird number patterns from the document.
 * Handles patterns like:
 * - 【400489077423502†L40-L67】
 * - [288914753644591†L299-L356]
 *
 * @returns {Promise<string>} A message indicating the number of instances removed.
 */
export async function removeWeirdNumbers(): Promise<string> {
  try {
    return await Word.run(async (context) => {
      console.log("Starting weird number removal process...");

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("text,style,styleBuiltIn,alignment");
      await context.sync();

      // Build paragraph metadata and find TOC section to skip
      const metas = buildParagraphMeta(paragraphs);
      const tocRange = findTOCSectionRange(metas);
      const referenceStartIndex = findReferenceStartIndex(metas);
      console.log(`TOC section range: ${tocRange.startIndex} to ${tocRange.endIndex}`);
      console.log(`Reference section start index: ${referenceStartIndex}`);

      // Regex to find patterns like 【...】 or [...] with the specified format
      const weirdNumberPattern = /[【[]\d+.*?[†+t].*?[】\]]\S*/g;
      let totalRemoved = 0;

      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        const meta = metas[i];

        if (isInReferenceSection(i, referenceStartIndex)) {
          console.log(`Skipping reference paragraph ${i + 1}`);
          continue;
        }

        // Skip TOC section paragraphs
        if (tocRange.startIndex !== -1 && i >= tocRange.startIndex && i <= tocRange.endIndex) {
          console.log(`Skipping TOC paragraph ${i + 1}`);
          continue;
        }

        // Skip individual TOC-like entries
        if (looksLikeTOC(meta.text)) {
          console.log(`Skipping TOC-like entry at paragraph ${i + 1}`);
          continue;
        }

        const originalText = paragraph.text;
        const matches = originalText.match(weirdNumberPattern);

        if (matches && matches.length > 0) {
          totalRemoved += matches.length;
          // Replace weird markers and normalize repeated spaces/tabs only.
          // Do not normalize newlines/vertical tabs to avoid visual paragraph merges.
          const newText = originalText.replace(weirdNumberPattern, "").replace(/[ \t]{2,}/g, " ");
          // Replace only paragraph content (not paragraph mark) to preserve boundaries.
          paragraph.getRange(Word.RangeLocation.content).insertText(newText, Word.InsertLocation.replace);
        }
      }

      await context.sync();

      const successMessage = `Removed ${totalRemoved} weird number instances.`;
      console.log(successMessage);
      return successMessage;
    });
  } catch (error) {
    console.error("Error in removeWeirdNumbers:", error);
    throw new Error(`Error removing weird numbers: ${error.message}`);
  }
}

/**
 * Removes bold formatting from body paragraphs while preserving headings or titles.
 */
export async function normalizeBodyBold(): Promise<string> {
  try {
    return await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("text,style,styleBuiltIn,alignment");
      await context.sync();

      const metas = buildParagraphMeta(paragraphs);
      const tocRange = findTOCSectionRange(metas);
      const referenceStartIndex = findReferenceStartIndex(metas);
      console.log(`TOC section range: ${tocRange.startIndex} to ${tocRange.endIndex}`);
      console.log(`Reference section start index: ${referenceStartIndex}`);
      let updatedCount = 0;

      for (const meta of metas) {
        if (!meta.text) continue;
        if (isHeadingOrTitle(meta)) continue;

        if (isInReferenceSection(meta.index, referenceStartIndex)) {
          console.log(`Skipping reference paragraph ${meta.index + 1}`);
          continue;
        }

        // Skip TOC section paragraphs
        if (tocRange.startIndex !== -1 && meta.index >= tocRange.startIndex && meta.index <= tocRange.endIndex) {
          console.log(`Skipping TOC paragraph ${meta.index + 1}`);
          continue;
        }

        // Skip individual TOC-like entries
        if (looksLikeTOC(meta.text)) {
          console.log(`Skipping TOC-like entry at paragraph ${meta.index + 1}`);
          continue;
        }

        const paragraph = paragraphs.items[meta.index];
        paragraph.font.bold = false;
        updatedCount++;
      }

      await context.sync();
      const result = `Removed bold formatting from ${updatedCount} paragraph${updatedCount === 1 ? "" : "s"}.`;
      console.log(result);
      return result;
    });
  } catch (error) {
    console.error("Error in normalizeBodyBold:", error);
    throw new Error(`Error normalizing bold text: ${error.message}`);
  }
}

/**
 * Paraphrase all body paragraphs in the document using the local API
 */
interface ParaphraseParagraphMeta {
  id: string;
  text: string;
}

const PARAPHRASE_DELIMITER = "qbpdelim123";
const ZERO_WIDTH_TEXT_RE = /[\u200B-\u200D\uFEFF]/g;
type SinglePassMode = "standard" | "ludicrous";

function stripParaphraseDelimiterToken(text: string): string {
  return (text || "")
    .replace(ZERO_WIDTH_TEXT_RE, "")
    .replace(new RegExp(PARAPHRASE_DELIMITER, "ig"), " ")
    .replace(/\s+([,.;:!?])/g, "$1")
    .replace(/\s{2,}/g, " ")
    .trim();
}

export function sanitizeParaphraseOutputText(text: string): string {
  return stripParaphraseDelimiterToken(text);
}

export async function paraphraseDocument(): Promise<ParaphraseResult> {
  try {
    return await Word.run(async (context) => {
      console.log("Starting document paraphrase (Simple + Short mode with batch API)...");
      const allWarnings: string[] = [];

      // Health check (non-blocking)
      const healthResult = await checkServiceHealth();
      if (healthResult.warnings.length > 0) {
        console.warn("Health check warnings:", healthResult.warnings);
        allWarnings.push(...healthResult.warnings);
      }

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/text, items/outlineLevel, items/uniqueLocalId");
      await context.sync();

      // First pass: identify reference section
      let referenceStartIndex = -1;
      for (let i = paragraphs.items.length - 1; i >= 0; i--) {
        const text = paragraphs.items[i].text ? paragraphs.items[i].text.trim() : "";
        if (REFERENCE_HEADERS.some((regex) => regex.test(text))) {
          referenceStartIndex = i;
          console.log(`Found reference section at index ${i}: "${text}"`);
          break;
        }
      }

      const metas: ParaphraseParagraphMeta[] = [];

      for (let i = 0; i < paragraphs.items.length; i++) {
        const p = paragraphs.items[i];
        const text = p.text ? p.text.trim() : "";
        if (!text) continue;

        // Skip paragraphs in or after reference section
        if (referenceStartIndex !== -1 && i >= referenceStartIndex) {
          console.log(`Skipping paragraph in reference section (index: ${i}): "${text.substring(0, 50)}..."`);
          continue;
        }

        // Skip very short paragraphs (likely titles/headings)
        // 15+ words is a good threshold to identify real body content
        const wordCount = text.split(/\s+/).filter(Boolean).length;
        if (wordCount < 15) {
          console.log(`Skipping short paragraph (${wordCount} words): "${text}"`);
          continue;
        }

        // Skip main headings (outlineLevel 1, 2, 3 are Heading 1, 2, 3)
        // Levels 4-9 and 10 (body) are allowed if they have enough words
        const outlineLevel = typeof p.outlineLevel === "number" ? p.outlineLevel : Number(p.outlineLevel);
        if (outlineLevel >= 1 && outlineLevel <= 3) {
          console.log(`Skipping heading (outlineLevel: ${outlineLevel}): "${text.substring(0, 50)}..."`);
          continue;
        }

        // Paragraph has enough words to be considered body content
        console.log(
          `Including paragraph (${wordCount} words, outlineLevel: ${p.outlineLevel}): "${text.substring(0, 50)}..."`
        );

        metas.push({
          id: p.uniqueLocalId,
          text: p.text,
        });
      }

      if (metas.length === 0) {
        return { message: "No body paragraphs found to paraphrase.", warnings: allWarnings };
      }

      console.log(`Found ${metas.length} body paragraphs to paraphrase.`);

      // Calculate total word count and capture original preview
      const originalWordCount = metas.reduce((sum, meta) => {
        return sum + meta.text.split(/\s+/).filter(Boolean).length;
      }, 0);
      const originalPreview = metas.length > 0 ? metas[0].text.substring(0, 50) + "..." : "";
      console.log(`Total word count: ${originalWordCount}`);

      const selection = chooseAdaptiveAccountCount(
        originalWordCount,
        "dual",
        healthResult.statusData
      );
      const numAccounts = selection.count;
      console.log(
        `Using ${numAccounts} account(s) for processing via batch API (mode=dual, est=${selection.estimatedSeconds.toFixed(
          1
        )}s, accounts=${selection.accounts.join("/")}, capacity=${Math.round(selection.effectiveCapacity)})`
      );

      // Build payloads helper
      const buildPayload = (metaArray: ParaphraseParagraphMeta[]) => {
        const chunks: string[] = [];
        for (const meta of metaArray) {
          chunks.push(PARAPHRASE_DELIMITER);
          chunks.push(meta.text);
        }
        return chunks.join("\n\n");
      };

      // Split metas across accounts
      const chunkSize = Math.ceil(metas.length / numAccounts);
      const chunks: ParaphraseParagraphMeta[][] = [];
      for (let i = 0; i < numAccounts; i++) {
        const start = i * chunkSize;
        const end = Math.min(start + chunkSize, metas.length);
        chunks.push(metas.slice(start, end));
      }

      // Build batch request payload
      const batchPayload: Record<string, string> = { mode: "dual" };
      const usedAccounts: AccountKey[] = [];

      chunks.forEach((chunk, idx) => {
        if (chunk.length > 0) {
          const accountKey = ACCOUNT_KEYS[idx];
          batchPayload[accountKey] = buildPayload(chunk);
          usedAccounts.push(accountKey);
          console.log(`Account ${accountKey}: ${chunk.length} paragraphs`);
        }
      });

      // Send single batch request
      console.log("Sending batch request to service...");
      const response = await fetch(BATCH_API_URL, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(batchPayload),
      });

      if (!response.ok) {
        throw new Error(`Batch API request failed with status ${response.status}`);
      }

      console.log("Batch response received, parsing...");
      const data: BatchApiResponse = await response.json();

      // Collect warnings from response
      const batchWarnings = collectBatchWarnings(data, usedAccounts);
      allWarnings.push(...batchWarnings);

      // Extract paraphrased text from each account (using secondMode for dual mode)
      const paraphrasedChunks: string[] = [];
      for (const key of usedAccounts) {
        const result = data[key];
        if (!result) {
          throw new Error(`No response received for account ${key}`);
        }
        // For dual mode, we use secondMode (the shortened version after Simple + Shorten)
        const paraphrasedText = result.secondMode;
        if (!paraphrasedText) {
          // If secondMode is missing but we have an error, it's a complete failure
          if (result.error) {
            throw new Error(`Account ${key} failed: ${result.error}`);
          }
          throw new Error(`No paraphrased text received from account ${key}`);
        }
        paraphrasedChunks.push(paraphrasedText);
      }

      // Parse all responses
      const parseResponse = (text: string, expectedCount: number) => {
        // 1. Initial split by delimiter (relaxed regex without \\b to catch cases where delimiter is attached to words)
        let parts = text
          .split(new RegExp(`${PARAPHRASE_DELIMITER}`, "i"))
          .map((x: string) => sanitizeParaphraseOutputText(x))
          .filter((x: string) => x.length > 0);

        // 2. Recovery Strategy: If we have fewer parts than expected, check for merged paragraphs
        // separated by double newlines.
        if (parts.length < expectedCount) {
          console.warn(
            `Mismatch detected (Expected ${expectedCount}, got ${parts.length}). Attempting to recover merged paragraphs via line breaks...`
          );

          const recoveredParts: string[] = [];
          for (const part of parts) {
            // Check if this part might actually be multiple paragraphs
            // Heuristic: If it contains double newlines, it might be merged instances
            if (part.includes("\n\n")) {
              const subParts = part
                .split(/\n\n+/) // Split by 2 or more newlines
                .map((p) => sanitizeParaphraseOutputText(p))
                .filter((p) => p.length > 0);
              recoveredParts.push(...subParts);
            } else {
              const sanitizedPart = sanitizeParaphraseOutputText(part);
              if (sanitizedPart) {
                recoveredParts.push(sanitizedPart);
              }
            }
          }

          // Only use recovered parts if it gets us closer to or exactly the expected count
          if (recoveredParts.length === expectedCount) {
            console.log("Successfully recovered merged paragraphs!");
            return recoveredParts;
          } else if (Math.abs(recoveredParts.length - expectedCount) < Math.abs(parts.length - expectedCount)) {
            console.log(`Partial recovery: now have ${recoveredParts.length} parts.`);
            return recoveredParts;
          }
        }

        return parts;
      };

      const allParsedChunks = chunks.map((chunk, idx) => parseResponse(paraphrasedChunks[idx], chunk.length));

      // Log received counts and track mismatches
      let totalMismatch = 0;
      let mismatchDetails: string[] = [];

      allParsedChunks.forEach((parsed, idx) => {
        console.log(`Received ${parsed.length} parts from account ${usedAccounts[idx]}`);
        const sent = chunks[idx].length;
        const received = parsed.length;
        if (received !== sent) {
          const diff = sent - received;
          totalMismatch += diff;
          mismatchDetails.push(`${usedAccounts[idx]}: sent ${sent}, received ${received}`);
          console.warn(`Account ${usedAccounts[idx]} mismatch: sent ${sent}, received ${received} (${diff} missing)`);
        }
      });

      // If there's a mismatch, add a warning but continue with what we received
      if (totalMismatch > 0) {
        const warningMsg = `Partial paraphrase: ${totalMismatch} paragraph(s) could not be processed and were left unchanged. Details: ${mismatchDetails.join("; ")}`;
        console.warn(warningMsg);
        allWarnings.push(warningMsg);
      }

      // Combine results in order
      const allParts = allParsedChunks.flat();
      console.log(`Combined total: ${allParts.length} paraphrased paragraphs (expected ${metas.length})`);

      // Re-fetch paragraphs to ensure validity and apply changes
      const paragraphsToUpdate = context.document.body.paragraphs;
      paragraphsToUpdate.load("items/uniqueLocalId");
      await context.sync();

      const paragraphById = new Map<string, Word.Paragraph>();
      paragraphsToUpdate.items.forEach((p) => {
        paragraphById.set(p.uniqueLocalId, p);
      });

      // Collect original and new texts for comparison
      const originalTexts: string[] = [];
      const newTexts: string[] = [];

      let updatedCount = 0;
      let skippedCount = 0;
      let newPreview = "";
      const skippedPreviews: string[] = [];

      for (let i = 0; i < metas.length; i++) {
        const meta = metas[i];
        const p = paragraphById.get(meta.id);

        // Check if we have a paraphrased version for this paragraph
        if (i < allParts.length) {
          const newText = sanitizeParaphraseOutputText(allParts[i]);

          // Store texts for comparison
          originalTexts.push(meta.text);
          newTexts.push(newText);

          // Capture first paragraph as new preview
          if (i === 0) {
            newPreview = newText.substring(0, 50) + "...";
          }

          if (p) {
            p.insertText(newText, Word.InsertLocation.replace);
            p.font.bold = false;
            updatedCount++;
            console.log(`Updated paragraph ${i + 1}/${metas.length}`);
          } else {
            console.warn(`Could not find paragraph with ID ${meta.id}`);
          }
        } else {
          // No paraphrased version available - leave paragraph unchanged
          skippedCount++;
          // Capture first 8 words of skipped paragraph for warning
          const firstWords = meta.text.split(/\s+/).slice(0, 8).join(" ");
          skippedPreviews.push(`"${firstWords}..."`);
          console.log(`Skipped paragraph ${i + 1}/${metas.length} (no paraphrased version received): ${firstWords}`);
          // Still include original in metrics for accurate comparison
          originalTexts.push(meta.text);
          newTexts.push(meta.text);
        }
      }

      await context.sync();
      console.log(
        `Paraphrase complete. Updated ${updatedCount}/${metas.length} paragraphs, skipped ${skippedCount} using batch API.`
      );

      // Calculate real change metrics by comparing actual words
      const changeMetrics = calculateTextChangeMetrics(originalTexts, newTexts);

      // Build result message
      let resultMessage = `Successfully paraphrased ${updatedCount} body paragraphs (batch mode).`;
      if (skippedCount > 0) {
        resultMessage = `Paraphrased ${updatedCount} of ${metas.length} paragraphs. ${skippedCount} paragraph(s) left unchanged due to processing limits.`;
        // Add warning with previews of skipped paragraphs
        const skippedWarning = `Unchanged paragraphs: ${skippedPreviews.join(", ")}`;
        allWarnings.push(skippedWarning);
      }

      return {
        message: resultMessage,
        warnings: allWarnings,
        metrics: {
          originalWordCount: changeMetrics.totalOriginalWords,
          newWordCount: changeMetrics.totalNewWords,
          wordsChanged: changeMetrics.wordsChanged,
          wordChangePercent: changeMetrics.changePercent,
          reusedWords: changeMetrics.reusedWords,
          reusePercent: changeMetrics.reusePercent,
          addedWords: changeMetrics.addedWords,
          originalPreview,
          newPreview,
        },
      };
    });
  } catch (error) {
    console.error("Error in paraphraseDocument:", error);
    throw new Error(`Error paraphrasing document: ${error.message}`);
  }
}

/**
 * Paraphrase selected text using the batch API (Simple + Short mode / dual mode)
 */
export async function paraphraseSelectedText(): Promise<ParaphraseResult> {
  try {
    return await Word.run(async (context) => {
      const allWarnings: string[] = [];

      // Health check (non-blocking)
      const healthResult = await checkServiceHealth();
      if (healthResult.warnings.length > 0) {
        console.warn("Health check warnings:", healthResult.warnings);
        allWarnings.push(...healthResult.warnings);
      }

      // Get the selected range
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      // Get the selected text content
      const selectedText = selection.text.trim();

      if (!selectedText) {
        throw new Error("No text selected");
      }

      // Call the batch paraphrase API with dual mode
      const response = await fetch(BATCH_API_URL, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ acc1: selectedText, mode: "dual" }),
      });

      if (!response.ok) {
        throw new Error(`Batch API request failed with status ${response.status}`);
      }

      const data: BatchApiResponse = await response.json();

      // Collect warnings
      const batchWarnings = collectBatchWarnings(data, ["acc1"]);
      allWarnings.push(...batchWarnings);

      const result = data.acc1;
      if (!result) {
        throw new Error("No response received from batch API");
      }

      // For dual mode, use secondMode (the shortened result)
      const paraphrasedText = result.secondMode;

      if (!paraphrasedText) {
        if (result.error) {
          throw new Error(`Paraphrase failed: ${result.error}`);
        }
        throw new Error("Invalid response from paraphrase API");
      }

      const safeParaphrasedText = sanitizeParaphraseOutputText(paraphrasedText);

      // Replace the selected text with the paraphrased text
      selection.insertText(safeParaphrasedText, Word.InsertLocation.replace);
      selection.font.bold = false;
      await context.sync();

      return { message: "Text paraphrased successfully", warnings: allWarnings };
    });
  } catch (error) {
    console.error("Error in paraphraseSelectedText:", error);
    throw new Error(`Error paraphrasing text: ${error.message}`);
  }
}

/**
 * Paraphrase selected text using single-pass modes (standard/ludicrous)
 */
async function paraphraseSelectedTextSinglePass(mode: SinglePassMode): Promise<ParaphraseResult> {
  try {
    return await Word.run(async (context) => {
      const allWarnings: string[] = [];

      // Health check (non-blocking)
      const healthResult = await checkServiceHealth();
      if (healthResult.warnings.length > 0) {
        console.warn("Health check warnings:", healthResult.warnings);
        allWarnings.push(...healthResult.warnings);
      }

      // Get the selected range
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      // Get the selected text content
      const selectedText = selection.text.trim();

      if (!selectedText) {
        throw new Error("No text selected");
      }

      // Call the batch paraphrase API with selected single-pass mode
      const response = await fetch(BATCH_API_URL, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ acc1: selectedText, mode }),
      });

      if (!response.ok) {
        throw new Error(`Batch API request failed with status ${response.status}`);
      }

      const data: BatchApiResponse = await response.json();

      // Collect warnings
      const batchWarnings = collectBatchWarnings(data, ["acc1"]);
      allWarnings.push(...batchWarnings);

      const result = data.acc1;
      if (!result) {
        throw new Error("No response received from batch API");
      }

      // Single-pass modes return `result`.
      const paraphrasedText = result.result;

      if (!paraphrasedText) {
        if (result.error) {
          throw new Error(`Paraphrase failed: ${result.error}`);
        }
        throw new Error("Invalid response from paraphrase API");
      }

      const safeParaphrasedText = sanitizeParaphraseOutputText(paraphrasedText);

      // Replace the selected text with the paraphrased text
      selection.insertText(safeParaphrasedText, Word.InsertLocation.replace);
      selection.font.bold = false;
      await context.sync();

      return {
        message:
          mode === "standard"
            ? "Text paraphrased successfully (Standard mode)"
            : "Text paraphrased successfully (Ludicrous mode)",
        warnings: allWarnings,
      };
    });
  } catch (error) {
    console.error(`Error in paraphraseSelectedTextSinglePass (${mode}):`, error);
    throw new Error(`Error paraphrasing text: ${error.message}`);
  }
}

export async function paraphraseSelectedTextStandard(): Promise<ParaphraseResult> {
  return paraphraseSelectedTextSinglePass("standard");
}

export async function paraphraseSelectedTextLudicrous(): Promise<ParaphraseResult> {
  return paraphraseSelectedTextSinglePass("ludicrous");
}

/**
 * Paraphrase all body paragraphs in the document using single-pass modes (standard/ludicrous)
 */
async function paraphraseDocumentSinglePass(mode: SinglePassMode): Promise<ParaphraseResult> {
  try {
    return await Word.run(async (context) => {
      console.log(`Starting document paraphrase (${mode} mode with batch API)...`);
      const allWarnings: string[] = [];

      // Health check (non-blocking)
      const healthResult = await checkServiceHealth();
      if (healthResult.warnings.length > 0) {
        console.warn("Health check warnings:", healthResult.warnings);
        allWarnings.push(...healthResult.warnings);
      }

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/text, items/outlineLevel, items/uniqueLocalId");
      await context.sync();

      // First pass: identify reference section
      let referenceStartIndex = -1;
      for (let i = paragraphs.items.length - 1; i >= 0; i--) {
        const text = paragraphs.items[i].text ? paragraphs.items[i].text.trim() : "";
        if (REFERENCE_HEADERS.some((regex) => regex.test(text))) {
          referenceStartIndex = i;
          console.log(`Found reference section at index ${i}: "${text}"`);
          break;
        }
      }

      const metas: ParaphraseParagraphMeta[] = [];

      for (let i = 0; i < paragraphs.items.length; i++) {
        const p = paragraphs.items[i];
        const text = p.text ? p.text.trim() : "";
        if (!text) continue;

        // Skip paragraphs in or after reference section
        if (referenceStartIndex !== -1 && i >= referenceStartIndex) {
          console.log(`Skipping paragraph in reference section (index: ${i}): "${text.substring(0, 50)}..."`);
          continue;
        }

        // Skip very short paragraphs (likely titles/headings)
        // 15+ words is a good threshold to identify real body content
        const wordCount = text.split(/\s+/).filter(Boolean).length;
        if (wordCount < 15) {
          console.log(`Skipping short paragraph (${wordCount} words): "${text}"`);
          continue;
        }

        // Skip main headings (outlineLevel 1, 2, 3 are Heading 1, 2, 3)
        // Levels 4-9 and 10 (body) are allowed if they have enough words
        const outlineLevel = typeof p.outlineLevel === "number" ? p.outlineLevel : Number(p.outlineLevel);
        if (outlineLevel >= 1 && outlineLevel <= 3) {
          console.log(`Skipping heading (outlineLevel: ${outlineLevel}): "${text.substring(0, 50)}..."`);
          continue;
        }

        // Paragraph has enough words to be considered body content
        console.log(
          `Including paragraph (${wordCount} words, outlineLevel: ${p.outlineLevel}): "${text.substring(0, 50)}..."`
        );

        metas.push({
          id: p.uniqueLocalId,
          text: p.text,
        });
      }

      if (metas.length === 0) {
        return { message: "No body paragraphs found to paraphrase.", warnings: allWarnings };
      }

      console.log(`Found ${metas.length} body paragraphs to paraphrase.`);

      // Calculate total word count and capture original preview
      const originalWordCount = metas.reduce((sum, meta) => {
        return sum + meta.text.split(/\s+/).filter(Boolean).length;
      }, 0);
      const originalPreview = metas.length > 0 ? metas[0].text.substring(0, 50) + "..." : "";
      console.log(`Total word count: ${originalWordCount}`);

      const selection = chooseAdaptiveAccountCount(
        originalWordCount,
        mode,
        healthResult.statusData
      );
      const numAccounts = selection.count;
      console.log(
        `Using ${numAccounts} account(s) for processing via batch API (mode=${mode}, est=${selection.estimatedSeconds.toFixed(
          1
        )}s, accounts=${selection.accounts.join("/")}, capacity=${Math.round(selection.effectiveCapacity)})`
      );

      // Build payloads helper
      const buildPayload = (metaArray: ParaphraseParagraphMeta[]) => {
        const chunks: string[] = [];
        for (const meta of metaArray) {
          chunks.push(PARAPHRASE_DELIMITER);
          chunks.push(meta.text);
        }
        return chunks.join("\n\n");
      };

      // Split metas across accounts
      const chunkSize = Math.ceil(metas.length / numAccounts);
      const chunks: ParaphraseParagraphMeta[][] = [];
      for (let i = 0; i < numAccounts; i++) {
        const start = i * chunkSize;
        const end = Math.min(start + chunkSize, metas.length);
        chunks.push(metas.slice(start, end));
      }

      // Build batch request payload
      const batchPayload: Record<string, string> = { mode };
      const usedAccounts: AccountKey[] = [];

      chunks.forEach((chunk, idx) => {
        if (chunk.length > 0) {
          const accountKey = ACCOUNT_KEYS[idx];
          batchPayload[accountKey] = buildPayload(chunk);
          usedAccounts.push(accountKey);
          console.log(`Account ${accountKey}: ${chunk.length} paragraphs`);
        }
      });

      // Send single batch request
      console.log("Sending batch request to service...");
      const response = await fetch(BATCH_API_URL, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(batchPayload),
      });

      if (!response.ok) {
        throw new Error(`Batch API request failed with status ${response.status}`);
      }

      console.log("Batch response received, parsing...");
      const data: BatchApiResponse = await response.json();

      // Collect warnings from response
      const batchWarnings = collectBatchWarnings(data, usedAccounts);
      allWarnings.push(...batchWarnings);

      // Extract paraphrased text from each account (single-pass modes use `result`)
      const paraphrasedChunks: string[] = [];
      for (const key of usedAccounts) {
        const accountResult = data[key];
        if (!accountResult) {
          throw new Error(`No response received for account ${key}`);
        }
        // Single-pass modes use `result`.
        const paraphrasedText = accountResult.result;
        if (!paraphrasedText) {
          // If result is missing but we have an error, it's a complete failure
          if (accountResult.error) {
            throw new Error(`Account ${key} failed: ${accountResult.error}`);
          }
          throw new Error(`No paraphrased text received from account ${key}`);
        }
        paraphrasedChunks.push(paraphrasedText);
      }

      // Parse all responses
      const parseResponse = (text: string) => {
        return text
          .split(new RegExp(`${PARAPHRASE_DELIMITER}`, "i"))
          .map((x: string) => sanitizeParaphraseOutputText(x))
          .filter((x: string) => x.length > 0);
      };

      const allParsedChunks = paraphrasedChunks.map((chunk) => parseResponse(chunk));

      // Log received counts and track mismatches
      let totalMismatch = 0;
      let mismatchDetails: string[] = [];

      allParsedChunks.forEach((parsed, idx) => {
        console.log(`Received ${parsed.length} parts from account ${usedAccounts[idx]}`);
        const sent = chunks[idx].length;
        const received = parsed.length;
        if (received !== sent) {
          const diff = sent - received;
          totalMismatch += diff;
          mismatchDetails.push(`${usedAccounts[idx]}: sent ${sent}, received ${received}`);
          console.warn(`Account ${usedAccounts[idx]} mismatch: sent ${sent}, received ${received} (${diff} missing)`);
        }
      });

      // If there's a mismatch, add a warning but continue with what we received
      if (totalMismatch > 0) {
        const warningMsg = `Partial paraphrase: ${totalMismatch} paragraph(s) could not be processed and were left unchanged. Details: ${mismatchDetails.join("; ")}`;
        console.warn(warningMsg);
        allWarnings.push(warningMsg);
      }

      // Combine results in order
      const allParts = allParsedChunks.flat();
      console.log(`Combined total: ${allParts.length} paraphrased paragraphs (expected ${metas.length})`);

      // Re-fetch paragraphs to ensure validity and apply changes
      const paragraphsToUpdate = context.document.body.paragraphs;
      paragraphsToUpdate.load("items/uniqueLocalId");
      await context.sync();

      const paragraphById = new Map<string, Word.Paragraph>();
      paragraphsToUpdate.items.forEach((p) => {
        paragraphById.set(p.uniqueLocalId, p);
      });

      // Collect original and new texts for comparison
      const originalTexts: string[] = [];
      const newTexts: string[] = [];

      let updatedCount = 0;
      let skippedCount = 0;
      let newPreview = "";
      const skippedPreviews: string[] = [];

      for (let i = 0; i < metas.length; i++) {
        const meta = metas[i];
        const p = paragraphById.get(meta.id);

        // Check if we have a paraphrased version for this paragraph
        if (i < allParts.length) {
          const newText = sanitizeParaphraseOutputText(allParts[i]);

          // Store texts for comparison
          originalTexts.push(meta.text);
          newTexts.push(newText);

          // Capture first paragraph as new preview
          if (i === 0) {
            newPreview = newText.substring(0, 50) + "...";
          }

          if (p) {
            p.insertText(newText, Word.InsertLocation.replace);
            p.font.bold = false;
            updatedCount++;
            console.log(`Updated paragraph ${i + 1}/${metas.length}`);
          } else {
            console.warn(`Could not find paragraph with ID ${meta.id}`);
          }
        } else {
          // No paraphrased version available - leave paragraph unchanged
          skippedCount++;
          // Capture first 8 words of skipped paragraph for warning
          const firstWords = meta.text.split(/\s+/).slice(0, 8).join(" ");
          skippedPreviews.push(`"${firstWords}..."`);
          console.log(`Skipped paragraph ${i + 1}/${metas.length} (no paraphrased version received): ${firstWords}`);
          // Still include original in metrics for accurate comparison
          originalTexts.push(meta.text);
          newTexts.push(meta.text);
        }
      }

      await context.sync();
      console.log(
        `Paraphrase complete. Updated ${updatedCount}/${metas.length} paragraphs, skipped ${skippedCount} using batch API.`
      );

      // Calculate real change metrics by comparing actual words
      const changeMetrics = calculateTextChangeMetrics(originalTexts, newTexts);

      // Build result message
      let resultMessage = `Successfully paraphrased ${updatedCount} body paragraphs (batch mode).`;
      if (skippedCount > 0) {
        resultMessage = `Paraphrased ${updatedCount} of ${metas.length} paragraphs. ${skippedCount} paragraph(s) left unchanged due to processing limits.`;
        // Add warning with previews of skipped paragraphs
        const skippedWarning = `Unchanged paragraphs: ${skippedPreviews.join(", ")}`;
        allWarnings.push(skippedWarning);
      }

      return {
        message: resultMessage,
        warnings: allWarnings,
        metrics: {
          originalWordCount: changeMetrics.totalOriginalWords,
          newWordCount: changeMetrics.totalNewWords,
          wordsChanged: changeMetrics.wordsChanged,
          wordChangePercent: changeMetrics.changePercent,
          reusedWords: changeMetrics.reusedWords,
          reusePercent: changeMetrics.reusePercent,
          addedWords: changeMetrics.addedWords,
          originalPreview,
          newPreview,
        },
      };
    });
  } catch (error) {
    console.error(`Error in paraphraseDocumentSinglePass (${mode}):`, error);
    throw new Error(`Error paraphrasing document: ${error.message}`);
  }
}

export async function paraphraseDocumentStandard(): Promise<ParaphraseResult> {
  return paraphraseDocumentSinglePass("standard");
}

export async function paraphraseDocumentLudicrous(): Promise<ParaphraseResult> {
  return paraphraseDocumentSinglePass("ludicrous");
}
