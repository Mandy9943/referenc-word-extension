/* global Word console, fetch, AbortController, setTimeout, clearTimeout */
import { getFormattedReferences } from "./gemini";

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
  wordChangePercent: number;
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
  status: string;
  acc1?: { status: string };
  acc2?: { status: string };
  acc3?: { status: string };
  ready?: boolean;
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
async function checkServiceHealth(): Promise<{ ready: boolean; warnings: string[] }> {
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
      return { ready: false, warnings };
    }

    const data: HealthCheckResponse = await response.json();

    // Check each account status
    for (const key of ACCOUNT_KEYS) {
      const accountStatus = data[key]?.status;
      if (accountStatus && accountStatus !== "ready") {
        warnings.push(`Account ${key} is ${accountStatus}`);
      }
    }

    return { ready: data.ready ?? true, warnings };
  } catch (error) {
    if (error.name === "AbortError") {
      warnings.push("Health check timed out (service may be slow)");
    } else {
      warnings.push(`Health check failed: ${error.message}`);
    }
    return { ready: false, warnings };
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

function isHeadingOrTitle(meta: ParagraphMeta): boolean {
  const s = meta.style;
  const sb = meta.styleBuiltIn;

  if (
    s.includes("heading") ||
    s.includes("title") ||
    s.includes("subtitle") ||
    sb.includes("heading") ||
    sb.includes("title") ||
    sb.includes("subtitle")
  ) {
    return true;
  }

  const text = meta.text;
  if (!text) return false;

  const hasTerminalPunctuation = /[.!?]$/.test(text);
  const isShortish = meta.wordCount > 0 && meta.wordCount <= 15;
  const isCentered = meta.alignment === "centered";

  const words = text.split(/\s+/).filter(Boolean);
  const capitalised = words.filter((w) => /^[A-Z]/.test(w)).length;
  const isTitleCase = words.length > 1 && capitalised / words.length > 0.6;

  if (!hasTerminalPunctuation && isShortish && (isCentered || isTitleCase)) {
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

function findReferenceStartIndex(metas: ParagraphMeta[]): number {
  for (let i = metas.length - 1; i >= 0; i--) {
    if (REFERENCE_HEADERS.some((regex) => regex.test(metas[i].text))) {
      return i;
    }
  }
  return -1;
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

  // Limit per paragraph: 1–3 sentences max, chosen randomly
  if (slots.length <= 1) return slots;

  shuffleInPlace(slots);
  const maxForParagraph = Math.min(3, slots.length);
  const targetCount = 1 + Math.floor(Math.random() * maxForParagraph);
  return slots.slice(0, targetCount);
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
      const eligibleIndexes = metas
        .filter((meta) => {
          if (referenceStartIndex !== -1 && meta.index >= referenceStartIndex) {
            return false;
          }

          const insideExcludedSection = excludedSections.some((section) => {
            if (section.headingIndex === -1 || section.endIndex === -1) {
              return false;
            }
            return meta.index > section.headingIndex && meta.index < section.endIndex;
          });

          if (insideExcludedSection) {
            return false;
          }

          if (isHeadingOrTitle(meta)) {
            return false;
          }

          if (meta.wordCount < 11) {
            return false;
          }

          if (looksLikeTOC(meta.text)) {
            return false;
          }

          if (meta.text.trim().endsWith(":")) {
            return false;
          }

          return true;
        })
        .map((meta) => meta.index);

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
        const ranges = paragraph.getTextRanges([".", "?", "!"], true);
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

        const currentText = sentenceRange.text || "";
        const citation = references[referenceIndex];
        const newText = appendCitationAtSentenceEnd(currentText, citation);

        console.log("analyzeDocument => inserting reference", {
          paragraphIndex: slot.paragraphIndex,
          sentenceIndex: slot.sentenceIndex,
          referenceIndex,
        });

        sentenceRange.insertText(newText, Word.InsertLocation.replace);
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
      console.log(`TOC section range: ${tocRange.startIndex} to ${tocRange.endIndex}`);

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

            // Remove all matches and clean up extra spaces before periods
            text = text.replace(pattern, "").replace(/\s+\./g, ".");
            hadMatch = true;
            totalRemoved += matches.length;
          }
        }

        // Only update paragraph if we found and removed citations
        if (hadMatch) {
          // Replace the paragraph's content
          paragraph.getRange().insertText(text, Word.InsertLocation.replace);
          paragraph.font.bold = false;
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
      console.log(`TOC section range: ${tocRange.startIndex} to ${tocRange.endIndex}`);

      let paragraphsToProcess = paragraphs.items;

      if (!deleteAll) {
        const paragraphTexts = paragraphs.items.map((p) => p.text);

        const referenceHeaders = [
          "Reference List",
          "References List",
          "References",
          "REFERENCES LIST",
          "REFERENCE LIST",
          "REFERENCES",
          "Bibliography",
          "BIBLIOGRAPHY",
        ];

        // @ts-ignore
        const lastReferenceListIndex = paragraphTexts.findLastIndex((p) =>
          referenceHeaders.some((header) => p.includes(header))
        );

        console.log(`Reference section starts at paragraph index: ${lastReferenceListIndex}`);

        if (lastReferenceListIndex !== -1) {
          paragraphsToProcess = paragraphs.items.slice(0, lastReferenceListIndex);
        }
      }

      // Regex to find URL-like text, using word boundaries to correctly handle trailing punctuation.
      const urlRegex = /\b((https?:\/\/)?[\w.-]+(?:\.[\w.-]+)+)\b/g;
      let linksRemovedCount = 0;

      for (let i = 0; i < paragraphsToProcess.length; i++) {
        const paragraph = paragraphsToProcess[i];
        const paragraphIndex = paragraphs.items.indexOf(paragraph);
        const meta = metas[paragraphIndex];

        // Skip TOC section paragraphs
        if (
          tocRange.startIndex !== -1 &&
          paragraphIndex >= tocRange.startIndex &&
          paragraphIndex <= tocRange.endIndex
        ) {
          console.log(`Skipping TOC paragraph ${paragraphIndex + 1}`);
          continue;
        }

        // Skip individual TOC-like entries
        if (looksLikeTOC(meta.text)) {
          console.log(`Skipping TOC-like entry at paragraph ${paragraphIndex + 1}`);
          continue;
        }

        const originalText = paragraph.text;
        const matches = originalText.match(urlRegex);

        if (matches) {
          linksRemovedCount += matches.length;
          // Replace URL-like text and clean up spaces before punctuation.
          const newText = originalText.replace(urlRegex, "").replace(/\s+([.,;])/g, "$1");
          paragraph.insertText(newText, Word.InsertLocation.replace);
          paragraph.font.bold = false;
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
      console.log(`TOC section range: ${tocRange.startIndex} to ${tocRange.endIndex}`);

      // Regex to find patterns like 【...】 or [...] with the specified format
      const weirdNumberPattern = /[【[]\d+.*?[†+t].*?[】\]]\S*/g;
      let totalRemoved = 0;

      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        const meta = metas[i];

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
          // Replace the weird numbers and clean up potential double spaces.
          const newText = originalText.replace(weirdNumberPattern, "").replace(/\s{2,}/g, " ");
          paragraph.insertText(newText, Word.InsertLocation.replace);
          paragraph.font.bold = false;
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
      console.log(`TOC section range: ${tocRange.startIndex} to ${tocRange.endIndex}`);
      let updatedCount = 0;

      for (const meta of metas) {
        if (!meta.text) continue;
        if (isHeadingOrTitle(meta)) continue;

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
        const wordCount = text.split(/\s+/).filter(Boolean).length;
        if (wordCount < 15) {
          console.log(`Skipping short paragraph (${wordCount} words): "${text}"`);
          continue;
        }

        // Skip headings/titles - only include body text paragraphs
        // In Word API, outlineLevel 10 = body text, other numbers (1-9) = heading levels
        const outlineLevel = typeof p.outlineLevel === "number" ? p.outlineLevel : Number(p.outlineLevel);
        const isBody = outlineLevel === 10;

        if (!isBody) {
          console.log(`Skipping non-body paragraph (outlineLevel: ${p.outlineLevel}): "${text.substring(0, 50)}..."`);
          continue;
        }

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

      // Determine number of accounts to use based on word count
      let numAccounts = 1;
      if (originalWordCount > 1500) {
        numAccounts = 3;
      } else if (originalWordCount >= 500) {
        numAccounts = 2;
      }

      console.log(`Using ${numAccounts} account(s) for processing via batch API`);

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
          .map((x: string) => x.trim())
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
                .map((p) => p.trim())
                .filter((p) => p.length > 0);
              recoveredParts.push(...subParts);
            } else {
              recoveredParts.push(part);
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

      // Log received counts
      allParsedChunks.forEach((parsed, idx) => {
        console.log(`Received ${parsed.length} parts from account ${usedAccounts[idx]}`);
      });

      // Validate counts
      allParsedChunks.forEach((parsed, idx) => {
        if (parsed.length !== chunks[idx].length) {
          console.error(`Account ${usedAccounts[idx]} mismatch: sent ${chunks[idx].length}, received ${parsed.length}`);
          throw new Error(
            `Account ${usedAccounts[idx]} paraphrase count mismatch. Sent ${chunks[idx].length}, received ${parsed.length}. Aborting to prevent data loss.`
          );
        }
      });

      // Combine results in order
      const allParts = allParsedChunks.flat();
      console.log(`Combined total: ${allParts.length} paraphrased paragraphs`);

      // Re-fetch paragraphs to ensure validity and apply changes
      const paragraphsToUpdate = context.document.body.paragraphs;
      paragraphsToUpdate.load("items/uniqueLocalId");
      await context.sync();

      const paragraphById = new Map<string, Word.Paragraph>();
      paragraphsToUpdate.items.forEach((p) => {
        paragraphById.set(p.uniqueLocalId, p);
      });

      let updatedCount = 0;
      let newWordCount = 0;
      let newPreview = "";
      for (let i = 0; i < metas.length; i++) {
        const meta = metas[i];
        const newText = allParts[i];
        const p = paragraphById.get(meta.id);

        // Track new word count
        newWordCount += newText.split(/\s+/).filter(Boolean).length;

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
      }

      await context.sync();
      console.log(`Paraphrase complete. Updated ${updatedCount}/${metas.length} paragraphs using batch API.`);

      // Calculate change metrics
      const wordChangePercent =
        originalWordCount > 0 ? Math.round((Math.abs(originalWordCount - newWordCount) / originalWordCount) * 100) : 0;

      return {
        message: `Successfully paraphrased ${updatedCount} body paragraphs (batch mode).`,
        warnings: allWarnings,
        metrics: {
          originalWordCount,
          newWordCount,
          wordChangePercent,
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

      // Replace the selected text with the paraphrased text
      selection.insertText(paraphrasedText, Word.InsertLocation.replace);
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
 * Paraphrase selected text using the batch API (Standard mode)
 */
export async function paraphraseSelectedTextStandard(): Promise<ParaphraseResult> {
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

      // Call the batch paraphrase API with standard mode
      const response = await fetch(BATCH_API_URL, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ acc1: selectedText, mode: "standard" }),
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

      // For standard mode, use result field
      const paraphrasedText = result.result;

      if (!paraphrasedText) {
        if (result.error) {
          throw new Error(`Paraphrase failed: ${result.error}`);
        }
        throw new Error("Invalid response from paraphrase API");
      }

      // Replace the selected text with the paraphrased text
      selection.insertText(paraphrasedText, Word.InsertLocation.replace);
      selection.font.bold = false;
      await context.sync();

      return { message: "Text paraphrased successfully", warnings: allWarnings };
    });
  } catch (error) {
    console.error("Error in paraphraseSelectedTextStandard:", error);
    throw new Error(`Error paraphrasing text: ${error.message}`);
  }
}

/**
 * Paraphrase all body paragraphs in the document using the batch API (Standard mode)
 */
export async function paraphraseDocumentStandard(): Promise<ParaphraseResult> {
  try {
    return await Word.run(async (context) => {
      console.log("Starting document paraphrase (Standard mode with batch API)...");
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
        const wordCount = text.split(/\s+/).filter(Boolean).length;
        if (wordCount < 15) {
          console.log(`Skipping short paragraph (${wordCount} words): "${text}"`);
          continue;
        }

        // Skip headings/titles - only include body text paragraphs
        // In Word API, outlineLevel 10 = body text, other numbers (1-9) = heading levels
        const outlineLevel = typeof p.outlineLevel === "number" ? p.outlineLevel : Number(p.outlineLevel);
        const isBody = outlineLevel === 10;

        if (!isBody) {
          console.log(`Skipping non-body paragraph (outlineLevel: ${p.outlineLevel}): "${text.substring(0, 50)}..."`);
          continue;
        }

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

      // Determine number of accounts to use based on word count
      let numAccounts = 1;
      if (originalWordCount > 1500) {
        numAccounts = 3;
      } else if (originalWordCount >= 500) {
        numAccounts = 2;
      }

      console.log(`Using ${numAccounts} account(s) for processing via batch API`);

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
      const batchPayload: Record<string, string> = { mode: "standard" };
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

      // Extract paraphrased text from each account (using result for standard mode)
      const paraphrasedChunks: string[] = [];
      for (const key of usedAccounts) {
        const accountResult = data[key];
        if (!accountResult) {
          throw new Error(`No response received for account ${key}`);
        }
        // For standard mode, we use result field
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
          .split(new RegExp(`\\b${PARAPHRASE_DELIMITER}\\b`, "i"))
          .map((x: string) => x.trim())
          .filter((x: string) => x.length > 0);
      };

      const allParsedChunks = paraphrasedChunks.map((chunk) => parseResponse(chunk));

      // Log received counts
      allParsedChunks.forEach((parsed, idx) => {
        console.log(`Received ${parsed.length} parts from account ${usedAccounts[idx]}`);
      });

      // Validate counts
      allParsedChunks.forEach((parsed, idx) => {
        if (parsed.length !== chunks[idx].length) {
          console.error(`Account ${usedAccounts[idx]} mismatch: sent ${chunks[idx].length}, received ${parsed.length}`);
          throw new Error(
            `Account ${usedAccounts[idx]} paraphrase count mismatch. Sent ${chunks[idx].length}, received ${parsed.length}. Aborting to prevent data loss.`
          );
        }
      });

      // Combine results in order
      const allParts = allParsedChunks.flat();
      console.log(`Combined total: ${allParts.length} paraphrased paragraphs`);

      // Re-fetch paragraphs to ensure validity and apply changes
      const paragraphsToUpdate = context.document.body.paragraphs;
      paragraphsToUpdate.load("items/uniqueLocalId");
      await context.sync();

      const paragraphById = new Map<string, Word.Paragraph>();
      paragraphsToUpdate.items.forEach((p) => {
        paragraphById.set(p.uniqueLocalId, p);
      });

      let updatedCount = 0;
      let newWordCount = 0;
      let newPreview = "";
      for (let i = 0; i < metas.length; i++) {
        const meta = metas[i];
        const newText = allParts[i];
        const p = paragraphById.get(meta.id);

        // Track new word count
        newWordCount += newText.split(/\s+/).filter(Boolean).length;

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
      }

      await context.sync();
      console.log(`Paraphrase complete. Updated ${updatedCount}/${metas.length} paragraphs using batch API.`);

      // Calculate change metrics
      const wordChangePercent =
        originalWordCount > 0 ? Math.round((Math.abs(originalWordCount - newWordCount) / originalWordCount) * 100) : 0;

      return {
        message: `Successfully paraphrased ${updatedCount} body paragraphs (batch mode).`,
        warnings: allWarnings,
        metrics: {
          originalWordCount,
          newWordCount,
          wordChangePercent,
          originalPreview,
          newPreview,
        },
      };
    });
  } catch (error) {
    console.error("Error in paraphraseDocumentStandard:", error);
    throw new Error(`Error paraphrasing document: ${error.message}`);
  }
}
