/* global Word console, fetch */
import { getFormattedReferences } from "./gemini";

const TRAILING_PUNCTUATION = "[-:;,.!?–—]*";

const REFERENCE_HEADERS = [
  new RegExp(`^\\s*references?(?:\\s+list)?(?:\\s+section)?\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*reference\\s+list\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*references\\s+list\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*bibliograph(?:y|ies)\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*list\\s+of\\s+references\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
];

const CONCLUSION_HEADERS = [
  new RegExp(`^\\s*conclusions?(?:\\s+section)?\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*concluding\\s+remarks\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*final\\s+thoughts\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*summary(?:\\s+and\\s+future\\s+work)?\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*closing\\s+remarks\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*conclusion\\s+and\\s+recommendations\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
];

type ParagraphMeta = {
  index: number;
  text: string;
  wordCount: number;
  style: string;
  styleBuiltIn: string;
  alignment: string;
};

function buildParagraphMeta(paragraphs: Word.ParagraphCollection): ParagraphMeta[] {
  return paragraphs.items.map((p, index) => {
    const rawText = p.text || "";
    const text = rawText.trim();
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
  const dotLeaderWithPage = /\.{5,}.*\d+\s*$/.test(trimmed);
  const hasTab = /\t/.test(trimmed);
  return dotLeaderWithPage || hasTab;
}

function findReferenceStartIndex(metas: ParagraphMeta[]): number {
  for (let i = metas.length - 1; i >= 0; i--) {
    if (REFERENCE_HEADERS.some((regex) => regex.test(metas[i].text))) {
      return i;
    }
  }
  return -1;
}

function findConclusionRange(
  metas: ParagraphMeta[],
  referenceStartIndex: number
): { conclusionHeadingIndex: number; conclusionEndIndex: number } {
  let conclusionHeadingIndex = -1;

  for (let i = referenceStartIndex === -1 ? metas.length - 1 : referenceStartIndex - 1; i >= 0; i--) {
    const txt = metas[i].text;
    if (CONCLUSION_HEADERS.some((regex) => regex.test(txt))) {
      conclusionHeadingIndex = i;
      break;
    }
  }

  let conclusionEndIndex = -1;

  if (conclusionHeadingIndex !== -1) {
    for (let i = conclusionHeadingIndex + 1; i < metas.length; i++) {
      if (REFERENCE_HEADERS.some((regex) => regex.test(metas[i].text))) {
        conclusionEndIndex = i;
        break;
      }

      if (isHeadingOrTitle(metas[i]) && !CONCLUSION_HEADERS.some((regex) => regex.test(metas[i].text))) {
        conclusionEndIndex = i;
        break;
      }
    }

    if (conclusionEndIndex === -1 && referenceStartIndex !== -1) {
      conclusionEndIndex = referenceStartIndex;
    }
  }

  return { conclusionHeadingIndex, conclusionEndIndex };
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

      const { conclusionHeadingIndex, conclusionEndIndex } = findConclusionRange(metas, referenceStartIndex);
      console.log("analyzeDocument => conclusion range", { conclusionHeadingIndex, conclusionEndIndex });

      // Filter out short paragraphs, TOC lines, and those ending with ":"
      const eligibleIndexes = metas
        .filter((meta) => {
          if (referenceStartIndex !== -1 && meta.index >= referenceStartIndex) {
            console.log("eligible => skip reference section", meta.index);
            return false;
          }

          if (
            conclusionHeadingIndex !== -1 &&
            conclusionEndIndex !== -1 &&
            meta.index > conclusionHeadingIndex &&
            meta.index < conclusionEndIndex
          ) {
            console.log("eligible => skip conclusion", meta.index);
            return false;
          }

          if (isHeadingOrTitle(meta)) {
            console.log("eligible => skip heading", meta.index, meta.text);
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
          console.log("eligible => removing first paragraph", firstNonEmpty.index);
          eligibleIndexes.splice(idx, 1);
        }
      }

      if (eligibleIndexes.length === 0) {
        console.warn("analyzeDocument => abort: no eligible paragraphs");
        return "No eligible paragraphs found for inserting references";
      }
      console.log("analyzeDocument => eligible count", eligibleIndexes.length);

      // If insertEveryOther is true, only use every other paragraph
      let targetIndexes = [...eligibleIndexes];
      if (insertEveryOther && targetIndexes.length > 0) {
        // Sort available indexes to ensure we're working with ordered paragraphs
        targetIndexes.sort((a, b) => a - b);
        // Filter to only include every other paragraph
        targetIndexes = targetIndexes.filter((_, i) => i % 2 === 0);
        console.log("analyzeDocument => insertEveryOther applied", targetIndexes.length);
      }

      // Shuffle available indexes
      const randomIndex = [...targetIndexes].sort(() => Math.random() - 0.5);
      console.log("analyzeDocument => random order", randomIndex);

      // Insert references
      const usedReferences = new Set<number>();

      // Modify paragraphs with references
      for (const index of randomIndex) {
        let referenceIndex: number;

        if (usedReferences.size < references.length) {
          const unusedReferences = Array.from(Array(references.length).keys()).filter((i) => !usedReferences.has(i));
          referenceIndex = unusedReferences[Math.floor(Math.random() * unusedReferences.length)];
          usedReferences.add(referenceIndex);
        } else {
          referenceIndex = Math.floor(Math.random() * references.length);
        }

        const paragraph = paragraphs.items[index];
        const text = paragraph.text.trim();
        console.log("analyzeDocument => inserting reference", { paragraphIndex: index, referenceIndex });

        if (text.endsWith(".")) {
          paragraph.insertText(text.slice(0, -1) + ` ${references[referenceIndex]}.`, Word.InsertLocation.replace);
        } else {
          paragraph.insertText(` ${references[referenceIndex]}.`, Word.InsertLocation.end);
        }
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
      paragraphs.load("text");
      await context.sync();

      console.log(`Found ${paragraphs.items.length} paragraphs to process`);

      // Updated citation patterns to be more precise
      const citationPatterns = [
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
      paragraphs.load("text");
      await context.sync();

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

      for (const paragraph of paragraphsToProcess) {
        const originalText = paragraph.text;
        const matches = originalText.match(urlRegex);

        if (matches) {
          linksRemovedCount += matches.length;
          // Replace URL-like text and clean up spaces before punctuation.
          const newText = originalText.replace(urlRegex, "").replace(/\s+([.,;])/g, "$1");
          paragraph.insertText(newText, Word.InsertLocation.replace);
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
      paragraphs.load("text");
      await context.sync();

      // Regex to find patterns like 【...】 or [...] with the specified format
      const weirdNumberPattern = /[【[]\d+.*?[†+t].*?[】\]]\S*/g;
      let totalRemoved = 0;

      for (const paragraph of paragraphs.items) {
        const originalText = paragraph.text;
        const matches = originalText.match(weirdNumberPattern);

        if (matches && matches.length > 0) {
          totalRemoved += matches.length;
          // Replace the weird numbers and clean up potential double spaces.
          const newText = originalText.replace(weirdNumberPattern, "").replace(/\s{2,}/g, " ");
          paragraph.insertText(newText, Word.InsertLocation.replace);
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
 * Paraphrase selected text using the local API
 */
export async function paraphraseSelectedText(): Promise<string> {
  try {
    return await Word.run(async (context) => {
      // Get the selected range
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      // Get the selected text content
      const selectedText = selection.text.trim();

      if (!selectedText) {
        throw new Error("No text selected");
      }

      // Call the paraphrase API
      const response = await fetch("https://analizeai.com/paraphrase", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ text: selectedText }),
      });

      if (!response.ok) {
        throw new Error(`API request failed with status ${response.status}`);
      }

      const data = await response.json();
      const paraphrasedText = data.secondMode;

      if (!paraphrasedText) {
        throw new Error("Invalid response from paraphrase API");
      }

      // Replace the selected text with the paraphrased text
      selection.insertText(paraphrasedText, Word.InsertLocation.replace);
      await context.sync();

      return "Text paraphrased successfully";
    });
  } catch (error) {
    console.error("Error in paraphraseSelectedText:", error);
    throw new Error(`Error paraphrasing text: ${error.message}`);
  }
}
