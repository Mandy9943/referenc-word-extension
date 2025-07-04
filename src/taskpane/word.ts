/* global Word console */
import { Anthropic } from "@anthropic-ai/sdk";
import { getFormattedReferences } from "./gemini";

// Add a cancellation flag and helper methods
let isHumanizeCancelled = false;

export function requestCancelHumanize() {
  isHumanizeCancelled = true;
}

export function resetHumanizeCancelState() {
  isHumanizeCancelled = false;
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
      // Get all paragraphs from the document
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("text");
      await context.sync();

      // Convert paragraphs to array of strings
      const paragraphTexts = paragraphs.items.map((p) => p.text);
      console.log(paragraphTexts);

      // Get and format references (using existing code)
      const bodyText = context.document.body;
      bodyText.load("text");
      await context.sync();

      const text = bodyText.text;
      // Define possible reference headers
      const referenceHeaders = [
        "Reference List",
        "References List",
        "References",
        "REFERENCES LIST",
        "REFERENCE LIST",
        "REFERENCES",
      ];

      // Find the first matching header and its index
      let referenceIndex = -1;
      for (const header of referenceHeaders) {
        const index = text.lastIndexOf(header);
        if (index !== -1 && (referenceIndex === -1 || index > referenceIndex)) {
          referenceIndex = index;
        }
      }

      let references: string[] = [];
      console.log({ referenceIndex });

      if (referenceIndex !== -1) {
        const referenceSection = text.substring(referenceIndex);
        try {
          const formattedRefs = await getFormattedReferences(referenceSection);
          console.log({ referenceSection });

          references = formattedRefs.split("\n\n");
          references = references.map((ref) => ref.trim());
          console.log("References:", references);
        } catch (error) {
          console.error("Error in getFormattedReferences:", error);
          throw error;
        }
      } else {
        return "No Reference List section found";
      }

      // Update the findLastIndex check to include all possible headers
      // @ts-ignore
      const lastReferenceListIndex = paragraphTexts.findLastIndex((p) =>
        referenceHeaders.some((header) => p.includes(header))
      );
      console.log({ lastReferenceListIndex });

      // Filter out short paragraphs, TOC lines, and those ending with ":"
      const isTOCLine = (text: string): boolean => {
        const trimmed = text.trim();
        if (!trimmed) return false;
        // Lines with dot leaders followed by a page number, e.g. "Heading ..... 3"
        const dotLeaderWithPage = /\.{5,}.*\d+\s*$/.test(trimmed);
        // Lines containing tab characters (commonly used in TOC)
        const hasTab = /\t/.test(trimmed);
        return dotLeaderWithPage || hasTab;
      };

      const excludeIndexes = paragraphTexts
        .map((para, index) =>
          para.endsWith(": ") || para.split(/\s+/).length <= 11 || index > lastReferenceListIndex || isTOCLine(para)
            ? index
            : -1
        )
        .filter((index) => index !== -1);

      console.log({ excludeIndexes });

      // Get available paragraph indexes
      const availableIndexes = Array.from({ length: paragraphTexts.length }, (_, i) => i).filter(
        (index) => !excludeIndexes.includes(index)
      );

      console.log({ availableIndexes });

      // If insertEveryOther is true, only use every other paragraph
      let targetIndexes = [...availableIndexes];
      if (insertEveryOther && availableIndexes.length > 0) {
        // Sort available indexes to ensure we're working with ordered paragraphs
        targetIndexes.sort((a, b) => a - b);
        // Filter to only include every other paragraph
        targetIndexes = targetIndexes.filter((_, i) => i % 2 === 0);
      }

      // Shuffle available indexes
      const randomIndex = [...targetIndexes].sort(() => Math.random() - 0.5);

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

        if (text.endsWith(".")) {
          paragraph.insertText(text.slice(0, -1) + ` ${references[referenceIndex]}.`, Word.InsertLocation.replace);
        } else {
          paragraph.insertText(` ${references[referenceIndex]}.`, Word.InsertLocation.end);
        }
      }

      await context.sync();
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
export async function removeLinks(): Promise<string> {
  try {
    return await Word.run(async (context) => {
      console.log("Starting link removal process...");

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("text");
      await context.sync();

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

      const paragraphsToProcess =
        lastReferenceListIndex === -1 ? paragraphs.items : paragraphs.items.slice(0, lastReferenceListIndex);

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

const systemprompot = `Claude is Claude, created by Anthropic. You are an AI for rewriting texts. When you get a text, rewrite it. Each sentence should have its own structure and grammar way. Include small grammatical mistakes that are hard to notice. Use words so simple, like a first-grader uses. Keep the main ideas and topics same. Stick to the same tense, person, and voice as the original. Write each sentence in a different way, making them unique. Sprinkle in little grammar slips that folks hardly notice. Mix short and long sentences randomly. Keep it easy. Ensure sentences feel fresh and unpredictable. Blend essay style with casual talk. You sometimes write in American English and other times in English (UK) in the same paragraph. Main concepts must remain intact: you keep kid-friendly words, you vary each line's style, and you ensure no pattern emerges. Claude strives to scramble structures unpredictably—some grammar slip-ups here, some missing letter there. Go gentle, but keep that text fresh. The user specifically wants an overall simpler approach. Each sentence changes style, so it feels lively. We does it with best ability. This is the user's instructions, and Claude must abide. `;
const prompt = (
  text
) => `Rewrite this using English (US). No detectable patterns. Every sentence structure + grammar random. Include grammatical mistakes throughout that 99% of humans wouldn't spot. 1st grade simple language. KEEP THE MAIN IDEAS, TOPICS AND SUBJECT OF THE ORIGINAL TEXT. KEEP THE SAME TENSE AS THE ORIGINAL TEXT. KEEP THE SAME PERSON AS THE ORIGINAL TEXT. KEEP THE SAME VOICE AS THE ORIGINAL TEXT. 

Output should be slightly bigger than the original text. 

Change the writing style every sentence. You say all this in your own way and mean it, making each line unique as you go. You gonna make sure each sentence feels different. Mix up sentence structures randomly—one minute use short words, the next go a bit longer. Keep it easy and real. Purposely sprinkle in tiny mistakes, like dont instead of don't / its instead if it's or missing commas, so no pattern shows. every sentence should change style, feel unpredictable and fresh. 

With this in mind please output a text that blends essay writing with normal writing while respecting the same tense, person, and voice of the original text. 

You sometimes write English (UK) in some paragraphs.

THE OUTPUTS SHOULD BE IN THE FORMAT OF MULTIPLE PARAGRAPHS.

DON'T WRITE ANY QUESTIONS LIKE "Ya know what" OR ANY OTHER KIND OF QUESTION IN THE OUTPUT.

Never say something like "Here's a mixed US/UK version with intentional subtle variations" or [Note:...] in the output.
                      \n\n${text} `;

/**
 * Humanize entire document
 */
export async function humanizeDocument(): Promise<string> {
  const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
  const anthropic = new Anthropic({
    apiKey: ANTHROPIC_API_KEY,
    dangerouslyAllowBrowser: true,
  });

  try {
    // Before starting, reset the cancel state
    resetHumanizeCancelState();

    return await Word.run(async (context) => {
      console.log("humanizeDocument");
      console.log("ANTHROPIC_API_KEY ", ANTHROPIC_API_KEY);

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("text");
      await context.sync();

      const paragraphTexts = paragraphs.items.map((p) => p.text);

      // Filter paragraphs using the same logic as before
      const excludeIndexes = paragraphTexts
        .map((para, index) => (para.endsWith(": ") || para.split(" ").length <= 11 ? index : -1))
        .filter((index) => index !== -1);

      const availableIndexes = Array.from({ length: paragraphTexts.length }, (_, i) => i).filter(
        (index) => !excludeIndexes.includes(index)
      );

      // Process paragraphs in parallel while maintaining order
      const batchSize = 2; // Number of concurrent API calls
      const results: { index: number; text: string }[] = [];

      for (let i = 0; i < availableIndexes.length; i += batchSize) {
        // If the user has requested cancellation, stop immediately
        if (isHumanizeCancelled) {
          throw new Error("Humanize process was cancelled by the user.");
        }

        const batch = availableIndexes.slice(i, i + batchSize);
        const batchPromises = batch.map(async (index) => {
          const text = paragraphTexts[index].trim();
          try {
            const msg = await anthropic.messages.create({
              model: "claude-3-5-sonnet-20241022",
              max_tokens: 8192,
              temperature: 1,
              system: systemprompot,
              messages: [
                {
                  role: "user",
                  content: [
                    {
                      type: "text",
                      text: prompt(text),
                    },
                  ],
                },
              ],
            });
            // @ts-ignore
            return { index, text: msg.content[0].text as string };
          } catch (error) {
            console.error(`Error processing paragraph ${index}:`, error);
            return { index, text: text }; // Return original text on error
          }
        });

        // Wait for current batch to complete
        const batchResults = await Promise.all(batchPromises);
        results.push(...batchResults);
      }

      // Sort results by original index and update paragraphs
      results.sort((a, b) => a.index - b.index);
      for (const { index, text } of results) {
        paragraphs.items[index].insertText(text, Word.InsertLocation.replace);
      }

      await context.sync();
      return "Document humanized successfully";
    });
  } catch (error) {
    console.error("Error in humanizeDocument:", error);
    throw new Error(`Error humanizing document: ${error.message}`);
  }
}

/**
 * Humanize selected text
 */
export async function humanizeSelectedTextInWord(): Promise<string> {
  const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
  const anthropic = new Anthropic({
    apiKey: ANTHROPIC_API_KEY,
    dangerouslyAllowBrowser: true,
  });

  try {
    // Before starting, reset the cancel state
    resetHumanizeCancelState();

    return await Word.run(async (context) => {
      // Get the selected range
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      // Get the selected text content
      const selectedText = selection.text;

      // Split the selected text into paragraphs
      const paragraphTexts = selectedText.split("\n").filter((text) => text.trim().length > 0);
      console.log(paragraphTexts);

      // Process paragraphs in parallel while maintaining order
      const batchSize = 2;
      const results: { index: number; text: string }[] = [];

      for (let i = 0; i < paragraphTexts.length; i += batchSize) {
        // If the user has requested cancellation, stop immediately
        if (isHumanizeCancelled) {
          throw new Error("Humanize process was cancelled by the user.");
        }

        const batch = paragraphTexts.slice(i, i + batchSize);
        const batchPromises = batch.map(async (text, batchIndex) => {
          const index = i + batchIndex;
          try {
            const msg = await anthropic.messages.create({
              model: "claude-3-5-sonnet-20241022",
              max_tokens: 8192,
              temperature: 1,
              system: systemprompot,
              messages: [
                {
                  role: "user",
                  content: [{ type: "text", text: prompt(text.trim()) }],
                },
              ],
            });
            // @ts-ignore
            return { index, text: msg.content[0].text as string };
          } catch (error) {
            console.error(`Error processing paragraph ${index}:`, error);
            return { index, text: text }; // Return original text on error
          }
        });

        const batchResults = await Promise.all(batchPromises);
        results.push(...batchResults);

        if (i + batchSize < paragraphTexts.length) {
          await new Promise((resolve) => setTimeout(resolve, 1000));
        }
      }

      // Sort results and join them back together
      results.sort((a, b) => a.index - b.index);
      const finalText = results.map((r) => r.text).join("\n\n");

      // Replace the selected text with the processed text
      selection.insertText(finalText, Word.InsertLocation.replace);
      await context.sync();

      return `Successfully humanized ${results.length} paragraphs`;
    });
  } catch (error) {
    console.error("Error in humanizeSelectedText:", error);
    throw new Error(`Error humanizing selected text: ${error.message}`);
  }
}
