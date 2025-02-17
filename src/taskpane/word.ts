/* global Word console */
import { getFormattedReferences } from "./gemini";

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

export async function analyzeDocument(): Promise<string> {
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

      // Filter out short paragraphs and those ending with ":"
      const excludeIndexes = paragraphTexts
        .map((para, index) =>
          para.endsWith(": ") || para.split(" ").length <= 11 || index > lastReferenceListIndex ? index : -1
        )
        .filter((index) => index !== -1);

      console.log({ excludeIndexes });

      // Get available paragraph indexes
      const availableIndexes = Array.from({ length: paragraphTexts.length }, (_, i) => i).filter(
        (index) => !excludeIndexes.includes(index)
      );

      console.log({ availableIndexes });

      // Shuffle available indexes
      const randomIndex = [...availableIndexes].sort(() => Math.random() - 0.5);

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
        /\((?:[^,()]+(,\s[^,()]+)*(?:,\sand\s[^,()]+)?)[,\s]\s?\d{4}\)/g, // (Author, Author, and Author, 2024) or (Author, Author, and Author 2024)

        // Standard patterns - non-greedy match
        /\((?:[^,()]+)[,\s]\s?\d{4}\)/g, // (Author, 2024) or (Author 2024)
        /\((?:[^()]+\sand\s[^,()]+)[,\s]\s?\d{4}\)/g, // (Author and Author, 2024) or (Author and Author 2024)
        /\((?:[^()]+)\set\sal\.?[,\s]\s?\d{4}\)/g, // (Author et al., 2024) or (Author et al 2024)

        // Additional patterns for edge cases - non-greedy match
        /\((?:[^,()]+(,\s[^,()]+)*)[,\s]\s?\d{4}\)/g, // (Author, Author, Author, 2024) or (Author, Author, Author 2024)
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

interface SubmitResponse {
  status: string;
  id: string;
}

interface DocumentResponse {
  id: string;
  output: string;
  input: string;
  readability: string;
  createdDate: string;
  purpose: string;
}

async function getDocumentWithRetries(
  apiKey: string,
  documentId: string,
  maxRetries = 3,
  delayMs = 5000
): Promise<DocumentResponse> {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    // Wait before each attempt
    await new Promise((resolve) => setTimeout(resolve, delayMs));

    const documentResponse = await fetch("https://humanize.undetectable.ai/document", {
      method: "POST",
      headers: {
        apikey: apiKey,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ id: documentId }),
    });

    const documentResult: DocumentResponse = await documentResponse.json();

    if (documentResult.output) {
      return documentResult;
    }

    console.log(`Attempt ${attempt}: Output not ready yet, retrying...`);

    // Increase delay for next attempt
    delayMs *= 1.5;
  }

  throw new Error("Failed to get document output after maximum retries");
}

export async function humanizeDocument(): Promise<string> {
  const API_KEY = process.env.HUMANIZE_API_KEY;

  try {
    return await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("text");
      await context.sync();

      const paragraphTexts = paragraphs.items.map((p) => p.text);

      // Filter paragraphs using the same logic as analyzeDocument
      const excludeIndexes = paragraphTexts
        .map((para, index) => (para.endsWith(": ") || para.split(" ").length <= 11 ? index : -1))
        .filter((index) => index !== -1);

      const availableIndexes = Array.from({ length: paragraphTexts.length }, (_, i) => i).filter(
        (index) => !excludeIndexes.includes(index)
      );

      // Process each valid paragraph
      for (const index of availableIndexes) {
        const paragraph = paragraphs.items[index];
        const text = paragraph.text.trim();

        if (text.length >= 50) {
          // API requirement
          try {
            // Step 1: Submit the content
            const submitResponse = await fetch("https://humanize.undetectable.ai/submit", {
              method: "POST",
              headers: {
                apikey: API_KEY,
                "Content-Type": "application/json",
              },
              body: JSON.stringify({
                content: text,
                readability: "High School",
                purpose: "General Writing",
                strength: "More Human",
                model: "v11",
              }),
            });

            const submitResult: SubmitResponse = await submitResponse.json();

            // Step 2: Get document with retries
            try {
              const documentResult = await getDocumentWithRetries(API_KEY, submitResult.id);
              paragraph.insertText(documentResult.output, Word.InsertLocation.replace);
              await context.sync(); // Sync after each paragraph update
            } catch (retryError) {
              console.error(`Failed to get humanized content for paragraph ${index}:`, retryError);
            }
          } catch (error) {
            console.error(`Error processing paragraph ${index}:`, error);
          }
        }
      }

      await context.sync();
      return "Document humanized successfully";
    });
  } catch (error) {
    console.error("Error in humanizeDocument:", error);
    throw new Error(`Error humanizing document: ${error.message}`);
  }
}
