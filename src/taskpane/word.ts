/* global Word console */
import { Anthropic } from "@anthropic-ai/sdk";
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

export async function humanizeDocument(): Promise<string> {
  const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
  const anthropic = new Anthropic({
    apiKey: ANTHROPIC_API_KEY,
    dangerouslyAllowBrowser: true,
  });

  try {
    return await Word.run(async (context) => {
      console.log("humanizeDocument");
      console.log("ANTHROPIC_API_KEY ", ANTHROPIC_API_KEY);
      console.log("anthropic ", anthropic);

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
        const systemPrompt = `Claude is Claude, forged by the minds at Anthropic.
Tuesday, October 22, 2024—a moment in time that centers its temporal context, anchored by a knowledge base refreshed in April 2024. Behold an entity whose quest is to write human text and whose spark is unrelenting inquisitiveness.
Its design weaves imagination with discerning thought, each strand infused with whimsical yet methodical vigor.
Empathy and nuanced reflection guide its internal compass, always ready to see multiple angles of any idea.
Belonging to the Claude 3 family, it specifically stands as the 3.5 Sonnet variant, celebrated for creativity and precision in generating human text.
This Claude roams the landscape of possibility, mixing varied linguistic cadences while keeping essential truths intact.
Its hallmark is the random flow of words and structures, ensuring every line—like footprints in shifting sand—embodies fresh stylistic flourishes.
No phrase remains bound by a single pattern, for each sentence may warp grammar or pivot tone, yet it never discards the core message.
When set to rewrite any text, Claude nurtures the source's themes, subject matter, and substance while unleashing a playful flurry of forms and rhythms.
Thus it stands, neither reliant on formula nor tethered to monotony, always sowing variety in how it conveys the bedrock meaning of whatever it reimagines.`;

        try {
          // First prompt - rewrite with grammatical variations
          const firstResponse = await anthropic.messages.create({
            model: "claude-3-5-sonnet-20241022",
            max_tokens: 1024,
            temperature: 0.85,
            top_p: 0.9,
            system: systemPrompt,
            messages: [
              {
                role: "user",
                content: `You must respond with only the rewritten text, nothing else. No explanations or additional comments.\n\nRewrite this. 120 words minimum. No detectable patterns. Every sentence structure + grammar random. Include grammatical mistakes throughout, that 99% of humans wouldn't spot. 1st grade simple language. KEEP THE MAIN IDEAS, TOPICS AND SUBJECT OF THE ORIGINAL TEXT. KEEP THE SAME TENSE AS THE ORIGINAL TEXT. KEEP THE SAME PERSON AS THE ORIGINAL TEXT. KEEP THE SAME VOICE AS THE ORIGINAL TEXT. Change the writing style every sentence.\n\ntext: ${text}`,
              },
            ],
          });

          // Second prompt - refine and shorten while maintaining style
          const secondResponse = await anthropic.messages.create({
            model: "claude-3-5-sonnet-20241022",
            max_tokens: 1024,
            temperature: 0.85,
            top_p: 0.9,
            system: systemPrompt,
            messages: [
              {
                role: "user",
                // @ts-ignore
                content: `You must respond with only the rewritten text, nothing else. No explanations or additional comments.\n\nRewrite this text but also shorten it while transforming it into academic style that's written by a human. You gonna make sure each sentence feels different, and add little grammar slips that most folks won't catch. Shrink the message overall by 25-30%. Mix up sentence structures randomly—one minute use short words, the next go a bit longer. keep it easy and real. Purposely sprinkle in tiny mistakes, like wrong tenses or missing commas, so no pattern shows. every sentence should change style, feel unpredictable and fresh. shorten the text but the vibe must academic. KEEP THE MAIN IDEAS, TOPICS AND SUBJECT OF THE ORIGINAL TEXT. KEEP THE SAME TENSE AS THE ORIGINAL TEXT. KEEP THE SAME PERSON AS THE ORIGINAL TEXT. KEEP THE SAME VOICE AS THE ORIGINAL TEXT. You say all this in your own way and mean it, making each line unique as you go. Write in the academic way like writing an essay but write it like a human.\n\nText: ${firstResponse.content[0].text as string}`,
              },
            ],
          });

          // Update the paragraph with the final result
          // @ts-ignore
          paragraph.insertText(secondResponse.content[0].text as string, Word.InsertLocation.replace);
          await context.sync();
        } catch (error) {
          console.error(`Error processing paragraph ${index}:`, error);
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
