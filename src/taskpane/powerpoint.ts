/* eslint-disable office-addins/no-context-sync-in-loop */
/* global PowerPoint console, Office, fetch */
import { getFormattedReferences } from "./gemini";

const TRAILING_PUNCTUATION = "[-:;,.!?–—]*";
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

type ShapeMeta = {
  slideIndex: number;
  shapeId: string;
  text: string;
  wordCount: number;
  isTitle: boolean;
  index: number; // Global index for linear processing
};

function isHeadingOrTitle(meta: ShapeMeta): boolean {
  // Heuristics for PowerPoint where we don't have "styles" easily accessible
  if (meta.isTitle) return true;

  const text = meta.text;
  if (!text) return false;

  const hasTerminalPunctuation = /[.!?]$/.test(text);
  const isShortish = meta.wordCount > 0 && meta.wordCount <= 10;

  const words = text.split(/\s+/).filter(Boolean);
  const capitalised = words.filter((w) => /^[A-Z]/.test(w)).length;
  const isTitleCase = words.length > 1 && capitalised / words.length > 0.6;

  if (!hasTerminalPunctuation && isShortish) {
    return true;
  }
  if (isTitleCase && isShortish) {
    return true;
  }

  return false;
}

function findReferenceStartIndex(metas: ShapeMeta[]): number {
  for (let i = metas.length - 1; i >= 0; i--) {
    if (REFERENCE_HEADERS.some((regex) => regex.test(metas[i].text))) {
      return i;
    }
  }
  return -1;
}

function findConclusionRange(
  metas: ShapeMeta[],
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
  let reason = "End of document (no next section found)";

  if (conclusionHeadingIndex !== -1) {
    for (let i = conclusionHeadingIndex + 1; i < metas.length; i++) {
      if (REFERENCE_HEADERS.some((regex) => regex.test(metas[i].text))) {
        conclusionEndIndex = i;
        reason = `Found Reference Header at index ${i}: "${metas[i].text}"`;
        break;
      }

      if (isHeadingOrTitle(metas[i]) && !CONCLUSION_HEADERS.some((regex) => regex.test(metas[i].text))) {
        conclusionEndIndex = i;
        reason = `Found next Heading at index ${i}: "${metas[i].text}"`;
        break;
      }
    }

    if (conclusionEndIndex === -1 && referenceStartIndex !== -1) {
      conclusionEndIndex = referenceStartIndex;
      reason = `Reached Reference Start Index at ${referenceStartIndex}`;
    }
  }

  console.log(`Conclusion Detection: Start Index: ${conclusionHeadingIndex}`);
  if (conclusionHeadingIndex !== -1) {
    console.log(`Conclusion Detection: End Index: ${conclusionEndIndex}`);
    console.log(`Conclusion Detection: End Reason: ${reason}`);
  }

  return { conclusionHeadingIndex, conclusionEndIndex };
}

function splitIntoSentences(text: string): string[] {
  // Split text into sentences, keeping delimiters.
  // This regex matches a sequence of characters ending with . ! ? or end of string.
  // It tries to handle quotes/brackets after punctuation.
  return text.match(/[^.!?]+(?:[.!?]+["')\]]*)?|[^.!?]+$/g) || [text];
}

function appendCitationAtSentenceEnd(sentence: string, citation: string): string {
  const trimmed = sentence.trimEnd();
  const match = trimmed.match(/([.?!]["')\]]*)$/);
  if (!match) {
    const sep = /\s$/.test(trimmed) ? "" : " ";
    return trimmed + sep + citation;
  }
  const punctuation = match[1];
  const core = trimmed.slice(0, trimmed.length - punctuation.length);
  const sep = /\s$/.test(core) ? "" : " ";
  return `${core}${sep}${citation}${punctuation}`;
}

function shuffleInPlace<T>(array: T[]): T[] {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

export async function insertText(text: string) {
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();

      const slide = slides.items[0];
      const textBox = slide.shapes.addTextBox(text);
      textBox.fill.setSolidColor("white");
      textBox.lineFormat.color = "black";
      textBox.lineFormat.weight = 1;
      textBox.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.solid;
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function analyzeDocument(insertEveryOther: boolean = false): Promise<string> {
  try {
    return await PowerPoint.run(async (context) => {
      console.log("analyzeDocument => start", { insertEveryOther });
      const slides = context.presentation.slides;
      // Load all text from all shapes in all slides
      slides.load("items/shapes/items/textFrame/textRange/text, items/shapes/items/id, items/shapes/items/type");
      await context.sync();

      const metas: ShapeMeta[] = [];
      let globalIndex = 0;

      // Build metadata
      for (let i = 0; i < slides.items.length; i++) {
        const slide = slides.items[i];
        for (let j = 0; j < slide.shapes.items.length; j++) {
          const shape = slide.shapes.items[j];
          let text = "";
          try {
            // @ts-ignore
            text = shape.textFrame ? shape.textFrame.textRange.text : "";
          } catch (e) {
            continue;
          }

          text = text.replace(/[\u200B-\u200D\uFEFF]/g, "").trim();
          if (text) {
            const words = text.split(/\s+/).filter(Boolean);
            // Heuristic: if it's the first shape on slide and short, maybe title?
            // Or check shape type if possible.
            const isTitle = j === 0 && words.length < 10;

            metas.push({
              slideIndex: i,
              shapeId: shape.id,
              text,
              wordCount: words.length,
              isTitle,
              index: globalIndex++,
            });
          }
        }
      }

      console.log("analyzeDocument => shape count", metas.length);
      if (metas.length === 0) {
        return "No text content found in the presentation";
      }

      const referenceStartIndex = findReferenceStartIndex(metas);
      console.log({ referenceStartIndex });
      if (referenceStartIndex === -1) {
        return "No Reference List section found";
      }

      const referenceSection = metas
        .slice(referenceStartIndex)
        .map((m) => m.text)
        .join("\n");

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
        return "No valid references found in the Reference List section";
      }

      const { conclusionHeadingIndex, conclusionEndIndex } = findConclusionRange(metas, referenceStartIndex);

      // Filter eligible shapes
      const eligibleMetas = metas.filter((meta) => {
        if (referenceStartIndex !== -1 && meta.index >= referenceStartIndex) return false;
        if (
          conclusionHeadingIndex !== -1 &&
          conclusionEndIndex !== -1 &&
          meta.index > conclusionHeadingIndex &&
          meta.index < conclusionEndIndex
        )
          return false;

        if (isHeadingOrTitle(meta)) return false;
        if (meta.wordCount < 11) return false;
        if (meta.text.trim().endsWith(":")) return false;

        return true;
      });

      if (eligibleMetas.length === 0) {
        return "No eligible text shapes found for inserting references";
      }

      let targetMetas = [...eligibleMetas];
      if (insertEveryOther && targetMetas.length > 0) {
        targetMetas = targetMetas.filter((_, i) => i % 2 === 0);
      }

      // Sentence-level injection logic
      // Since we can't easily get sentence ranges in PPT, we'll work with text splitting
      type SentenceSlot = {
        meta: ShapeMeta;
        sentenceIndex: number;
        sentenceText: string;
      };

      const allSlots: SentenceSlot[] = [];
      const sentencesByMeta: Map<ShapeMeta, string[]> = new Map();

      for (const meta of targetMetas) {
        const sentences = splitIntoSentences(meta.text);
        sentencesByMeta.set(meta, sentences);

        // Identify slots
        for (let i = 0; i < sentences.length; i++) {
          const sText = sentences[i].trim();
          if (!sText) continue;

          // Skip first sentence if multi-sentence
          if (i === 0 && sentences.length > 1) continue;

          const wordCount = sText.split(/\s+/).filter(Boolean).length;
          if (wordCount < 8) continue;
          if (/\(\s*[^)]*?\d{4}[a-z]?\s*\)/.test(sText)) continue;

          const lower = sText.toLowerCase();
          if (
            lower.startsWith("in conclusion") ||
            lower.startsWith("to conclude") ||
            lower.startsWith("overall,") ||
            lower.startsWith("to sum up")
          )
            continue;

          allSlots.push({ meta, sentenceIndex: i, sentenceText: sText });
        }
      }

      // Limit per paragraph (shape) logic is tricky here because we flattened slots.
      // But we can shuffle all slots and then enforce limit when applying?
      // Or group slots by meta first.

      // Let's group slots by meta to enforce 0-3 limit
      const slotsByMeta = new Map<ShapeMeta, SentenceSlot[]>();
      for (const slot of allSlots) {
        if (!slotsByMeta.has(slot.meta)) slotsByMeta.set(slot.meta, []);
        slotsByMeta.get(slot.meta).push(slot);
      }

      const finalSlots: SentenceSlot[] = [];
      Array.from(slotsByMeta.values()).forEach((slots) => {
        if (slots.length === 0) return;
        shuffleInPlace(slots);
        const maxForShape = Math.min(3, slots.length);
        const targetCount = 1 + Math.floor(Math.random() * maxForShape);
        const selected = slots.slice(0, targetCount);
        finalSlots.push(...selected);
      });

      shuffleInPlace(finalSlots);
      console.log("analyzeDocument => sentence slots count", finalSlots.length);

      const usedReferences = new Set<number>();

      // We need to apply changes. Since multiple slots might be in same shape, we need to be careful.
      // We should group final slots by meta again to apply all changes to a shape at once.
      const changesByMeta = new Map<ShapeMeta, Map<number, string>>(); // meta -> (sentenceIndex -> newText)

      for (const slot of finalSlots) {
        let referenceIndex: number;
        if (usedReferences.size < references.length) {
          const unusedReferences = Array.from(Array(references.length).keys()).filter((i) => !usedReferences.has(i));
          referenceIndex = unusedReferences[Math.floor(Math.random() * unusedReferences.length)];
          usedReferences.add(referenceIndex);
        } else {
          referenceIndex = Math.floor(Math.random() * references.length);
        }

        const citation = references[referenceIndex];
        const newSentence = appendCitationAtSentenceEnd(slot.sentenceText, citation);

        if (!changesByMeta.has(slot.meta)) changesByMeta.set(slot.meta, new Map());
        changesByMeta.get(slot.meta).set(slot.sentenceIndex, newSentence);
      }

      // Apply changes
      Array.from(changesByMeta.entries()).forEach(([meta, changes]) => {
        const sentences = sentencesByMeta.get(meta);
        if (!sentences) return;

        // Reconstruct text
        const newSentences = sentences.map((s, i) => (changes.has(i) ? changes.get(i) : s));
        const reconstructedText = newSentences.join("");

        // Find the shape and update it
        const slide = slides.items[meta.slideIndex];
        // We need to find shape by ID.
        // We can iterate shapes again or use `getItem(id)` if available.
        // `slide.shapes.getItem(id)` exists.
        const shape = slide.shapes.getItem(meta.shapeId);
        // We need to load textFrame again to write? No, just write.
        // But we need to make sure we are writing to textRange.
        shape.textFrame.textRange.text = reconstructedText;
        // Unbold
        shape.textFrame.textRange.font.bold = false;
      });

      await context.sync();
      return "References added successfully";
    });
  } catch (error) {
    console.error("Error in analyzeDocument:", error);
    throw new Error(`Error modifying document: ${error.message}`);
  }
}

export async function removeReferences(): Promise<string> {
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/shapes/items/textFrame/textRange/text, items/shapes/items/textFrame/textRange/font");
      await context.sync();

      const citationPatterns = [
        /\((?:[^,()]+(,\s[^,()]+)*(?:,\sand\s[^,()]+)?)[,\s]\s?\d{4}[a-z]?\)/g,
        /\((?:[^,()]+)[,\s]\s?\d{4}[a-z]?\)/g,
        /\((?:[^()]+\sand\s[^,()]+)[,\s]\s?\d{4}[a-z]?\)/g,
        /\((?:[^()]+)\set\sal\.?[,\s]\s?\d{4}[a-z]?\)/g,
        /\((?:[^,()]+(,\s[^,()]+)*)[,\s]\s?\d{4}[a-z]?\)/g,
      ];

      let totalRemoved = 0;

      for (let slide of slides.items) {
        for (let shape of slide.shapes.items) {
          // @ts-ignore
          if (shape.textFrame && shape.textFrame.textRange) {
            let text = shape.textFrame.textRange.text;
            let hadMatch = false;

            for (const pattern of citationPatterns) {
              const matches = text.match(pattern) || [];
              if (matches.length > 0) {
                text = text.replace(pattern, "").replace(/\s+\./g, ".");
                hadMatch = true;
                totalRemoved += matches.length;
              }
            }

            if (hadMatch) {
              shape.textFrame.textRange.text = text;
              shape.textFrame.textRange.font.bold = false;
            }
          }
        }
      }

      await context.sync();
      return `Removed ${totalRemoved} citations and cleaned up formatting`;
    });
  } catch (error) {
    console.error("Error in removeReferences:", error);
    throw new Error(`Error removing references: ${error.message}`);
  }
}

export async function removeLinks(deleteAll: boolean = false): Promise<string> {
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/shapes/items/textFrame/textRange/text");
      await context.sync();

      // We need to find reference section to skip it if !deleteAll
      // This requires building the full text map again or just iterating.
      // For simplicity, let's iterate and check if we are in reference section.
      // But shapes are not linear like paragraphs.
      // We'll assume we process all shapes unless we detect we are in reference slide/shape?
      // Implementing "skip reference section" in PPT is harder without linear structure.
      // I'll implement simple version: remove links everywhere for now, or try to detect reference header in the shape itself.

      const urlRegex = /\b((https?:\/\/)?[\w.-]+(?:\.[\w.-]+)+)\b/g;
      let linksRemovedCount = 0;

      for (let slide of slides.items) {
        for (let shape of slide.shapes.items) {
          // @ts-ignore
          if (shape.textFrame && shape.textFrame.textRange) {
            let text = shape.textFrame.textRange.text;
            // Skip if text looks like reference header?
            if (!deleteAll && REFERENCE_HEADERS.some((r) => r.test(text))) {
              continue;
            }

            const matches = text.match(urlRegex);
            if (matches) {
              linksRemovedCount += matches.length;
              const newText = text.replace(urlRegex, "").replace(/\s+([.,;])/g, "$1");
              shape.textFrame.textRange.text = newText;
              shape.textFrame.textRange.font.bold = false;
            }
          }
        }
      }
      await context.sync();
      return `Removed ${linksRemovedCount} URL-like text snippets.`;
    });
  } catch (error) {
    console.error("Error in removeLinks:", error);
    throw new Error(`Error removing links: ${error.message}`);
  }
}

export async function removeWeirdNumbers(): Promise<string> {
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/shapes/items/textFrame/textRange/text");
      await context.sync();

      const weirdNumberPattern = /[【[]\d+.*?[†+t].*?[】\]]\S*/g;
      let totalRemoved = 0;

      for (let slide of slides.items) {
        for (let shape of slide.shapes.items) {
          // @ts-ignore
          if (shape.textFrame && shape.textFrame.textRange) {
            let text = shape.textFrame.textRange.text;
            const matches = text.match(weirdNumberPattern);
            if (matches && matches.length > 0) {
              totalRemoved += matches.length;
              const newText = text.replace(weirdNumberPattern, "").replace(/\s{2,}/g, " ");
              shape.textFrame.textRange.text = newText;
              shape.textFrame.textRange.font.bold = false;
            }
          }
        }
      }
      await context.sync();
      return `Removed ${totalRemoved} weird number instances.`;
    });
  } catch (error) {
    console.error("Error in removeWeirdNumbers:", error);
    throw new Error(`Error removing weird numbers: ${error.message}`);
  }
}

export async function paraphraseSelectedText(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, async (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
        return;
      }
      const selectedText = result.value as string;
      if (!selectedText || !selectedText.trim()) {
        reject(new Error("No text selected"));
        return;
      }

      try {
        const response = await fetch("https://analizeai.com/paraphrase", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ text: selectedText.trim() }),
        });

        if (!response.ok) {
          throw new Error(`API request failed with status ${response.status}`);
        }

        const data = await response.json();
        const paraphrasedText = data.secondMode;

        if (!paraphrasedText) {
          throw new Error("Invalid response from paraphrase API");
        }

        Office.context.document.setSelectedDataAsync(paraphrasedText, (setResult) => {
          if (setResult.status === Office.AsyncResultStatus.Failed) {
            reject(new Error(setResult.error.message));
          } else {
            resolve("Text paraphrased successfully");
          }
        });
      } catch (error) {
        reject(error);
      }
    });
  });
}
