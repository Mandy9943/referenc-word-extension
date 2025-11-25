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

function matchesReferenceHeader(rawText: string): boolean {
  if (!rawText) return false;
  const trimmed = rawText.trim();
  if (REFERENCE_HEADERS.some((regex) => regex.test(trimmed))) {
    return true;
  }

  const firstLine = trimmed.split(/\r?\n/)[0]?.trim();
  if (firstLine && firstLine !== trimmed) {
    return REFERENCE_HEADERS.some((regex) => regex.test(firstLine));
  }
  return false;
}

const CONCLUSION_HEADERS = [
  new RegExp(`^\\s*${NUMBERING_PREFIX}conclusions?(?:\\s+section)?\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}concluding\\s+remarks\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}final\\s+thoughts\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}summary(?:\\s+and\\s+future\\s+work)?\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}closing\\s+remarks\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
  new RegExp(`^\\s*${NUMBERING_PREFIX}conclusions?\\s+and\\s+recommendations\\s*${TRAILING_PUNCTUATION}\\s*$`, "i"),
];

type ShapeMeta = {
  slideId: string;
  slideIndex: number;
  shapeId: string;
  shapeIndex: number;
  shapeType: string;
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
    if (matchesReferenceHeader(metas[i].text)) {
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
      if (matchesReferenceHeader(metas[i].text)) {
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

const TEXT_CAPABLE_SHAPE_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder", "SmartArt", "Group"]);
const PARAPHRASE_DELIMITER = "qbpdelim123";

type TextShapeHandle = {
  slide: PowerPoint.Slide;
  shape: PowerPoint.Shape;
  slideIndex: number;
  shapeIndex: number;
};

function describeShape(handle: TextShapeHandle): string {
  return `slide ${handle.slideIndex} shape ${handle.shapeIndex}`;
}

async function loadTextCapableShapes(context: PowerPoint.RequestContext): Promise<TextShapeHandle[]> {
  const slides = context.presentation.slides;
  slides.load("items/index, items/id, items/shapes/items/id, items/shapes/items/type");
  await context.sync();

  const handles: TextShapeHandle[] = [];

  for (let i = 0; i < slides.items.length; i++) {
    const slide = slides.items[i];
    for (let j = 0; j < slide.shapes.items.length; j++) {
      const shape = slide.shapes.items[j];
      if (!TEXT_CAPABLE_SHAPE_TYPES.has(shape.type)) continue;
      handles.push({ slide, shape, slideIndex: i, shapeIndex: j });
    }
  }

  return handles;
}

async function tryLoadShapeText(handle: TextShapeHandle, context: PowerPoint.RequestContext): Promise<string | null> {
  const location = describeShape(handle);

  try {
    handle.shape.textFrame.textRange.load("text, font");
  } catch (error) {
    console.warn(`tryLoadShapeText => unable to queue text load for ${location}`, error);
    return null;
  }

  try {
    await context.sync();
  } catch (error) {
    console.warn(`tryLoadShapeText => context sync failed for ${location}`, error);
    return null;
  }

  try {
    return handle.shape.textFrame.textRange.text as string;
  } catch (error) {
    console.warn(`tryLoadShapeText => unable to read text for ${location}`, error);
    return null;
  }
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
      const slides = context.presentation.slides;

      // Load basic properties first to identify shapes with text
      slides.load("items/id, items/shapes/items/id, items/shapes/items/type");
      await context.sync();

      for (let slide of slides.items) {
        for (let shape of slide.shapes.items) {
          // Only try to load text for shapes that can have text
          if (TEXT_CAPABLE_SHAPE_TYPES.has(shape.type)) {
            shape.textFrame.textRange.load("text");
          }
        }
      }
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
            text = shape.textFrame.textRange.text;
          } catch (e) {
            continue;
          }

          text = text.replace(/[\u200B-\u200D\uFEFF]/g, "").trim();
          if (text) {
            const words = text.split(/\s+/).filter(Boolean);
            const isTitle = j === 0 && words.length < 10;

            metas.push({
              slideId: slide.id,
              slideIndex: i,
              shapeId: shape.id,
              shapeIndex: j,
              shapeType: shape.type,
              text,
              wordCount: words.length,
              isTitle,
              index: globalIndex++,
            });
          }
        }
      }

      if (metas.length === 0) {
        return "No text content found in the presentation";
      }

      const referenceStartIndex = findReferenceStartIndex(metas);
      if (referenceStartIndex === -1) {
        return "No Reference List section found";
      }

      const headerMeta = metas[referenceStartIndex];
      const previousSameSlideMetas = metas.filter(
        (meta) => meta.slideId === headerMeta.slideId && meta.index < referenceStartIndex
      );
      const tailAfterHeader = metas.slice(referenceStartIndex + 1);
      const referenceTailMetas = [headerMeta, ...previousSameSlideMetas, ...tailAfterHeader];
      console.log(
        "analyzeDocument => reference metas preview:",
        referenceTailMetas.slice(0, 5).map((meta) => ({
          slide: meta.slideIndex,
          shape: meta.shapeIndex,
          words: meta.wordCount,
          preview: meta.text.substring(0, 120),
        }))
      );

      const referenceSection = referenceTailMetas
        .map((m) => m.text)
        .join("\n")
        .trim();

      console.log("analyzeDocument => reference section detected:", referenceSection);
      let references: string[] = [];
      try {
        console.log("analyzeDocument => sending payload to formatter:\n", referenceSection);
        const formattedRefs = await getFormattedReferences(referenceSection);
        console.log("analyzeDocument => AI response:\n", formattedRefs);
        references = formattedRefs
          .split(/\n\s*\n/)
          .map((ref) => ref.trim())
          .filter(Boolean);
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

      const usedReferences = new Set<number>();

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

      // Apply changes sequentially, per shape, to identify problematic slides precisely
      const slidesById = new Map(slides.items.map((slide) => [slide.id, slide]));
      const entries = Array.from(changesByMeta.entries());

      for (const [meta, changes] of entries) {
        const sentences = sentencesByMeta.get(meta);
        if (!sentences) continue;

        const slide = slidesById.get(meta.slideId);
        if (!slide) {
          console.warn(`analyzeDocument => slide not found for meta slideId=${meta.slideId}`);
          continue;
        }

        let shape: PowerPoint.Shape;
        try {
          shape = slide.shapes.getItem(meta.shapeId);
        } catch (error) {
          console.error(`analyzeDocument => unable to fetch shape ${meta.shapeId} on slide ${meta.slideId}`, error);
          continue;
        }

        const newSentences = sentences.map((s, i) => (changes.has(i) ? changes.get(i) : s));
        const reconstructedText = newSentences.join("");

        try {
          shape.textFrame.textRange.load("text");
          await context.sync();
        } catch (error) {
          console.error(`analyzeDocument => failed loading shape ${meta.shapeId}`, error);
          continue;
        }

        try {
          shape.textFrame.textRange.text = reconstructedText;

          shape.textFrame.textRange.font.bold = false;
          await context.sync();
        } catch (error) {
          console.error(`analyzeDocument => failed to apply update for shape ${meta.shapeId}`, error);
        }
      }

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
      console.log("removeReferences => start");
      const textShapes = await loadTextCapableShapes(context);
      console.log(`removeReferences => loaded ${textShapes.length} text-capable shapes`);
      if (textShapes.length === 0) {
        return "No text content available to clean";
      }

      const citationPatterns = [
        /\((?:[^,()]+(,\s[^,()]+)*(?:,\sand\s[^,()]+)?)[,\s]\s?\d{4}[a-z]?\)/g,
        /\((?:[^,()]+)[,\s]\s?\d{4}[a-z]?\)/g,
        /\((?:[^()]+\sand\s[^,()]+)[,\s]\s?\d{4}[a-z]?\)/g,
        /\((?:[^()]+)\set\sal\.?[,\s]\s?\d{4}[a-z]?\)/g,
        /\((?:[^,()]+(,\s[^,()]+)*)[,\s]\s?\d{4}[a-z]?\)/g,
      ];

      let totalRemoved = 0;
      let shapesUpdated = 0;

      for (const handle of textShapes) {
        const location = describeShape(handle);
        console.log(`removeReferences => inspecting ${location}`);
        const text = await tryLoadShapeText(handle, context);
        if (!text) {
          console.warn(`removeReferences => skipping ${location} (unable to load text)`);
          continue;
        }

        let updatedText = text;
        let hadMatch = false;

        for (const pattern of citationPatterns) {
          const matches = updatedText.match(pattern) || [];
          if (matches.length > 0) {
            updatedText = updatedText.replace(pattern, "");
            hadMatch = true;
            totalRemoved += matches.length;
          }
        }

        if (!hadMatch) continue;

        updatedText = updatedText
          .replace(/\s+\./g, ".")
          .replace(/\s{2,}/g, " ")
          .trim();
        console.log(`removeReferences => updating ${location} (removed citations)`);
        handle.shape.textFrame.textRange.text = updatedText;
        handle.shape.textFrame.textRange.font.bold = false;
        shapesUpdated++;
      }

      await context.sync();
      console.log(`removeReferences => finished sync (totalRemoved=${totalRemoved}, shapesUpdated=${shapesUpdated})`);
      if (totalRemoved === 0) {
        return "No citations detected";
      }

      return `Removed ${totalRemoved} citations across ${shapesUpdated} shapes`;
    });
  } catch (error) {
    console.error("Error in removeReferences:", error);
    throw new Error(`Error removing references: ${error.message}`);
  }
}

export async function removeLinks(deleteAll: boolean = false): Promise<string> {
  try {
    return await PowerPoint.run(async (context) => {
      console.log("removeLinks => start", { deleteAll });
      const textShapes = await loadTextCapableShapes(context);
      console.log(`removeLinks => loaded ${textShapes.length} text-capable shapes`);
      if (textShapes.length === 0) {
        return "No text content available to scrub links";
      }

      const urlRegex = /\b((https?:\/\/)?[\w.-]+(?:\.[\w.-]+)+)\b/g;
      let linksRemovedCount = 0;

      for (const handle of textShapes) {
        const location = describeShape(handle);
        console.log(`removeLinks => inspecting ${location}`);
        const text = await tryLoadShapeText(handle, context);
        if (!text) {
          console.warn(`removeLinks => skipping ${location} (unable to load text)`);
          continue;
        }

        if (!deleteAll && matchesReferenceHeader(text)) {
          continue;
        }

        const matches = text.match(urlRegex) || [];
        if (matches.length === 0) continue;

        linksRemovedCount += matches.length;
        const newText = text
          .replace(urlRegex, "")
          .replace(/\s+([.,;])/g, "$1")
          .replace(/\s{2,}/g, " ")
          .trim();

        handle.shape.textFrame.textRange.text = newText;
        handle.shape.textFrame.textRange.font.bold = false;
        console.log(`removeLinks => updated ${location} (removed ${matches.length} links)`);
      }

      await context.sync();
      console.log(`removeLinks => finished sync (linksRemoved=${linksRemovedCount})`);
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
      console.log("removeWeirdNumbers => start");
      const textShapes = await loadTextCapableShapes(context);
      console.log(`removeWeirdNumbers => loaded ${textShapes.length} text-capable shapes`);
      if (textShapes.length === 0) {
        return "No text content available to clean";
      }

      const weirdNumberPattern = /[【[]\d+.*?[†+t].*?[】\]]\S*/g;
      let totalRemoved = 0;

      for (const handle of textShapes) {
        const location = describeShape(handle);
        console.log(`removeWeirdNumbers => inspecting ${location}`);
        const text = await tryLoadShapeText(handle, context);
        if (!text) {
          console.warn(`removeWeirdNumbers => skipping ${location} (unable to load text)`);
          continue;
        }

        const matches = text.match(weirdNumberPattern) || [];
        if (matches.length === 0) continue;

        totalRemoved += matches.length;
        const newText = text
          .replace(weirdNumberPattern, "")
          .replace(/\s{2,}/g, " ")
          .trim();
        handle.shape.textFrame.textRange.text = newText;
        handle.shape.textFrame.textRange.font.bold = false;
        console.log(`removeWeirdNumbers => updated ${location} (removed ${matches.length} tokens)`);
      }

      await context.sync();
      console.log(`removeWeirdNumbers => finished sync (totalRemoved=${totalRemoved})`);
      return `Removed ${totalRemoved} weird number instances.`;
    });
  } catch (error) {
    console.error("Error in removeWeirdNumbers:", error);
    throw new Error(`Error removing weird numbers: ${error.message}`);
  }
}

export async function normalizeBodyBold(): Promise<string> {
  try {
    return await PowerPoint.run(async (context) => {
      console.log("normalizeBodyBold => start");
      const textShapes = await loadTextCapableShapes(context);
      if (textShapes.length === 0) {
        return "No text content available to normalize";
      }

      let updatedCount = 0;
      let globalIndex = 0;

      for (const handle of textShapes) {
        const location = describeShape(handle);
        console.log(`normalizeBodyBold => inspecting ${location}`);
        const text = await tryLoadShapeText(handle, context);
        if (!text) {
          console.warn(`normalizeBodyBold => skipping ${location} (unable to load text)`);
          continue;
        }

        const sanitized = text.replace(/[\u200B-\u200D\uFEFF]/g, "").trim();
        if (!sanitized) {
          globalIndex++;
          continue;
        }

        const words = sanitized.split(/\s+/).filter(Boolean);
        const meta: ShapeMeta = {
          slideId: handle.slide.id,
          slideIndex: handle.slideIndex,
          shapeId: handle.shape.id,
          shapeIndex: handle.shapeIndex,
          shapeType: handle.shape.type,
          text: sanitized,
          wordCount: words.length,
          isTitle: handle.shapeIndex === 0 && words.length < 10,
          index: globalIndex++,
        };

        if (isHeadingOrTitle(meta)) {
          continue;
        }

        handle.shape.textFrame.textRange.font.bold = false;
        updatedCount++;
      }

      await context.sync();
      const result = `Removed bold formatting from ${updatedCount} text shape${updatedCount === 1 ? "" : "s"}.`;
      console.log(`normalizeBodyBold => completed (${result})`);
      return result;
    });
  } catch (error) {
    console.error("Error in normalizeBodyBold:", error);
    throw new Error(`Error normalizing bold text: ${error.message}`);
  }
}

export async function paraphraseDocument(): Promise<string> {
  try {
    return await PowerPoint.run(async (context) => {
      console.log("paraphraseDocument => start");

      const textShapes = await loadTextCapableShapes(context);
      console.log(`paraphraseDocument => text-capable shapes loaded: ${textShapes.length}`);

      if (!textShapes.length) {
        return "No text content available to paraphrase";
      }

      const handleByShapeId = new Map<string, TextShapeHandle>();
      const metas: ShapeMeta[] = [];
      let globalIndex = 0;

      for (const handle of textShapes) {
        handleByShapeId.set(handle.shape.id, handle);
        const rawText = await tryLoadShapeText(handle, context);
        if (!rawText) {
          continue;
        }

        const sanitized = rawText.replace(/[\u200B-\u200D\uFEFF]/g, "").trim();
        if (!sanitized) {
          continue;
        }

        const words = sanitized.split(/\s+/).filter(Boolean);
        const meta: ShapeMeta = {
          slideId: handle.slide.id,
          slideIndex: handle.slideIndex,
          shapeId: handle.shape.id,
          shapeIndex: handle.shapeIndex,
          shapeType: handle.shape.type,
          text: sanitized,
          wordCount: words.length,
          isTitle: handle.shapeIndex === 0 && words.length < 10,
          index: globalIndex++,
        };

        metas.push(meta);
      }

      console.log(`paraphraseDocument => collected metas: ${metas.length}`);
      if (metas.length === 0) {
        return "No text content found in the presentation";
      }

      const referenceStartIndex = findReferenceStartIndex(metas);
      const { conclusionHeadingIndex, conclusionEndIndex } = findConclusionRange(metas, referenceStartIndex);
      console.log(
        "paraphraseDocument => sections",
        JSON.stringify({ referenceStartIndex, conclusionHeadingIndex, conclusionEndIndex })
      );

      const eligibleMetas = metas.filter((meta) => {
        if (referenceStartIndex !== -1 && meta.index >= referenceStartIndex) return false;
        if (
          conclusionHeadingIndex !== -1 &&
          conclusionEndIndex !== -1 &&
          meta.index > conclusionHeadingIndex &&
          meta.index < conclusionEndIndex
        )
          return false;
        if (matchesReferenceHeader(meta.text)) return false;
        if (isHeadingOrTitle(meta)) return false;
        if (meta.wordCount < 11) return false;
        if (meta.text.trim().endsWith(":")) return false;
        return true;
      });

      console.log(
        "paraphraseDocument => eligible summary",
        eligibleMetas
          .map((meta) => ({
            slide: meta.slideIndex,
            shape: meta.shapeIndex,
            words: meta.wordCount,
            preview: meta.text.substring(0, 80),
          }))
          .slice(0, 5)
      );

      if (eligibleMetas.length === 0) {
        return "No eligible text shapes found to paraphrase.";
      }

      const chunks: string[] = [];
      for (const meta of eligibleMetas) {
        chunks.push(PARAPHRASE_DELIMITER);
        chunks.push(meta.text);
      }
      const payloadText = chunks.join("\n\n");

      console.log(
        `paraphraseDocument => sending ${payloadText.length} chars to paraphrase API (chunks=${eligibleMetas.length})`
      );

      let response;
      try {
        response = await fetch("https://analizeai.com/paraphrase", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ text: payloadText, freeze: [PARAPHRASE_DELIMITER] }),
        });
      } catch (networkError) {
        console.error("paraphraseDocument => network error", networkError);
        throw networkError;
      }

      if (!response.ok) {
        throw new Error(`API request failed with status ${response.status}`);
      }

      const data = await response.json();
      const paraphrasedWholeText = data.secondMode;

      if (!paraphrasedWholeText) {
        throw new Error("Invalid response from paraphrase API");
      }

      const parts = paraphrasedWholeText
        .split(new RegExp(`\\b${PARAPHRASE_DELIMITER}\\b`, "i"))
        .map((part: string) => part.trim())
        .filter((part: string) => part.length > 0);

      if (parts.length !== eligibleMetas.length) {
        console.error(
          `paraphraseDocument => mismatch: sent ${eligibleMetas.length}, received ${parts.length} paraphrased parts`
        );
        throw new Error(
          `Paraphrase count mismatch. Sent ${eligibleMetas.length}, received ${parts.length}. Aborting to prevent data loss.`
        );
      }

      let updatedCount = 0;

      for (let i = 0; i < eligibleMetas.length; i++) {
        const meta = eligibleMetas[i];
        const newText = parts[i];
        const handle = handleByShapeId.get(meta.shapeId);
        if (!handle) {
          console.warn(`paraphraseDocument => missing handle for shape ${meta.shapeId}`);
          continue;
        }

        const shape = handle.shape;

        try {
          shape.textFrame.textRange.text = newText;
          shape.textFrame.textRange.font.bold = false;
          await context.sync();
          updatedCount++;
          console.log(
            `paraphraseDocument => updated slide ${meta.slideIndex} shape ${meta.shapeIndex} (${meta.wordCount} words)`,
            {
              oldPreview: meta.text.substring(0, 80),
              newPreview: newText.substring(0, 80),
            }
          );
        } catch (error) {
          console.error(`paraphraseDocument => failed updating shape ${meta.shapeId}`, error);
        }
      }

      return `Successfully paraphrased ${updatedCount} text shape${updatedCount === 1 ? "" : "s"}.`;
    });
  } catch (error) {
    console.error("Error in paraphraseDocument:", error);
    throw new Error(`Error paraphrasing document: ${error.message}`);
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
