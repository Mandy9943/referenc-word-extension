/* eslint-disable office-addins/no-context-sync-in-loop */
/* global PowerPoint console, Office, fetch, AbortController, setTimeout, clearTimeout */
import { calculateTextChangeMetrics } from "./changeMetrics";
import { getFormattedReferences } from "./gemini";
import type { ParaphraseResult } from "./word";

// ==================== Batch API Configuration ====================
const BATCH_API_URL = "https://analizeai.com/paraphrase-batch";
const HEALTH_CHECK_URL = "https://analizeai.com/health";
const HEALTH_CHECK_TIMEOUT = 2000; // 2 seconds
const ACCOUNT_KEYS = ["acc1", "acc2", "acc3"] as const;

type AccountKey = (typeof ACCOUNT_KEYS)[number];

interface HealthCheckResponse {
  status: string;
  acc1?: { status: string };
  acc2?: { status: string };
  acc3?: { status: string };
  ready?: boolean;
}

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

export async function paraphraseDocument(): Promise<ParaphraseResult> {
  try {
    return await PowerPoint.run(async (context) => {
      console.log("paraphraseDocument => start (current slide only)");

      const allWarnings: string[] = [];
      const healthResult = await checkServiceHealth();
      if (healthResult.warnings.length > 0) {
        console.warn("Health check warnings:", healthResult.warnings);
        allWarnings.push(...healthResult.warnings);
      }

      // Get the currently selected slide
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items");
      await context.sync();

      if (selectedSlides.items.length === 0) {
        return { message: "No slide selected. Please select a slide first.", warnings: allWarnings };
      }

      const currentSlide = selectedSlides.items[0];
      currentSlide.load("shapes");
      await context.sync();

      console.log(`paraphraseDocument => processing slide with ${currentSlide.shapes.items.length} shapes`);

      // Collect text shapes from current slide only
      const eligibleShapes: Array<{ shape: PowerPoint.Shape; text: string; shapeIndex: number }> = [];

      for (let i = 0; i < currentSlide.shapes.items.length; i++) {
        const shape = currentSlide.shapes.items[i];
        shape.load("type, id");
        await context.sync();

        if (!TEXT_CAPABLE_SHAPE_TYPES.has(shape.type)) continue;

        // Load text
        try {
          shape.textFrame.textRange.load("text");
          await context.sync();

          const text = shape.textFrame.textRange.text.replace(/[\u200B-\u200D\uFEFF]/g, "").trim();
          if (!text) continue;

          const words = text.split(/\s+/).filter(Boolean);
          const isTitle = i === 0 && words.length < 10;

          // Skip titles and very short text
          if (isTitle || words.length < 11 || text.trim().endsWith(":")) {
            console.log(`paraphraseDocument => skipping shape ${i} (title or too short)`);
            continue;
          }

          eligibleShapes.push({ shape, text, shapeIndex: i });
          console.log(`paraphraseDocument => queued shape ${i} for paraphrase (${words.length} words)`);
        } catch (error) {
          console.warn(`paraphraseDocument => unable to load text for shape ${i}`, error);
        }
      }

      if (eligibleShapes.length === 0) {
        return { message: "No eligible text found on current slide to paraphrase.", warnings: allWarnings };
      }

      // Calculate total word count and capture original preview
      const originalWordCount = eligibleShapes.reduce((sum, item) => {
        return sum + item.text.split(/\s+/).filter(Boolean).length;
      }, 0);
      const originalPreview = eligibleShapes.length > 0 ? eligibleShapes[0].text.substring(0, 50) + "..." : "";

      // Determine number of accounts to use based on word count
      let numAccounts = 1;
      if (originalWordCount > 1500) {
        numAccounts = 3;
      } else if (originalWordCount >= 500) {
        numAccounts = 2;
      }

      console.log(
        `paraphraseDocument => totalWords=${originalWordCount}, using ${numAccounts} account(s) via batch API`
      );

      const buildPayload = (items: Array<{ text: string }>) => {
        const chunks: string[] = [];
        for (const item of items) {
          chunks.push(PARAPHRASE_DELIMITER);
          chunks.push(item.text);
        }
        return chunks.join("\n\n");
      };

      // Split eligible shapes across accounts
      const chunkSize = Math.ceil(eligibleShapes.length / numAccounts);
      const shapeChunks: Array<Array<{ shape: PowerPoint.Shape; text: string; shapeIndex: number }>> = [];
      for (let i = 0; i < numAccounts; i++) {
        const start = i * chunkSize;
        const end = Math.min(start + chunkSize, eligibleShapes.length);
        shapeChunks.push(eligibleShapes.slice(start, end));
      }

      const batchPayload: Record<string, string> = { mode: "dual" };
      const usedAccounts: AccountKey[] = [];

      shapeChunks.forEach((chunk, idx) => {
        if (chunk.length > 0) {
          const accountKey = ACCOUNT_KEYS[idx];
          batchPayload[accountKey] = buildPayload(chunk);
          usedAccounts.push(accountKey);
          console.log(`paraphraseDocument => ${accountKey}: ${chunk.length} shapes`);
        }
      });

      console.log(`paraphraseDocument => sending batch request (${eligibleShapes.length} shapes total)`);
      const response = await fetch(BATCH_API_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(batchPayload),
      });

      if (!response.ok) {
        throw new Error(`Batch API request failed with status ${response.status}`);
      }

      const data: BatchApiResponse = await response.json();

      const batchWarnings = collectBatchWarnings(data, usedAccounts);
      allWarnings.push(...batchWarnings);

      const paraphrasedChunks: string[] = [];
      for (const key of usedAccounts) {
        const result = data[key];
        if (!result) {
          throw new Error(`No response received for account ${key}`);
        }
        const paraphrasedText = result.secondMode;
        if (!paraphrasedText) {
          if (result.error) {
            throw new Error(`Account ${key} failed: ${result.error}`);
          }
          throw new Error(`No paraphrased text received from account ${key}`);
        }
        paraphrasedChunks.push(paraphrasedText);
      }

      const parseResponse = (text: string, expectedCount: number) => {
        let parts = text
          .split(new RegExp(`${PARAPHRASE_DELIMITER}`, "i"))
          .map((x: string) => x.trim())
          .filter((x: string) => x.length > 0);

        if (parts.length < expectedCount) {
          console.warn(
            `Mismatch detected (Expected ${expectedCount}, got ${parts.length}). Attempting to recover merged parts...`
          );

          const recoveredParts: string[] = [];
          for (const part of parts) {
            if (part.includes("\n\n")) {
              const subParts = part
                .split(/\n\n+/)
                .map((p) => p.trim())
                .filter((p) => p.length > 0);
              recoveredParts.push(...subParts);
            } else {
              recoveredParts.push(part);
            }
          }

          if (recoveredParts.length === expectedCount) {
            console.log("Successfully recovered merged parts!");
            return recoveredParts;
          } else if (Math.abs(recoveredParts.length - expectedCount) < Math.abs(parts.length - expectedCount)) {
            console.log(`Partial recovery: now have ${recoveredParts.length} parts.`);
            return recoveredParts;
          }
        }

        return parts;
      };

      const parsedChunks = shapeChunks.map((chunk, idx) => parseResponse(paraphrasedChunks[idx], chunk.length));

      for (let idx = 0; idx < parsedChunks.length; idx++) {
        const expected = shapeChunks[idx].length;
        const actual = parsedChunks[idx].length;
        if (actual !== expected) {
          console.error(`paraphraseDocument => mismatch in chunk ${idx}: sent ${expected}, received ${actual}`);
          throw new Error(
            `Paraphrase count mismatch. Sent ${eligibleShapes.length} shapes, received ${parsedChunks.reduce((s, p) => s + p.length, 0)}. Aborting to prevent data loss.`
          );
        }
      }

      // Collect original and new texts for comparison
      const originalTexts: string[] = [];
      const newTexts: string[] = [];

      // Apply paraphrased text to shapes
      let updatedCount = 0;
      let newPreview = "";
      let isFirstShape = true;
      for (let chunkIndex = 0; chunkIndex < shapeChunks.length; chunkIndex++) {
        const shapesInChunk = shapeChunks[chunkIndex];
        const parts = parsedChunks[chunkIndex];

        for (let i = 0; i < shapesInChunk.length; i++) {
          const item = shapesInChunk[i];
          const newText = parts[i];

          // Store texts for comparison
          originalTexts.push(item.text);
          newTexts.push(newText);

          // Capture first shape as new preview
          if (isFirstShape) {
            newPreview = newText.substring(0, 50) + "...";
            isFirstShape = false;
          }

          try {
            console.log(`paraphraseDocument => updating shape ${item.shapeIndex}...`);

            item.shape.textFrame.textRange.text = newText;
            item.shape.textFrame.textRange.font.bold = false;
            await context.sync();

            updatedCount++;
            console.log(`paraphraseDocument => ✓ updated shape ${item.shapeIndex}`, {
              oldPreview: item.text.substring(0, 60),
              newPreview: newText.substring(0, 60),
            });
          } catch (error) {
            console.error(`paraphraseDocument => failed updating shape ${item.shapeIndex}`, error);
          }
        }
      }

      // Calculate real change metrics by comparing actual words
      const changeMetrics = calculateTextChangeMetrics(originalTexts, newTexts);

      return {
        message: `Successfully paraphrased ${updatedCount} text box${updatedCount === 1 ? "" : "es"} on current slide.`,
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
    throw new Error(`Error paraphrasing slide: ${error.message}`);
  }
}

export async function paraphraseDocumentStandard(): Promise<ParaphraseResult> {
  try {
    return await PowerPoint.run(async (context) => {
      console.log("paraphraseDocumentStandard => start (current slide only)");

      const allWarnings: string[] = [];
      const healthResult = await checkServiceHealth();
      if (healthResult.warnings.length > 0) {
        console.warn("Health check warnings:", healthResult.warnings);
        allWarnings.push(...healthResult.warnings);
      }

      // Get the currently selected slide
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items");
      await context.sync();

      if (selectedSlides.items.length === 0) {
        throw new Error("No slide selected");
      }

      const currentSlide = selectedSlides.items[0];
      currentSlide.load("shapes");
      await context.sync();

      console.log(`paraphraseDocumentStandard => processing slide with ${currentSlide.shapes.items.length} shapes`);

      // Collect text shapes from current slide only
      const eligibleShapes: Array<{ shape: PowerPoint.Shape; text: string; shapeIndex: number }> = [];

      for (let i = 0; i < currentSlide.shapes.items.length; i++) {
        const shape = currentSlide.shapes.items[i];
        if (!TEXT_CAPABLE_SHAPE_TYPES.has(shape.type)) continue;

        const text = await tryLoadShapeText({ slide: currentSlide, shape, slideIndex: 0, shapeIndex: i }, context);
        if (!text || !text.trim()) continue;

        const trimmed = text.trim();
        const wordCount = trimmed.split(/\s+/).filter(Boolean).length;

        // Skip very short text boxes (likely titles)
        if (wordCount < 15) {
          console.log(`paraphraseDocumentStandard => skipping short shape ${i} (${wordCount} words)`);
          continue;
        }

        // Skip reference headers
        if (matchesReferenceHeader(trimmed)) {
          console.log(`paraphraseDocumentStandard => skipping reference header at shape ${i}`);
          continue;
        }

        eligibleShapes.push({ shape, text: trimmed, shapeIndex: i });
      }

      if (eligibleShapes.length === 0) {
        return {
          message: "No eligible text boxes found on current slide (need at least 15 words)",
          warnings: allWarnings,
        };
      }

      // Calculate total word count and capture original preview
      const originalWordCount = eligibleShapes.reduce((sum, item) => {
        return sum + item.text.split(/\s+/).filter(Boolean).length;
      }, 0);
      const originalPreview = eligibleShapes.length > 0 ? eligibleShapes[0].text.substring(0, 50) + "..." : "";

      let numAccounts = 1;
      if (originalWordCount > 1500) {
        numAccounts = 3;
      } else if (originalWordCount >= 500) {
        numAccounts = 2;
      }

      console.log(
        `paraphraseDocumentStandard => totalWords=${originalWordCount}, using ${numAccounts} account(s) via batch API`
      );

      const buildPayload = (items: Array<{ text: string }>) => {
        const chunks: string[] = [];
        for (const item of items) {
          chunks.push(PARAPHRASE_DELIMITER);
          chunks.push(item.text);
        }
        return chunks.join("\n\n");
      };

      const chunkSize = Math.ceil(eligibleShapes.length / numAccounts);
      const shapeChunks: Array<Array<{ shape: PowerPoint.Shape; text: string; shapeIndex: number }>> = [];
      for (let i = 0; i < numAccounts; i++) {
        const start = i * chunkSize;
        const end = Math.min(start + chunkSize, eligibleShapes.length);
        shapeChunks.push(eligibleShapes.slice(start, end));
      }

      const batchPayload: Record<string, string> = { mode: "standard" };
      const usedAccounts: AccountKey[] = [];

      shapeChunks.forEach((chunk, idx) => {
        if (chunk.length > 0) {
          const accountKey = ACCOUNT_KEYS[idx];
          batchPayload[accountKey] = buildPayload(chunk);
          usedAccounts.push(accountKey);
          console.log(`paraphraseDocumentStandard => ${accountKey}: ${chunk.length} shapes`);
        }
      });

      console.log(`paraphraseDocumentStandard => sending batch request (${eligibleShapes.length} shapes total)`);
      const response = await fetch(BATCH_API_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(batchPayload),
      });

      if (!response.ok) {
        throw new Error(`Batch API request failed with status ${response.status}`);
      }

      const data: BatchApiResponse = await response.json();

      const batchWarnings = collectBatchWarnings(data, usedAccounts);
      allWarnings.push(...batchWarnings);

      const paraphrasedChunks: string[] = [];
      for (const key of usedAccounts) {
        const result = data[key];
        if (!result) {
          throw new Error(`No response received for account ${key}`);
        }
        const paraphrasedText = result.result;
        if (!paraphrasedText) {
          if (result.error) {
            throw new Error(`Account ${key} failed: ${result.error}`);
          }
          throw new Error(`No paraphrased text received from account ${key}`);
        }
        paraphrasedChunks.push(paraphrasedText);
      }

      const parseResponse = (text: string, expectedCount: number) => {
        let parts = text
          .split(new RegExp(`${PARAPHRASE_DELIMITER}`, "i"))
          .map((x: string) => x.trim())
          .filter((x: string) => x.length > 0);

        if (parts.length < expectedCount) {
          console.warn(
            `Mismatch detected (Expected ${expectedCount}, got ${parts.length}). Attempting to recover merged parts...`
          );

          const recoveredParts: string[] = [];
          for (const part of parts) {
            if (part.includes("\n\n")) {
              const subParts = part
                .split(/\n\n+/)
                .map((p) => p.trim())
                .filter((p) => p.length > 0);
              recoveredParts.push(...subParts);
            } else {
              recoveredParts.push(part);
            }
          }

          if (recoveredParts.length === expectedCount) {
            console.log("Successfully recovered merged parts!");
            return recoveredParts;
          } else if (Math.abs(recoveredParts.length - expectedCount) < Math.abs(parts.length - expectedCount)) {
            console.log(`Partial recovery: now have ${recoveredParts.length} parts.`);
            return recoveredParts;
          }
        }

        return parts;
      };

      const parsedChunks = shapeChunks.map((chunk, idx) => parseResponse(paraphrasedChunks[idx], chunk.length));

      for (let idx = 0; idx < parsedChunks.length; idx++) {
        const expected = shapeChunks[idx].length;
        const actual = parsedChunks[idx].length;
        if (actual !== expected) {
          console.error(`paraphraseDocumentStandard => mismatch in chunk ${idx}: sent ${expected}, received ${actual}`);
          throw new Error(
            `Paraphrase count mismatch. Sent ${eligibleShapes.length} text boxes, received ${parsedChunks.reduce((s, p) => s + p.length, 0)}. Aborting to prevent data loss.`
          );
        }
      }

      // Collect original and new texts for comparison
      const originalTexts: string[] = [];
      const newTexts: string[] = [];

      // Apply paraphrased text to shapes
      let updatedCount = 0;
      let newPreview = "";
      let isFirstShape = true;
      for (let chunkIndex = 0; chunkIndex < shapeChunks.length; chunkIndex++) {
        const shapesInChunk = shapeChunks[chunkIndex];
        const parts = parsedChunks[chunkIndex];

        for (let i = 0; i < shapesInChunk.length; i++) {
          const item = shapesInChunk[i];
          const newText = parts[i];

          // Store texts for comparison
          originalTexts.push(item.text);
          newTexts.push(newText);

          // Capture first shape as new preview
          if (isFirstShape) {
            newPreview = newText.substring(0, 50) + "...";
            isFirstShape = false;
          }

          try {
            item.shape.textFrame.textRange.text = newText;
            item.shape.textFrame.textRange.font.bold = false;
            await context.sync();

            console.log(
              `paraphraseDocumentStandard => updated shape ${item.shapeIndex}\nOLD: ${item.text.substring(0, 80)}...\nNEW: ${newText.substring(0, 80)}...`
            );
            updatedCount++;
          } catch (updateError) {
            console.error(`paraphraseDocumentStandard => failed to update shape ${item.shapeIndex}:`, updateError);
          }
        }
      }

      // Calculate real change metrics by comparing actual words
      const changeMetrics = calculateTextChangeMetrics(originalTexts, newTexts);

      return {
        message: `Successfully paraphrased ${updatedCount} text box${updatedCount === 1 ? "" : "es"} on current slide (Standard mode).`,
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
    console.error("Error in paraphraseDocumentStandard:", error);
    throw new Error(`Error paraphrasing slide: ${error.message}`);
  }
}

export async function paraphraseSelectedText(): Promise<ParaphraseResult> {
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

      const allWarnings: string[] = [];
      const healthResult = await checkServiceHealth();
      if (healthResult.warnings.length > 0) {
        console.warn("Health check warnings:", healthResult.warnings);
        allWarnings.push(...healthResult.warnings);
      }

      try {
        const response = await fetch(BATCH_API_URL, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ acc1: selectedText.trim(), mode: "dual" }),
        });

        if (!response.ok) {
          throw new Error(`Batch API request failed with status ${response.status}`);
        }

        const data: BatchApiResponse = await response.json();
        const batchWarnings = collectBatchWarnings(data, ["acc1"]);
        allWarnings.push(...batchWarnings);

        const resultAcc1 = data.acc1;
        if (!resultAcc1) {
          throw new Error("No response received for account acc1");
        }
        const paraphrasedText = resultAcc1.secondMode;
        if (!paraphrasedText) {
          if (resultAcc1.error) {
            throw new Error(`Account acc1 failed: ${resultAcc1.error}`);
          }
          throw new Error("Invalid batch response: missing secondMode");
        }

        Office.context.document.setSelectedDataAsync(paraphrasedText, (setResult) => {
          if (setResult.status === Office.AsyncResultStatus.Failed) {
            reject(new Error(setResult.error.message));
          } else {
            resolve({ message: "Text paraphrased successfully", warnings: allWarnings });
          }
        });
      } catch (error) {
        reject(error);
      }
    });
  });
}
