import { insertText as insertTextInExcel } from "./excel";
import { insertText as insertTextInOneNote } from "./onenote";
import { insertText as insertTextInOutlook } from "./outlook";
import {
  analyzeDocument as analyzeDocumentInPowerPoint,
  insertText as insertTextInPowerPoint,
  normalizeBodyBold as normalizeBodyBoldInPowerPoint,
  paraphraseDocument as paraphraseDocumentInPowerPoint,
  paraphraseDocumentStandard,
  paraphraseSelectedText as paraphraseSelectedTextInPowerPointSelection,
  paraphraseSelectedTextStandard as paraphraseSelectedTextStandardInPowerPointSelection,
  removeLinks as removeLinksInPowerPoint,
  removeReferences as removeReferencesInPowerPoint,
  removeWeirdNumbers as removeWeirdNumbersInPowerPoint,
} from "./powerpoint";
import { insertText as insertTextInProject } from "./project";
import {
  analyzeDocument as analyzeDocumentInWord,
  ChangeMetrics,
  insertText as insertTextInWord,
  normalizeBodyBold as normalizeBodyBoldInWord,
  paraphraseDocument as paraphraseDocumentInWord,
  paraphraseDocumentStandard as paraphraseDocumentStandardInWord,
  ParaphraseResult,
  removeLinks as removeLinksInWord,
  removeReferences as removeReferencesInWord,
  removeWeirdNumbers as removeWeirdNumbersInWord,
} from "./word";

// Re-export the types for use in components
export type { ChangeMetrics, ParaphraseResult };

/* global Office */

export async function insertText(text: string) {
  Office.onReady(async (info) => {
    switch (info.host) {
      case Office.HostType.Excel:
        await insertTextInExcel(text);
        break;
      case Office.HostType.OneNote:
        await insertTextInOneNote(text);
        break;
      case Office.HostType.Outlook:
        await insertTextInOutlook(text);
        break;
      case Office.HostType.Project:
        await insertTextInProject(text);
        break;
      case Office.HostType.PowerPoint:
        await insertTextInPowerPoint(text);
        break;
      case Office.HostType.Word:
        await insertTextInWord(text);
        break;
      default: {
        throw new Error("Don't know how to insert text when running in ${info.host}.");
      }
    }
  });
}

export async function analyzeDocument(insertEveryOther: boolean = false) {
  await Office.onReady();
  switch (Office.context.host) {
    case Office.HostType.Word:
      return await analyzeDocumentInWord(insertEveryOther);
    case Office.HostType.PowerPoint:
      return await analyzeDocumentInPowerPoint(insertEveryOther);
    default:
      throw new Error("This function is only available in Word and PowerPoint");
  }
}

export async function removeReferences() {
  await Office.onReady();
  switch (Office.context.host) {
    case Office.HostType.Word:
      return await removeReferencesInWord();
    case Office.HostType.PowerPoint:
      return await removeReferencesInPowerPoint();
    default:
      throw new Error("This function is only available in Word and PowerPoint");
  }
}

export async function removeLinks(deleteAll: boolean = false) {
  await Office.onReady();
  switch (Office.context.host) {
    case Office.HostType.Word:
      return await removeLinksInWord(deleteAll);
    case Office.HostType.PowerPoint:
      return await removeLinksInPowerPoint(deleteAll);
    default:
      throw new Error("This function is only available in Word and PowerPoint");
  }
}

export async function removeWeirdNumbers() {
  await Office.onReady();
  switch (Office.context.host) {
    case Office.HostType.Word:
      return await removeWeirdNumbersInWord();
    case Office.HostType.PowerPoint:
      return await removeWeirdNumbersInPowerPoint();
    default:
      throw new Error("This function is only available in Word and PowerPoint");
  }
}

export async function paraphraseSelectedText(): Promise<ParaphraseResult> {
  await Office.onReady();
  switch (Office.context.host) {
    case Office.HostType.Word:
      return await paraphraseDocumentInWord();
    case Office.HostType.PowerPoint: {
      try {
        return await paraphraseSelectedTextInPowerPointSelection();
      } catch (error) {
        if (shouldFallbackToSlideParaphrase(error)) {
          return await paraphraseDocumentInPowerPoint();
        }
        throw error;
      }
    }
    default:
      throw new Error("This function is only available in Word and PowerPoint");
  }
}

export async function paraphraseSelectedTextStandard(): Promise<ParaphraseResult> {
  await Office.onReady();
  switch (Office.context.host) {
    case Office.HostType.Word:
      return await paraphraseDocumentStandardInWord();
    case Office.HostType.PowerPoint: {
      try {
        return await paraphraseSelectedTextStandardInPowerPointSelection();
      } catch (error) {
        if (shouldFallbackToSlideParaphrase(error)) {
          return await paraphraseDocumentStandard();
        }
        throw error;
      }
    }
    default:
      throw new Error("This function is only available in Word and PowerPoint");
  }
}

function shouldFallbackToSlideParaphrase(error: unknown): boolean {
  const message = String((error as Error)?.message || "").toLowerCase();
  const selectionFallbackMarkers = [
    "no text selected",
    "there is no text selected",
    "coercion type",
    "cannot coerce",
    "not a text selection",
  ];

  return (
    selectionFallbackMarkers.some((marker) => message.includes(marker))
  );
}

export async function normalizeBoldText() {
  await Office.onReady();
  switch (Office.context.host) {
    case Office.HostType.Word:
      return await normalizeBodyBoldInWord();
    case Office.HostType.PowerPoint:
      return await normalizeBodyBoldInPowerPoint();
    default:
      throw new Error("This function is only available in Word and PowerPoint");
  }
}
