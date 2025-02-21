import { insertText as insertTextInExcel } from "./excel";
import { insertText as insertTextInOneNote } from "./onenote";
import { insertText as insertTextInOutlook } from "./outlook";
import {
  analyzeDocument as analyzeDocumentInPowerPoint,
  humanizeDocument as humanizeDocumentInPowerPoint,
  insertText as insertTextInPowerPoint,
  removeReferences as removeReferencesInPowerPoint,
} from "./powerpoint";
import { insertText as insertTextInProject } from "./project";
import {
  analyzeDocument as analyzeDocumentInWord,
  humanizeDocument as humanizeDocumentInWord,
  humanizeSelectedTextInWord,
  insertText as insertTextInWord,
  removeReferences as removeReferencesInWord,
  requestCancelHumanize,
} from "./word";

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

export async function analyzeDocument() {
  await Office.onReady();
  switch (Office.context.host) {
    case Office.HostType.Word:
      return await analyzeDocumentInWord();
    case Office.HostType.PowerPoint:
      return await analyzeDocumentInPowerPoint();
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

export async function humanizeDocument() {
  await Office.onReady();
  switch (Office.context.host) {
    case Office.HostType.Word:
      return await humanizeDocumentInWord();
    case Office.HostType.PowerPoint:
      return await humanizeDocumentInPowerPoint();
    default:
      throw new Error("This function is only available in Word and PowerPoint");
  }
}

export async function humanizeSelectedText() {
  await Office.onReady();
  switch (Office.context.host) {
    case Office.HostType.Word:
      return await humanizeSelectedTextInWord();
    default:
      throw new Error("This function is only available in Word");
  }
}

/**
 * Synchronously stops the in-progress "humanize" operation in Word.
 */
export function stopHumanizeProcess() {
  // This will trigger the cancel logic in "word.ts"
  requestCancelHumanize();
}
