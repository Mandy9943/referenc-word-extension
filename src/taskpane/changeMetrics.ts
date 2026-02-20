export interface CalculatedTextChangeMetrics {
  wordsChanged: number;
  totalOriginalWords: number;
  totalNewWords: number;
  changePercent: number;
  reusedWords: number;
  reusePercent: number;
  addedWords: number;
}

// Split text into language-friendly tokens while preserving contractions.
function tokenizeWords(text: string): string[] {
  const matches = text
    .toLowerCase()
    .match(/[a-z0-9\u00c0-\u024f]+(?:['’-][a-z0-9\u00c0-\u024f]+)*/g);
  if (!matches) {
    return [];
  }

  const normalized: string[] = [];
  for (const token of matches) {
    const normalizedToken =
      typeof token.normalize === "function" ? token.normalize("NFKD") : token;
    const folded = normalizedToken.replace(/[\u0300-\u036f]+/g, "");
    normalized.push(folded);
  }
  return normalized;
}

/**
 * Calculates lexical reuse/change metrics between original and paraphrased texts.
 * - `reusedWords`: original words still present in paraphrased output (frequency-aware)
 * - `wordsChanged`: original words replaced/removed in paraphrased output
 * - `addedWords`: new words introduced by paraphrased output
 */
export function calculateTextChangeMetrics(
  originalTexts: string[],
  newTexts: string[]
): CalculatedTextChangeMetrics {
  const originalWords = tokenizeWords(originalTexts.join(" "));
  const newWords = tokenizeWords(newTexts.join(" "));

  const originalWordFreq = new Map<string, number>();
  for (const word of originalWords) {
    originalWordFreq.set(word, (originalWordFreq.get(word) || 0) + 1);
  }

  const newWordFreq = new Map<string, number>();
  for (const word of newWords) {
    newWordFreq.set(word, (newWordFreq.get(word) || 0) + 1);
  }

  let reusedWords = 0;
  originalWordFreq.forEach((originalCount, word) => {
    const newCount = newWordFreq.get(word) || 0;
    reusedWords += Math.min(originalCount, newCount);
  });

  const totalOriginalWords = originalWords.length;
  const totalNewWords = newWords.length;
  const wordsChanged = Math.max(0, totalOriginalWords - reusedWords);
  const addedWords = Math.max(0, totalNewWords - reusedWords);
  const changePercent =
    totalOriginalWords > 0
      ? Math.round((wordsChanged / totalOriginalWords) * 100)
      : 0;
  const reusePercent =
    totalOriginalWords > 0
      ? Math.round((reusedWords / totalOriginalWords) * 100)
      : 0;

  return {
    wordsChanged,
    totalOriginalWords,
    totalNewWords,
    changePercent,
    reusedWords,
    reusePercent,
    addedWords,
  };
}
