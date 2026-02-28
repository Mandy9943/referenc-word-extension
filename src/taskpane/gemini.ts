const GEMINI_FORMAT_ENDPOINT = "/api/gemini-format";

export async function getFormattedReferences(references: string): Promise<string> {
  const response = await fetch(GEMINI_FORMAT_ENDPOINT, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ references }),
  });

  let payload: any = null;
  try {
    payload = await response.json();
  } catch {
    payload = null;
  }

  if (!response.ok) {
    const detail = payload?.error || `Gemini formatter request failed with status ${response.status}`;
    throw new Error(String(detail));
  }

  const text = payload?.text;
  if (!text || typeof text !== "string") {
    throw new Error("Gemini formatter returned an invalid response.");
  }

  return text;
}
