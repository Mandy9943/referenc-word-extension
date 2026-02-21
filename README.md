## Essay Manager – Reference Automation Add‑in

Essay Manager is a Word and PowerPoint task pane add‑in that helps writers clean, cite, and paraphrase academic content directly inside Office. It blends a React + Fluent UI interface with host-specific Office.js scripts and a couple of curated AI services to automate inserting references, stripping unwanted citations, and paraphrasing selected passages.

### Table of contents

1. [Key capabilities](#key-capabilities)
2. [How it works](#how-it-works)
3. [Task pane buttons (what & how)](#task-pane-buttons-what--how)
4. [External services & configuration](#external-services--configuration)
5. [Getting started](#getting-started)
6. [Everyday usage](#everyday-usage)
7. [Bulk DOCX automation](#bulk-docx-automation)
8. [Bulk PPTX automation](#bulk-pptx-automation)
9. [Testing & quality](#testing--quality)
10. [Troubleshooting & tips](#troubleshooting--tips)
11. [Contributing & license](#contributing--license)

## Key capabilities

| Feature | Word | PowerPoint | Description |
| --- | --- | --- | --- |
| **Add References** (`analyzeDocument`) | ✅ | ✅ | Detects the “Reference List” section, reformats entries with Gemini, and weaves randomized `(Author, Year)` citations back into paragraph text or shapes while skipping headings/TOC entries. |
| **Clean Document** (`removeReferences`, `removeLinks`, `removeWeirdNumbers`) | ✅ | ⚪️ | Sequentially removes inline citations, URL-like strings outside the references section, and clipboard artifacts such as ``. |
| **Paraphrase Document** (`paraphraseDocument`) | ✅ | ✅ | Sends all eligible body paragraphs or text shapes to AnalizeAI, then replaces them with the returned `secondMode` paraphrase while skipping headings, titles, and reference sections. |
| **Insert Text helpers** (`insertText`) | ✅ | ✅ | Demo helpers per host used by tests and potential ribbon commands. |

> ⚪️ = feature not available on that host yet. The add-in automatically checks `Office.context.host` before executing a command to avoid unsupported scenarios.

## How it works

### High-level flow

1. **Task pane UI (`src/taskpane/components`)** – A React app styled with Fluent UI surfaces the visible actions (Add References, Clean, Paraphrase) plus the *Insert references every other paragraph* toggle.
2. **Host router (`src/taskpane/taskpane.ts`)** – Each button calls into a host-aware wrapper that runs only after `Office.onReady`. It forwards the call to the Word or PowerPoint implementation and throws when the host is unsupported.
3. **Host implementations** – Files inside `src/taskpane/<host>.ts` use Office.js APIs to read and mutate document content:
	- `word.ts` walks `context.document.body.paragraphs`, extracts reference sections, calls Gemini formatting, removes citations/links/odd markers, and performs selection-based paraphrasing.
	- `powerpoint.ts` iterates slides/shapes, loads each `textFrame`, and performs similar reference injection logic adapted to shapes.
4. **AI helpers** – Utility modules encapsulate third-party services:
	- `gemini.ts` formats reference blocks using Google Generative AI (Gemini 2.5 Flash Lite).
	- The Word/PowerPoint paraphrase routines POST to `https://analizeai.com/paraphrase`, preserving a frozen delimiter so each returned `secondMode` block maps back to the originating paragraph or shape.
5. **Feedback loop** – The React layer shows loading / success / error banners, a stopwatch for long paraphrase calls, and automatically resets UI state once asynchronous work finishes.

### Notable implementation details

* **Reference detection** – Multiple header spellings (“Reference List”, “REFERENCES”, etc.) are supported. Only the last occurrence is considered so appendices remain untouched.
* **Smart filtering** – Before injecting citations, the add-in filters out paragraphs that look like TOC entries, heading lines ending with `:`, or those shorter than 11 words.
* **Host parity** – Word and PowerPoint share the same action names but use host-specific object models (`Word.run`, `PowerPoint.run`). Other hosts (Excel, Outlook, OneNote, Project) retain demo `insertText` shims for future growth.

## Task pane buttons (what & how)

| Button | Visible behavior | Under the hood |
| --- | --- | --- |
| **Clean** | Single magenta button; status pill switches from “Processing…” to green success (or red error) after the pipeline finishes. | 1) `removeReferences()` loops through every paragraph, matches APA-like inline citations with several regexes, removes them, and fixes doubled spaces/periods.<br>2) `removeLinks(false)` rebuilds the paragraph list, slices everything **before** the references section, and strips bare URLs using a boundary-aware regex so punctuation stays intact.<br>3) `removeWeirdNumbers()` looks for clipboard artifacts such as ``, deletes them, and normalizes whitespace. Each function runs inside its own `Word.run`, so text is updated immediately before the next stage begins. |
| **Paraphrase** | Lilac button paraphrases the entire Word document or PowerPoint presentation; no selection required. The stopwatch under the buttons still tracks elapsed API time for transparency. | 1) The host-specific routine (`paraphraseDocument`) enumerates all paragraphs/shapes, skips headings, titles, short snippets, and anything inside/after the references slide/section, and builds a delimiter-separated payload.<br>2) Issues `fetch("https://analizeai.com/paraphrase")` with `{ text, freeze: ["qbpdelim123"] }` so the response can be split safely.<br>3) Validates that the returned `secondMode` block count matches the number of eligible paragraphs/shapes; if so, replaces each text range and clears bold formatting.<br>4) React timer logic mirrors the long-running action status so users know how long the paraphrase call took. |
| **Add References** | Green primary button with a helper checkbox “Insert references every other paragraph”. When done, the status pill announces success and Word/PowerPoint content now contains inline citations. | 1) `Word.run` (or `PowerPoint.run`) loads every paragraph/shape, building a plain-text copy of the whole document/presentation.<br>2) Searches from the bottom up for any header variant (“Reference List”, “REFERENCES”, etc.); if none are found the action exits early.<br>3) Slices the text from that header to the end and feeds the raw block to `getFormattedReferences()`, which calls Gemini 2.5 Flash Lite. The AI responds with blank-line-separated `(Author, Year)` entries that we split and trim.<br>4) Builds a list of eligible paragraphs by excluding anything shorter than 11 words, ending with `:`, or matching TOC patterns (dot leaders or tabbed page numbers). If the checkbox is checked, the code keeps only every other eligible index to reduce density.<br>5) Randomizes the eligible index list, then for each target paragraph inserts one of the formatted citations at the end (or before the trailing period). Used citations are tracked so Gemini outputs are reused evenly before repeating. PowerPoint follows the same steps but operates on shape text ranges instead of body paragraphs. |

No other buttons are wired today—there is intentionally **no** humanize or DeepSeek functionality left in the codebase, which keeps configuration limited to Gemini + AnalizeAI.

## External services & configuration

| Service | File(s) | Purpose | Required variables |
| --- | --- | --- | --- |
| Google Generative AI (Gemini 2.5 Flash Lite) | `src/taskpane/gemini.ts` | Rewrites raw reference sections into `(Author, Year)` snippets. | `GEMINI_API_KEY` |
| AnalizeAI paraphrase API | `src/taskpane/word.ts`, `src/taskpane/powerpoint.ts` | Remotely paraphrases eligible body text or shapes. | _none_ (public HTTPS call) |

Create a `.env` (or `.env.local`) in the project root before running dev builds:

```bash
GEMINI_API_KEY=AIza...
```

> The bundler injects environment variables via `dotenv-webpack`. Never expose production keys in source control—use local `.env` files or CI secrets.

## Getting started

1. **Prerequisites**
	- Node.js 16.x–20.x and npm 7–10 (see `package.json` `engines`).
	- Office desktop (Word or PowerPoint) or Office on the web with sideloading enabled.
	- Optional: Microsoft 365 Developer Program subscription for sandbox tenants.

2. **Install dependencies**
	```bash
	npm install
	```

3. **Run the dev server (hot reload for the React pane)**
	```bash
	npm run dev-server
	```

4. **Sideload into Word (desktop)**
	```bash
	npm run start:desktop -- --app word
	```
	The script uses `manifest.xml`. Additional manifests exist for Excel, PowerPoint, Outlook, OneNote, and Project in the repo root.

5. **Stop debugging**
	```bash
	npm run stop
	```

6. **Build production assets**
	```bash
	npm run build
	```

## Everyday usage

1. Launch Word or PowerPoint with the add-in loaded and open the **Essay Manager** task pane.
2. Choose an action:
	- **Add References** – Optionally enable *Insert references every other paragraph*, then press the green button. The add-in reads the document, calls Gemini to normalize references, and injects citations while logging progress to the console for troubleshooting.
	- **Clean** – Runs `removeReferences`, `removeLinks`, and `removeWeirdNumbers` in sequence. Useful after pasting AI output or bibliographies from PDFs.
	- **Paraphrase** – Click **Paraphrase** to rewrite every eligible paragraph (Word) or text shape (PowerPoint). The live stopwatch shows how long the AnalizeAI round-trip takes while references, headings, and short titles are automatically skipped.
3. Review the status indicator: grey (idle), brand blue (processing), green (success), or red (error). Errors are also printed in the Office task pane console (Edge DevTools).

### Behind the scenes

* All Word mutations happen within a single `Word.run` batch to keep context state consistent.
* PowerPoint text updates re-load each `textFrame.textRange` to avoid stale object errors.
* The `Clean` pipeline is intentionally sequential: removing citations first prevents dangling periods before URL removal and weird-number cleanup.

## Bulk DOCX automation

For one-command DOCX processing, use the offline script pipeline. It enforces the exact sequence:

1. Clean weird artifacts + existing in-text citations (without touching headings/subtitles)
2. Paraphrase body content (`SIMPLE+SHORT` by default)
3. Detect reference section and insert new in-text references
4. Write output as `pr <original-name>.docx`

Commands:

```bash
# SIMPLE+SHORT mode
npm run doc

# STANDARD mode
npm run doc standard
```

Input behavior:

* If no file path is passed, the script auto-selects a single `.docx` on Desktop.
* It ignores temporary Word lock files (`~$*.docx`) and generated `pr *.docx` outputs when the original source file is also present.

Output behavior:

* Default output is `pr <input-name>.docx` in the same folder as the input file.
* If references cannot be detected or inserted, the script exits with a terminal error and does **not** write an output file.

Useful options:

```bash
python3 scripts/paraphrase_docx.py --mode dual "/path/in.docx"
python3 scripts/paraphrase_docx.py --mode standard "/path/in.docx"
python3 scripts/paraphrase_docx.py standard "/path/in.docx" --dry-run
```

## Bulk PPTX automation

For one-command PPTX processing, use the offline script pipeline. It enforces the exact sequence:

1. Clean weird artifacts + existing in-text citations (without touching headings/subtitles or detected references section)
2. Paraphrase slide text and speaker notes (`SIMPLE+SHORT` by default)
3. Detect references and insert new in-text citations into both slides and notes
4. Write output as `pr <original-name>.pptx`

Commands:

```bash
# SIMPLE+SHORT mode
npm run pptx

# STANDARD mode
npm run pptx standard
```

Input behavior:

* If no file path is passed, the script auto-selects a single `.pptx` on Desktop.
* It ignores temporary PowerPoint lock files (`~$*.pptx`) and generated `pr *.pptx` outputs when the original source file is also present.

Output behavior:

* Default output is `pr <input-name>.pptx` in the same folder as the input file.
* If references cannot be detected or inserted, the script exits with a terminal error and does **not** write an output file.

Useful options:

```bash
python3 scripts/paraphrase_pptx.py "/path/in.pptx"
python3 scripts/paraphrase_pptx.py standard "/path/in.pptx"
python3 scripts/paraphrase_pptx.py --mode dual "/path/in.pptx" --dry-run
python3 scripts/paraphrase_pptx.py --no-notes "/path/in.pptx"
```

## Testing & quality

| Command | Purpose |
| --- | --- |
| `npm run test:unit` | Runs Mocha tests under `test/unit` that exercise host logic via the Office mock runtime. |
| `npm run test:e2e` | Launches Mocha + Playwright-style helpers in `test/end-to-end` to verify manifest + UI wiring. |
| `npm run lint` / `npm run lint:fix` | Applies the Office Add-in ESLint + Prettier config to TypeScript/React files. |

Developers should run lint + unit tests before opening a PR; E2E coverage is slower but recommended after major host-behavior changes.

## Troubleshooting & tips

- **Add-in fails to sideload** – Run `npm run convert-to-single-host` to shrink the manifest to a single host, or double-check that Office “Trusted Add-in Catalog” is configured.
- **AI calls hang** – Confirm environment variables are exposed to the bundler and that your tenant/network allows outbound HTTPS to Google AI (Gemini) and AnalizeAI.
- **Office.js object invalid** – Ensure every `load` is followed by `await context.sync()` before accessing properties, especially when iterating PowerPoint shapes.

## Contributing & license

Pull requests and issues are welcome. Please review **CONTRIBUTING.md**, **SECURITY.md**, and the Microsoft Open Source Code of Conduct before submitting changes.

This project remains under the MIT license (see `LICENSE`).
