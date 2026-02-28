#!/usr/bin/env node
"use strict";

const express = require("express");
const multer = require("multer");
const os = require("os");
const path = require("path");
const crypto = require("crypto");
const fsSync = require("fs");
const fs = require("fs/promises");
const { spawn } = require("child_process");

const repoRoot = path.resolve(__dirname, "..");
const webRoot = path.join(__dirname, "workflow-web");
const stagingRoot = path.join(os.tmpdir(), "referenc-workflow-upload-staging");
const jobsRoot = path.join(os.tmpdir(), "referenc-workflow-jobs");
const pptToDocScript = "/Users/andreibucur/tools/pptx_to_docx_with_notes.py";
const GEMINI_MODEL = "gemini-2.5-flash-lite";
const GEMINI_API_URL = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;
const GEMINI_API_KEY = String(process.env.GEMINI_API_KEY || "").trim();

const workflows = {
  "ppt-to-doc": {
    label: "ppt to doc",
    accepts: [".pptx"],
    allowMultiple: false,
  },
  "ppt-to-doc-batch": {
    label: "ppt to doc --batch",
    accepts: [".pptx"],
    allowMultiple: true,
  },
  "doc-dual": {
    label: "npm run doc (SIMPLE+SHORT)",
    accepts: [".docx"],
    allowMultiple: true,
  },
  "doc-standard": {
    label: "npm run doc:standard",
    accepts: [".docx"],
    allowMultiple: true,
  },
  "ppt-dual": {
    label: "npm run ppt (SIMPLE+SHORT)",
    accepts: [".pptx"],
    allowMultiple: true,
  },
  "ppt-standard": {
    label: "npm run ppt:standard",
    accepts: [".pptx"],
    allowMultiple: true,
  },
};

const app = express();
const port = Number(process.env.WORKFLOW_WEB_PORT || 4312);

fsSync.mkdirSync(stagingRoot, { recursive: true });
fsSync.mkdirSync(jobsRoot, { recursive: true });

const upload = multer({
  dest: stagingRoot,
  limits: {
    fileSize: 200 * 1024 * 1024,
    files: 50,
  },
});

app.use(express.json());
app.use(express.static(webRoot));

function buildGeminiPrompt(references) {
  return `
Rewrite the following list of references like this: (author name, year).
Put blank space between references and write them one below the other.
Don't number them and don't use bullet points.
Never add intro text like: "Here is the list of references formatted in the style you requested".
If the list includes these words, never output them: Books, Journal Articles, Websites, Additional References.
Always use citation format: (author name, year), for example: (Buzan, 2010).

Example:
\`\`\`
(Buzan, 2010)

(Buzan, 2010)
\`\`\`

Here's the list:
${references}
`.trim();
}

async function callGeminiFormatter(references) {
  if (!GEMINI_API_KEY) {
    const error = new Error("Gemini API key is not configured on server.");
    error.statusCode = 500;
    throw error;
  }

  const prompt = buildGeminiPrompt(references);
  const response = await fetch(`${GEMINI_API_URL}?key=${encodeURIComponent(GEMINI_API_KEY)}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
    }),
  });

  let payload = null;
  try {
    payload = await response.json();
  } catch (_) {
    payload = null;
  }

  if (!response.ok) {
    const message =
      payload?.error?.message ||
      `Gemini request failed with HTTP ${response.status}`;
    const error = new Error(message);
    error.statusCode = response.status >= 400 && response.status < 500 ? 400 : 502;
    throw error;
  }

  const text =
    payload?.candidates?.[0]?.content?.parts
      ?.map((part) => (typeof part?.text === "string" ? part.text : ""))
      .join("")
      .trim() || "";

  if (!text) {
    const error = new Error("Gemini returned no text.");
    error.statusCode = 502;
    throw error;
  }

  return text;
}

function safeBaseName(name) {
  const parsed = path.parse(name || "file");
  const raw = parsed.name || "file";
  return raw.replace(/[^A-Za-z0-9._ -]/g, " ").replace(/\s+/g, " ").trim() || "file";
}

function normalizeExt(filePath) {
  return path.extname(filePath || "").toLowerCase();
}

async function runCommand(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: options.cwd || repoRoot,
      env: { ...process.env, ...(options.env || {}) },
      stdio: ["ignore", "pipe", "pipe"],
    });

    let stdout = "";
    let stderr = "";

    child.stdout.on("data", (chunk) => {
      const text = chunk.toString();
      stdout += text;
      if (options.onOutput) options.onOutput(text, "stdout");
    });

    child.stderr.on("data", (chunk) => {
      const text = chunk.toString();
      stderr += text;
      if (options.onOutput) options.onOutput(text, "stderr");
    });

    child.on("error", (error) => reject(error));

    child.on("close", (code) => {
      if (code === 0) {
        resolve({ stdout, stderr, code });
        return;
      }
      const summary = [
        `Command failed (${code}): ${command} ${args.join(" ")}`,
        stderr.trim(),
      ]
        .filter(Boolean)
        .join("\n\n");
      const error = new Error(summary);
      error.code = code;
      error.stdout = stdout;
      error.stderr = stderr;
      reject(error);
    });
  });
}

async function removePath(targetPath) {
  if (!targetPath) return;
  try {
    await fs.rm(targetPath, { recursive: true, force: true });
  } catch (_) {
    // best-effort cleanup
  }
}

function outputNameForParaphrase(originalName) {
  return `pr ${originalName}`;
}

async function prepareUploadedFiles(reqFiles, workflowConfig, inputDir) {
  const files = reqFiles || [];
  if (files.length === 0) {
    throw new Error("Please upload at least one file.");
  }

  if (!workflowConfig.allowMultiple && files.length > 1) {
    throw new Error(`${workflowConfig.label} accepts exactly one file.`);
  }

  const prepared = [];
  const usedNames = new Set();

  for (let i = 0; i < files.length; i += 1) {
    const file = files[i];
    const ext = normalizeExt(file.originalname || file.filename);
    if (!workflowConfig.accepts.includes(ext)) {
      throw new Error(
        `Invalid file type for ${workflowConfig.label}. Allowed: ${workflowConfig.accepts.join(", ")}`
      );
    }

    let name = `${safeBaseName(file.originalname || file.filename)}${ext}`;
    if (usedNames.has(name.toLowerCase())) {
      name = `${safeBaseName(file.originalname || file.filename)}-${i + 1}${ext}`;
    }
    usedNames.add(name.toLowerCase());

    const destination = path.join(inputDir, name);
    await fs.rename(file.path, destination);
    prepared.push({ originalName: file.originalname, path: destination, name, ext });
  }

  return prepared;
}

async function runPptToDoc(files, logger) {
  if (!fsSync.existsSync(pptToDocScript)) {
    throw new Error(`Missing converter script: ${pptToDocScript}`);
  }

  const inputPaths = files.map((file) => file.path);
  const args = [
    "run",
    "--with",
    "python-docx",
    "--with",
    "python-pptx",
    pptToDocScript,
    ...inputPaths,
  ];

  logger(`[runner] uv ${args.join(" ")}`);
  await runCommand("uv", args, {
    cwd: repoRoot,
    onOutput: (text) => logger(text.trimEnd()),
  });

  const outputs = [];
  for (const file of files) {
    const parsed = path.parse(file.path);
    const outputPath = path.join(parsed.dir, `${parsed.name}_slides_notes.docx`);
    if (!fsSync.existsSync(outputPath)) {
      throw new Error(`Conversion output missing for ${path.basename(file.path)}.`);
    }
    outputs.push(outputPath);
  }

  return outputs;
}

async function runDocWorkflow(files, mode, outputsDir, logger) {
  const outputs = [];

  for (const file of files) {
    const outputPath = path.join(outputsDir, outputNameForParaphrase(file.name));
    const args = [
      path.join("scripts", "paraphrase_docx.py"),
      "--mode",
      mode,
      file.path,
      "-o",
      outputPath,
    ];

    logger(`[runner] python3 ${args.join(" ")}`);
    await runCommand("python3", args, {
      cwd: repoRoot,
      onOutput: (text) => logger(text.trimEnd()),
    });

    if (!fsSync.existsSync(outputPath)) {
      throw new Error(`DOC output missing for ${file.name}.`);
    }

    outputs.push(outputPath);
  }

  return outputs;
}

async function runPptWorkflow(files, mode, outputsDir, logger) {
  const outputs = [];

  for (const file of files) {
    const outputPath = path.join(outputsDir, outputNameForParaphrase(file.name));
    const args = [
      path.join("scripts", "paraphrase_pptx.py"),
      "--mode",
      mode,
      file.path,
      "-o",
      outputPath,
    ];

    logger(`[runner] python3 ${args.join(" ")}`);
    await runCommand("python3", args, {
      cwd: repoRoot,
      onOutput: (text) => logger(text.trimEnd()),
    });

    if (!fsSync.existsSync(outputPath)) {
      throw new Error(`PPT output missing for ${file.name}.`);
    }

    outputs.push(outputPath);
  }

  return outputs;
}

async function createZipArchive(outputPath, files, logger) {
  if (files.length === 0) {
    throw new Error("Nothing to archive.");
  }

  const args = ["-jq", outputPath, ...files];
  logger(`[runner] zip ${args.join(" ")}`);
  await runCommand("zip", args, {
    cwd: repoRoot,
    onOutput: (text) => logger(text.trimEnd()),
  });

  if (!fsSync.existsSync(outputPath)) {
    throw new Error("ZIP archive generation failed.");
  }
}

app.get("/api/health", (_req, res) => {
  res.json({
    ok: true,
    workflows: Object.keys(workflows),
    geminiConfigured: Boolean(GEMINI_API_KEY),
  });
});

app.post("/api/gemini-format", async (req, res) => {
  const references = typeof req.body?.references === "string" ? req.body.references.trim() : "";
  if (!references) {
    res.status(400).json({ error: "Missing references text." });
    return;
  }

  try {
    const text = await callGeminiFormatter(references);
    res.json({ text });
  } catch (error) {
    const statusCode = Number.isInteger(error?.statusCode) ? error.statusCode : 500;
    res.status(statusCode).json({ error: error?.message || "Gemini formatting failed." });
  }
});

app.post("/api/process", upload.array("files", 50), async (req, res) => {
  const workflowKey = String(req.body.workflow || "").trim();
  const workflow = workflows[workflowKey];
  const stagedPaths = (req.files || []).map((file) => file.path);

  const requestId = crypto.randomUUID();
  const jobDir = path.join(jobsRoot, requestId);
  const inputDir = path.join(jobDir, "inputs");
  const outputDir = path.join(jobDir, "outputs");
  const logs = [];

  const log = (line) => {
    if (!line) return;
    logs.push(line);
    console.log(`[workflow:${requestId}] ${line}`);
  };

  try {
    if (!workflow) {
      throw new Error("Invalid workflow selected.");
    }

    await fs.mkdir(inputDir, { recursive: true });
    await fs.mkdir(outputDir, { recursive: true });

    const uploadedFiles = await prepareUploadedFiles(req.files, workflow, inputDir);

    log(`workflow=${workflow.label}`);
    log(`files=${uploadedFiles.map((f) => f.name).join(", ")}`);

    let outputs = [];

    if (workflowKey === "ppt-to-doc" || workflowKey === "ppt-to-doc-batch") {
      outputs = await runPptToDoc(uploadedFiles, log);
    } else if (workflowKey === "doc-dual") {
      outputs = await runDocWorkflow(uploadedFiles, "dual", outputDir, log);
    } else if (workflowKey === "doc-standard") {
      outputs = await runDocWorkflow(uploadedFiles, "standard", outputDir, log);
    } else if (workflowKey === "ppt-dual") {
      outputs = await runPptWorkflow(uploadedFiles, "dual", outputDir, log);
    } else if (workflowKey === "ppt-standard") {
      outputs = await runPptWorkflow(uploadedFiles, "standard", outputDir, log);
    } else {
      throw new Error("Workflow is not implemented.");
    }

    if (outputs.length === 0) {
      throw new Error("No output file generated.");
    }

    const cleanup = async () => {
      await removePath(jobDir);
      for (const stagedPath of stagedPaths) {
        await removePath(stagedPath);
      }
    };
    res.setHeader("X-Workflow-Id", requestId);

    if (outputs.length === 1) {
      const outputFile = outputs[0];
      res.download(outputFile, path.basename(outputFile), async (error) => {
        if (error) {
          console.error(`[workflow:${requestId}] download error`, error);
        }
        await cleanup();
      });
      return;
    }

    const zipPath = path.join(outputDir, "processed-files.zip");
    await createZipArchive(zipPath, outputs, log);

    res.download(zipPath, "processed-files.zip", async (error) => {
      if (error) {
        console.error(`[workflow:${requestId}] zip download error`, error);
      }
      await cleanup();
    });
  } catch (error) {
    await removePath(jobDir);
    for (const stagedPath of stagedPaths) {
      await removePath(stagedPath);
    }
    const message = error && error.message ? error.message : String(error);
    const details = logs.slice(-25).join("\n");
    res.status(400).json({
      ok: false,
      error: message,
      details,
      workflow: workflow ? workflow.label : null,
    });
  }
});

app.get("*", (_req, res) => {
  res.sendFile(path.join(webRoot, "index.html"));
});

app.listen(port, () => {
  console.log(`Workflow web app running on http://localhost:${port}`);
  console.log(`Repo root: ${repoRoot}`);
});
