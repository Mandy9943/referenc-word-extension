/* eslint-disable no-undef */

const fs = require("fs");
const os = require("os");
const path = require("path");
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");
const Dotenv = require("dotenv-webpack");

const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION
const addinLogPath = path.join(os.homedir(), "Desktop", "add-in-logs.md");
const addinLogHeader = `# Add-in Pipeline Logs

Auto-generated run telemetry for DOCX/PPTX automation.
This file logs success/failure runs, timings, request behavior, and improvement hints.

`;

function ensureAddinLogHeader() {
  if (fs.existsSync(addinLogPath)) {
    return;
  }
  fs.mkdirSync(path.dirname(addinLogPath), { recursive: true });
  fs.writeFileSync(addinLogPath, addinLogHeader, "utf8");
}

function safeString(value, maxLength = 2000) {
  const text = String(value ?? "");
  return text.length > maxLength ? text.slice(0, maxLength) : text;
}

function parseMaybeNumber(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
}

function appendManualAddinTelemetry(payload) {
  ensureAddinLogHeader();

  const timestamp = safeString(payload.timestamp || new Date().toISOString(), 64);
  const action = safeString(payload.action || "unknown_action", 120);
  const status = safeString(payload.status || "unknown", 32).toUpperCase();
  const mode = safeString(payload.mode || "unknown", 40);
  const source = safeString(payload.source || "office-addin", 40);
  const host = safeString(payload.host || "unknown", 40);
  const sessionId = safeString(payload.sessionId || "unknown", 120);
  const durationMs = parseMaybeNumber(payload.durationMs);
  const warningCount = parseMaybeNumber(payload.warningCount) ?? 0;
  const errorMessage = safeString(payload.errorMessage || "", 1600);
  const metadata = payload && typeof payload.metadata === "object" ? payload.metadata : {};

  const lines = [];
  lines.push(`## ${timestamp} | ADDIN | ${status}`);
  lines.push("");
  lines.push(`- \`source\`: \`${source}\``);
  lines.push(`- \`host\`: \`${host}\``);
  lines.push(`- \`session_id\`: \`${sessionId}\``);
  lines.push(`- \`action\`: \`${action}\``);
  lines.push(`- \`mode\`: \`${mode}\``);
  if (durationMs !== null) {
    lines.push(`- \`duration_s\`: \`${(durationMs / 1000).toFixed(2)}\``);
  }
  lines.push(`- \`warning_count\`: \`${warningCount}\``);
  if (errorMessage) {
    lines.push(`- \`error\`: \`${errorMessage}\``);
  }
  lines.push("");
  lines.push("### Raw Event");
  lines.push("```json");
  lines.push(JSON.stringify(payload, null, 2));
  lines.push("```");
  lines.push("");
  lines.push("---");
  lines.push("");

  fs.appendFileSync(addinLogPath, lines.join("\n"), "utf8");
}

function parseJsonBody(req) {
  if (req.body && typeof req.body === "object") {
    return Promise.resolve(req.body);
  }

  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", (chunk) => {
      body += String(chunk);
      if (body.length > 2 * 1024 * 1024) {
        reject(new Error("Telemetry payload too large"));
      }
    });
    req.on("end", () => {
      if (!body.trim()) {
        resolve({});
        return;
      }
      try {
        resolve(JSON.parse(body));
      } catch (error) {
        reject(error);
      }
    });
    req.on("error", reject);
  });
}

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      vendor: ["react", "react-dom", "core-js", "@fluentui/react-components", "@fluentui/react-icons"],
      taskpane: ["./src/taskpane/index.tsx", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.ts",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: ["ts-loader"],
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|ttf|woff|woff2|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new Dotenv({
        path: `./.env`,
        systemvars: true,
        safe: true,
        defaults: true,
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "vendor", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
    ],
    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      allowedHosts: "all",
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      setupMiddlewares: (middlewares, devServer) => {
        if (devServer && devServer.app) {
          devServer.app.post("/telemetry/addin", async (req, res) => {
            try {
              const payload = await parseJsonBody(req);
              appendManualAddinTelemetry(payload);
              res.status(200).json({ status: "ok" });
            } catch (error) {
              console.error("Failed to write add-in telemetry:", error);
              res.status(400).json({
                status: "error",
                message: error instanceof Error ? error.message : String(error),
              });
            }
          });
        }
        return middlewares;
      },
    },
  };

  return config;
};
