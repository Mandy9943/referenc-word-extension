/* global Office */

type TelemetryStatus = "start" | "success" | "error";

interface AddinTelemetryPayload {
  action: string;
  status: TelemetryStatus;
  mode?: string;
  durationMs?: number;
  warningCount?: number;
  errorMessage?: string;
  metadata?: Record<string, unknown>;
}

const DEFAULT_LOCAL_TELEMETRY_ORIGIN = "https://localhost:3000";
const TELEMETRY_PATH = "/telemetry/addin";
const SESSION_STORAGE_KEY = "addinTelemetrySessionId";

function getSessionId(): string {
  try {
    const storage = window?.sessionStorage;
    if (!storage) {
      return `anon-${Date.now()}`;
    }

    const existing = storage.getItem(SESSION_STORAGE_KEY);
    if (existing) {
      return existing;
    }

    const generated =
      typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
        ? crypto.randomUUID()
        : `session-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
    storage.setItem(SESSION_STORAGE_KEY, generated);
    return generated;
  } catch {
    return `anon-${Date.now()}`;
  }
}

function getHostLabel(): string {
  try {
    const host = Office?.context?.host;
    if (!host) {
      return "unknown";
    }
    return String(host).toLowerCase();
  } catch {
    return "unknown";
  }
}

function sanitizeText(value: unknown, maxLength: number = 1000): string {
  const text = String(value || "");
  return text.length > maxLength ? text.slice(0, maxLength) : text;
}

export function createTelemetryRequestId(action: string, mode: string): string {
  const host = getHostLabel();
  const safeAction = sanitizeText(action, 64).replace(/[^a-zA-Z0-9_-]+/g, "-");
  const safeMode = sanitizeText(mode || "unknown", 32).replace(/[^a-zA-Z0-9_-]+/g, "-");
  return `addon-${host}-${safeAction}-${safeMode}-${Date.now()}`;
}

export async function emitAddinTelemetry(payload: AddinTelemetryPayload): Promise<void> {
  try {
    const body = {
      source: "office-addin",
      host: getHostLabel(),
      sessionId: getSessionId(),
      timestamp: new Date().toISOString(),
      action: sanitizeText(payload.action, 120),
      status: payload.status,
      mode: sanitizeText(payload.mode || "unknown", 40),
      durationMs:
        typeof payload.durationMs === "number" && Number.isFinite(payload.durationMs)
          ? Math.max(0, Math.round(payload.durationMs))
          : null,
      warningCount:
        typeof payload.warningCount === "number" && Number.isFinite(payload.warningCount)
          ? Math.max(0, Math.round(payload.warningCount))
          : 0,
      errorMessage: payload.errorMessage ? sanitizeText(payload.errorMessage, 1200) : "",
      metadata: payload.metadata || {},
    };

    const origin = (() => {
      try {
        return window.location.origin;
      } catch {
        return "";
      }
    })();

    const endpoints: string[] = [];
    if (origin) {
      endpoints.push(`${origin}${TELEMETRY_PATH}`);
    } else {
      endpoints.push(TELEMETRY_PATH);
    }
    if (origin !== DEFAULT_LOCAL_TELEMETRY_ORIGIN) {
      endpoints.push(`${DEFAULT_LOCAL_TELEMETRY_ORIGIN}${TELEMETRY_PATH}`);
    }

    for (const endpoint of endpoints) {
      try {
        await fetch(endpoint, {
          method: "POST",
          mode: "cors",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(body),
          keepalive: true,
        });
        break;
      } catch {
        // Try next endpoint.
      }
    }
  } catch {
    // Telemetry must never block user actions.
  }
}
