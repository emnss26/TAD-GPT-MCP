import "dotenv/config";
import express from "express";
import fetch from "node-fetch";
import cors from "cors";

const app = express();

// Soportar JSON y también form-urlencoded
app.use(express.json({ limit: "2mb" }));
app.use(express.urlencoded({ extended: true }));
app.use(cors());

const {
  PORT = 8080,
  BRIDGE_URL,
  MCP_API_KEY,
  PUBLIC_BEARER,
  FETCH_TIMEOUT_MS = 30000,
} = process.env;

if (!BRIDGE_URL) {
  console.error("❌ Missing BRIDGE_URL in .env");
  process.exit(1);
}

// --- Auth simple del gateway (permitir /health y /openapi.json sin auth) ---
app.use((req, res, next) => {
  if (req.path === "/health" || req.path === "/openapi.json") return next();
  const auth = req.headers.authorization || "";
  if (PUBLIC_BEARER && auth !== `Bearer ${PUBLIC_BEARER}`) {
    return res.status(401).json({ ok: false, message: "Unauthorized (gateway)" });
  }
  next();
});

// --- Helpers ---
function withTimeout(ms) {
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), Number(ms));
  return { signal: controller.signal, cancel: () => clearTimeout(t) };
}

function bridgeBase() {
  const u = new URL(BRIDGE_URL); // BRIDGE_URL apunta a .../mcp
  return `${u.protocol}//${u.host}`;
}

async function callBridge(action, args = {}) {
  const { signal, cancel } = withTimeout(FETCH_TIMEOUT_MS);
  try {
    const r = await fetch(BRIDGE_URL, {
      method: "POST",
      headers: {
        "content-type": "application/json",
        ...(MCP_API_KEY ? { Authorization: `Bearer ${MCP_API_KEY}` } : {})
      },
      body: JSON.stringify({ action, args }),
      signal
    });
    const json = await r.json().catch(() => ({}));
    if (!r.ok) return { ok: false, message: `Bridge HTTP ${r.status}`, detail: json || null };
    return json;
  } catch (err) {
    const msg = err?.name === "AbortError" ? "Bridge timeout" : (err?.message || "Bridge error");
    return { ok: false, message: msg };
  } finally { cancel(); }
}

async function callBridgeGet(path) {
  try {
    const r = await fetch(`${bridgeBase()}${path}`, {
      method: "GET",
      headers: { ...(MCP_API_KEY ? { Authorization: `Bearer ${MCP_API_KEY}` } : {}) }
    });
    return await r.json();
  } catch (e) {
    return null;
  }
}

// --- Normalizador de acciones ---
const manualAliases = {
  "model.levels.list": "levels.list",
  "qto.walls.types": "qto.walls",
};
function normalizeActionName(name, known = []) {
  if (!name) return name;
  const n = String(name).trim();
  if (manualAliases[n]) return manualAliases[n];

  const candidates = new Set([
    n,
    n.toLowerCase(),
    n.replaceAll('_', '.'),
    n.replaceAll(' ', ''),
    n.replaceAll(' ', '_').replaceAll('__', '_'),
  ]);
  for (const c of candidates) if (known.includes(c)) return c;
  return n; // si no encontramos match, lo pasamos tal cual
}

// --- Hints opcionales para ayudar al LLM ---
const ACTION_HINTS = {
  "level.create": {
    args: { name: "string", "elevation_m|elevation_ft": "number" },
    examples: [
      { action: "level.create", args: { name: "Nivel 2", elevation_m: 3.5 } },
      { action: "level.create", args: { name: "Roof", elevation_ft: 49.21 } }
    ]
  }
};

// --- Coerción de argumentos ---
function toCamelCaseKey(k) {
  if (!k || typeof k !== "string") return k;
  if (!k.includes("_")) return k;
  return k.split("_").map((s, i) => i === 0 ? s : (s.charAt(0).toUpperCase() + s.slice(1))).join("");
}

function parseNumericWithUnits(raw) {
  if (typeof raw === "number") return { num: raw, unit: null };
  if (typeof raw !== "string") return { num: raw, unit: null };
  const s = raw.trim().toLowerCase();
  const m = s.match(/^(-?\d+(\.\d+)?)(\s*(mm|cm|m|ft|feet|in|inch|inches|"))?$/i);
  if (!m) return { num: isNaN(Number(s)) ? raw : Number(s), unit: null };
  const value = Number(m[1]);
  const unit = (m[4] || (s.endsWith('"') ? 'in' : null))?.toLowerCase() || null;
  return { num: value, unit };
}

function toMeters(n, unit) {
  if (typeof n !== "number" || isNaN(n)) return n;
  switch (unit) {
    case "mm": return n / 1000.0;
    case "cm": return n / 100.0;
    case "m":  return n;
    case "ft":
    case "feet": return n * 0.3048;
    case "in":
    case "inch":
    case "inches":
    case '"': return n * 0.0254;
    default: return n; // si no hay unidad, asumimos metros
  }
}

function maybeCoerceScalar(key, val) {
  if (typeof val === "string") {
    const low = val.trim().toLowerCase();
    if (low === "true") return true;
    if (low === "false") return false;
    if (low.includes(",") && !/^{|\[/.test(low)) {
      const parts = val.split(",").map(s => s.trim()).filter(Boolean);
      if (parts.length > 1) return parts;
    }
  }
  if (typeof val === "string") {
    const { num, unit } = parseNumericWithUnits(val);
    if (typeof num === "number" && !isNaN(num)) {
      if (/(elevation|offset|diameter|length|height|thickness|size)/i.test(key) || unit) {
        return toMeters(num, unit);
      }
      return num;
    }
  }
  if (typeof val === "number") {
    if (/_mm$/i.test(key)) return val / 1000.0;
    if (/_cm$/i.test(key)) return val / 100.0;
    if (/_ft$/i.test(key)) return val * 0.3048;
    if (/_in$/i.test(key)) return val * 0.0254;
    if (/_m$/i.test(key))  return val;
  }
  return val;
}

function coerceArgsForAction(action, rawArgs) {
  if (!rawArgs || typeof rawArgs !== "object") return rawArgs || {};
  const out = {};
  for (const [k, v] of Object.entries(rawArgs)) {
    const camel = toCamelCaseKey(k);
    const coerced = maybeCoerceScalar(camel, v);
    if (/(^|.*)(viewId|levelId|typeId|templateId|elementId|categoryId)$/i.test(camel)) {
      out[camel] = typeof coerced === "string" ? parseInt(coerced, 10) : coerced;
      continue;
    }
    out[camel] = coerced;
  }
  if (out.elevation_m != null && out.elevation == null) out.elevation = out.elevation_m;
  if (out.elevation_ft != null && out.elevation == null) out.elevation = toMeters(Number(out.elevation_ft), "ft");
  return out;
}

function mergeArgsTolerant(bodyOrQuery) {
  const src = bodyOrQuery || {};
  const { action, args, ...rest } = src;
  let merged = {};
  if (args && typeof args === "string") {
    try { merged = JSON.parse(args); } catch { merged = {}; }
  } else if (args && typeof args === "object") {
    merged = { ...args };
  }
  // el resto plano también cuenta como args
  return { action, args: { ...rest, ...merged } };
}

// --- Health
app.get("/health", (_req, res) => res.json({ ok: true, bridge: BRIDGE_URL }));

// --- Actions + hints
app.get("/actions", async (_req, res) => {
  const j = await callBridgeGet("/actions").catch(() => null);
  if (!j?.ok) return res.status(502).json({ ok:false, message:"Bridge /actions failed", detail:j });
  res.json({ ok: true, data: { actions: j.data.actions || [], hints: ACTION_HINTS } });
});

// --- /mcp (JSON tolerante)
app.post("/mcp", async (req, res) => {
  // acepta { action, args } o { action, ...flat }
  const { action, args } = mergeArgsTolerant(req.body);
  if (!action) return res.status(400).json({ ok: false, message: "Missing 'action'." });

  let known = [];
  try {
    const actionsResp = await callBridgeGet("/actions");
    if (actionsResp?.ok) known = actionsResp.data.actions || actionsResp.data || [];
  } catch (_) {}

  const normalized = normalizeActionName(action, known);
  const coerced = coerceArgsForAction(normalized, args || {});
  res.json(await callBridge(normalized, coerced));
});

// --- /mcp.flat (body plano)
app.post("/mcp.flat", async (req, res) => {
  const { action, args } = mergeArgsTolerant(req.body);
  if (!action) return res.status(400).json({ ok: false, message: "Missing 'action'." });

  let known = [];
  try {
    const actionsResp = await callBridgeGet("/actions");
    if (actionsResp?.ok) known = actionsResp.data.actions || actionsResp.data || [];
  } catch (_) {}

  const normalized = normalizeActionName(action, known);
  const coerced = coerceArgsForAction(normalized, args || {});
  res.json(await callBridge(normalized, coerced));
});

// --- /exec GET (query) y POST (form/json) — evita UnrecognizedKwargsError de builders
app.get("/exec", async (req, res) => {
  const { action, args } = mergeArgsTolerant(req.query);
  if (!action) return res.status(400).json({ ok: false, message: "Missing 'action'." });

  let known = [];
  try {
    const actionsResp = await callBridgeGet("/actions");
    if (actionsResp?.ok) known = actionsResp.data.actions || actionsResp.data || [];
  } catch (_) {}

  const normalized = normalizeActionName(action, known);
  const coerced = coerceArgsForAction(normalized, args || {});
  res.json(await callBridge(normalized, coerced));
});

app.post("/exec", async (req, res) => {
  // soporta x-www-form-urlencoded, JSON y mezcla
  const mix = { ...req.body, ...req.query };
  const { action, args } = mergeArgsTolerant(mix);
  if (!action) return res.status(400).json({ ok: false, message: "Missing 'action'." });

  let known = [];
  try {
    const actionsResp = await callBridgeGet("/actions");
    if (actionsResp?.ok) known = actionsResp.data.actions || actionsResp.data || [];
  } catch (_) {}

  const normalized = normalizeActionName(action, known);
  const coerced = coerceArgsForAction(normalized, args || {});
  res.json(await callBridge(normalized, coerced));
});

// --- OpenAPI dinámico
app.get("/openapi.json", async (req, res) => {
  let actions = [];
  try {
    const a = await callBridgeGet("/actions");
    if (a?.ok) actions = a.data.actions || [];
  } catch (_) {}

  const pubUrl = `${req.protocol}://${req.get("host")}`;
  const spec = {
    openapi: "3.1.0",
    info: { title: "TAD-GPT Gateway", version: "1.2.0" },
    servers: [{ url: pubUrl }],
    security: [{ bearerAuth: [] }],
    paths: {
      "/health": {
        get: {
          operationId: "health",
          summary: "Healthcheck",
          responses: { "200": { description: "OK" } }
        }
      },
      "/actions": {
        get: {
          operationId: "listActions",
          summary: "List available bridge actions",
          responses: { "200": { description: "OK" } }
        }
      },
      "/mcp": {
        post: {
          operationId: "mcpProxy",
          summary: "Generic MCP Proxy (tolerant JSON)",
          requestBody: {
            required: true,
            content: {
              "application/json": {
                schema: { $ref: "#/components/schemas/McpEnvelope" }
              }
            }
          },
          responses: { "200": { description: "Bridge response" } }
        }
      },
      "/mcp.flat": {
        post: {
          operationId: "mcpProxyFlat",
          summary: "Generic MCP Proxy (flat JSON)",
          requestBody: {
            required: true,
            content: {
              "application/json": { schema: { $ref: "#/components/schemas/McpFlatEnvelope" } }
            }
          },
          responses: { "200": { description: "Bridge response" } }
        }
      },
      "/exec": {
        get: {
          operationId: "execGet",
          summary: "Execute action via query",
          parameters: [
            { in: "query", name: "action", required: true, schema: { type: "string" } },
            { in: "query", name: "args", required: false, schema: { type: "string" } },
            { in: "query", name: "name", required: false, schema: { type: "string" } },
            { in: "query", name: "elevation", required: false, schema: { type: "number" } },
            { in: "query", name: "elevation_m", required: false, schema: { type: "number" } },
            { in: "query", name: "elevation_ft", required: false, schema: { type: "number" } }
          ],
          responses: { "200": { description: "Bridge response" } }
        },
        post: {
          operationId: "execPost",
          summary: "Execute action via form/json",
          requestBody: {
            required: true,
            content: {
              "application/x-www-form-urlencoded": {
                schema: {
                  type: "object",
                  properties: {
                    action: { type: "string" },
                    args:   { type: "string" },
                    name:   { type: "string" },
                    elevation: { type: "number" },
                    elevation_m: { type: "number" },
                    elevation_ft: { type: "number" }
                  }
                }
              },
              "application/json": {
                schema: { type: "object", additionalProperties: true }
              }
            }
          },
          responses: { "200": { description: "Bridge response" } }
        }
      }
    },
    components: {
      securitySchemes: { bearerAuth: { type: "http", scheme: "bearer" } },
      schemas: {
        McpEnvelope: {
          type: "object",
          required: ["action"],
          properties: {
            action: { type: "string", enum: actions.length ? actions : undefined },
            args: { type: "object", additionalProperties: true }
          }
        },
        McpFlatEnvelope: {
          type: "object",
          required: ["action"],
          additionalProperties: true,
          properties: { action: { type: "string", enum: actions.length ? actions : undefined } }
        }
      }
    }
  };
  res.json(spec);
});

// --- Start
app.listen(PORT, () => console.log(`Gateway listening on http://localhost:${PORT}`));
