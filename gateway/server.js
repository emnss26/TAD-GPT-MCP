import "dotenv/config";
import express from "express";
import fetch from "node-fetch";
import cors from "cors";

const app = express();
app.use(express.json({ limit: "2mb" }));
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

// --- Auth simple del gateway (permite /health y /openapi.json sin auth para debug) ---
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
  const u = new URL(BRIDGE_URL); // BRIDGE_URL ya apunta a .../mcp
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
  const r = await fetch(`${bridgeBase()}${path}`, {
    method: "GET",
    headers: {
      ...(MCP_API_KEY ? { Authorization: `Bearer ${MCP_API_KEY}` } : {})
    }
  });
  return r.json();
}

// Normalizador de acciones (anti-nombres inventados)
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
  return n; // si no encontramos match, pasamos tal cual
}

// --- Health ---
app.get("/health", (_req, res) => res.json({ ok: true, bridge: BRIDGE_URL }));

// --- Listado de acciones (proxy a bridge) ---
app.get("/actions", async (_req, res) => {
  const j = await callBridgeGet("/actions").catch(() => null);
  if (!j?.ok) return res.status(502).json({ ok:false, message:"Bridge /actions failed", detail:j });
  res.json(j);
});

// --- Proxy genérico ---
app.post("/mcp", async (req, res) => {
  const { action, args } = req.body || {};
  if (!action) return res.status(400).json({ ok: false, message: "Missing 'action'." });

  let known = [];
  try {
    const actionsResp = await callBridgeGet("/actions");
    if (actionsResp?.ok) known = actionsResp.data.actions || [];
  } catch (_) {}

  const normalized = normalizeActionName(action, known);
  res.json(await callBridge(normalized, args || {}));
});

// --- OpenAPI dinámico (opcional, útil para importar directo en Builder) ---
app.get("/openapi.json", async (req, res) => {
  let actions = [];
  try {
    const a = await callBridgeGet("/actions");
    if (a?.ok) actions = a.data.actions || [];
  } catch (_) {}
  const pubUrl = `${req.protocol}://${req.get("host")}`;

  const spec = {
    openapi: "3.1.0",
    info: { title: "TAD-GPT Gateway", version: "1.0.0" },
    servers: [{ url: pubUrl }],
    security: [{ bearerAuth: [] }],
    paths: {
      "/health": {
        get: {
          operationId: "health",
          summary: "Healthcheck",
          responses: {
            "200": {
              description: "OK",
              content: {
                "application/json": {
                  schema: {
                    type: "object",
                    properties: { ok: { type: "boolean" }, bridge: { type: "string" } }
                  }
                }
              }
            }
          }
        }
      },
      "/actions": {
        get: {
          operationId: "listActions",
          summary: "List available bridge actions",
          responses: {
            "200": {
              description: "OK",
              content: {
                "application/json": {
                  schema: {
                    type: "object",
                    properties: {
                      ok: { type: "boolean" },
                      data: { type: "object", properties: { actions: { type: "array", items: { type: "string" } } } }
                    }
                  }
                }
              }
            }
          }
        }
      },
      "/mcp": {
        post: {
          operationId: "mcpProxy",
          summary: "Generic MCP Proxy",
          requestBody: {
            required: true,
            content: {
              "application/json": {
                schema: { $ref: "#/components/schemas/McpEnvelope" },
                examples: {
                  levels: { value: { action: "levels.list", args: {} } },
                  qtoWalls: { value: { action: "qto.walls", args: {} } }
                }
              }
            }
          },
          responses: {
            "200": {
              description: "Bridge response",
              content: { "application/json": { schema: { $ref: "#/components/schemas/McpResponse" } } }
            }
          }
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
        McpResponse: {
          type: "object",
          properties: {
            ok: { type: "boolean" },
            message: { type: "string" },
            detail: { type: "object", additionalProperties: true },
            data: {
              oneOf: [
                { type: "object", additionalProperties: true },
                { type: "array", items: {} },
                { type: "string" },
                { type: "number" },
                { type: "integer" },
                { type: "boolean" },
                { type: "null" }
              ]
            }
          }
        }
      }
    }
  };

  res.json(spec);
});

// --- Start ---
app.listen(PORT, () => console.log(`Gateway listening on http://localhost:${PORT}`));
