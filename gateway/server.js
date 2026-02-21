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
  "qto.floor": "qto.floors",
  "qto.floor.types": "qto.floors",
  "qto.wall": "qto.walls",
  "view.category.override": "view.category.override_color",
  "view.category.color": "view.category.override_color",
  "view.category.set_color": "view.category.override_color",
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
  // Query / catalog
  "levels.list": {
    args: {},
    examples: [{ action: "levels.list", args: {} }]
  },
  "view.active": {
    args: {},
    examples: [{ action: "view.active", args: {} }]
  },
  "views.list": {
    args: {},
    examples: [{ action: "views.list", args: {} }]
  },
  "categories.list": {
    args: {},
    examples: [{ action: "categories.list", args: {} }]
  },
  "materials.list": {
    args: {},
    examples: [{ action: "materials.list", args: {} }]
  },
  "walltypes.list": {
    args: {},
    examples: [{ action: "walltypes.list", args: {} }]
  },
  "families.types.list": {
    args: {},
    examples: [{ action: "families.types.list", args: {} }]
  },
  "selection.info": {
    args: { includeParameters: "boolean?", topNParams: "int?" },
    examples: [
      { action: "selection.info", args: { includeParameters: false } },
      { action: "selection.info", args: { includeParameters: true, topNParams: 30 } }
    ]
  },
  "element.info": {
    args: { elementId: "int (required)", includeParameters: "boolean?", topNParams: "int?" },
    examples: [
      { action: "element.info", args: { elementId: 12345 } },
      { action: "element.info", args: { elementId: 12345, includeParameters: true, topNParams: 40 } }
    ]
  },

  // Create / edit
  "level.create": {
    args: { name: "string?", "elevation_m|elevation_ft": "number (required one)" },
    examples: [
      { action: "level.create", args: { name: "Nivel 2", elevation_m: 3.5 } },
      { action: "level.create", args: { name: "Roof", elevation_ft: 49.21 } }
    ]
  },
  "floor.create": {
    args: {
      level: "string?",
      floorType: "string?",
      profile: "Pt2[] (required, min 3) -> [{x:number,y:number}, ...]"
    },
    examples: [
      {
        action: "floor.create",
        args: {
          level: "ARQ-P.B.",
          floorType: "Generic 150mm",
          profile: [{ x: 0, y: 0 }, { x: 8, y: 0 }, { x: 8, y: 6 }, { x: 0, y: 6 }]
        }
      }
    ]
  },

  // QTO
  "qto.walls": {
    args: {
      groupBy: "string[]? -> any of: type, level, phase",
      includeIds: "boolean?",
      filter: "{ typeIds?:int[], typeNames?:string[], levels?:string[], phase?:string, useSelection?:boolean }?"
    },
    examples: [
      { action: "qto.walls", args: {} },
      { action: "qto.walls", args: { groupBy: ["type"] } },
      { action: "qto.walls", args: { groupBy: ["level", "type"], includeIds: true } },
      { action: "qto.walls", args: { filter: { levels: ["ARQ-P.B."], useSelection: false } } }
    ]
  },
  "qto.floors": {
    args: {
      groupBy: "string[]? -> any of: type, level",
      includeIds: "boolean?"
    },
    examples: [
      { action: "qto.floors", args: {} },
      { action: "qto.floors", args: { groupBy: ["type"] } },
      { action: "qto.floors", args: { groupBy: ["level", "type"], includeIds: true } }
    ],
    notes: [
      "This endpoint supports grouping.",
      "Use groupBy=['type'] to get totals per floor type."
    ]
  },
  "qto.ceilings": {
    args: {
      groupBy: "string[]? -> any of: type, level",
      includeIds: "boolean?"
    },
    examples: [
      { action: "qto.ceilings", args: {} },
      { action: "qto.ceilings", args: { groupBy: ["type"] } }
    ]
  },
  "qto.railings": {
    args: {
      groupBy: "string[]? -> any of: type, level",
      includeIds: "boolean?"
    },
    examples: [
      { action: "qto.railings", args: {} },
      { action: "qto.railings", args: { groupBy: ["level"] } }
    ]
  },
  "qto.families.count": {
    args: {
      groupBy: "string[]? -> any of: category, family, type, level",
      includeIds: "boolean?",
      categories: "string[]? (ex: OST_Doors, Doors)"
    },
    examples: [
      { action: "qto.families.count", args: { groupBy: ["category", "family", "type"] } },
      { action: "qto.families.count", args: { groupBy: ["type"], categories: ["OST_Doors"] } }
    ]
  },

  // Graphics
  "view.category.set_visibility": {
    args: {
      categories: "string[] (required)",
      visible: "boolean? (default true)",
      forceDetachTemplate: "boolean? (default false)",
      viewId: "int?"
    },
    examples: [
      { action: "view.category.set_visibility", args: { categories: ["Floors"], visible: false } },
      { action: "view.category.set_visibility", args: { categories: ["OST_Floors"], visible: true, forceDetachTemplate: true } }
    ]
  },
  "view.category.clear_overrides": {
    args: {
      categories: "string[] (required)",
      forceDetachTemplate: "boolean? (default false)",
      viewId: "int?"
    },
    examples: [
      { action: "view.category.clear_overrides", args: { categories: ["Floors"] } }
    ]
  },
  "view.category.override_color": {
    args: {
      categories: "string[] (required)",
      color: "string hex #RRGGBB OR {r:int,g:int,b:int} (required)",
      transparency: "int 0..100?",
      halftone: "boolean?",
      surfaceSolid: "boolean? (default true)",
      projectionLines: "boolean? (default false)",
      forceDetachTemplate: "boolean? (default false)",
      viewId: "int?"
    },
    examples: [
      {
        action: "view.category.override_color",
        args: {
          categories: ["Floors"],
          color: "#FF0000",
          projectionLines: true,
          surfaceSolid: true
        }
      },
      {
        action: "view.category.override_color",
        args: {
          categories: ["OST_Floors"],
          color: { r: 255, g: 0, b: 0 },
          transparency: 0,
          forceDetachTemplate: true
        }
      }
    ]
  },
  "view.set_scale": {
    args: { viewId: "int?", scale: "int >= 1 (required)" },
    examples: [{ action: "view.set_scale", args: { scale: 100 } }]
  },
  "view.set_detail_level": {
    args: { viewId: "int?", detailLevel: "coarse|medium|fine (required)" },
    examples: [{ action: "view.set_detail_level", args: { detailLevel: "fine" } }]
  },
  "view.set_discipline": {
    args: { viewId: "int?", discipline: "architectural|structural|mechanical|coordination (required)" },
    examples: [{ action: "view.set_discipline", args: { discipline: "coordination" } }]
  },
  "view.set_phase": {
    args: { viewId: "int?", phase: "string?" },
    examples: [{ action: "view.set_phase", args: { phase: "New Construction" } }]
  },

  // Parameters
  "params.get": {
    args: {
      elementIds: "int[] (required)",
      paramNames: "string[] (required)",
      includeValueString: "boolean?"
    },
    examples: [
      { action: "params.get", args: { elementIds: [344900], paramNames: ["Comments"], includeValueString: true } }
    ]
  },
  "params.set": {
    args: {
      updates: "array (required) -> [{ elementId:int, param|bip|guid:string, value:any }, ...]"
    },
    examples: [
      {
        action: "params.set",
        args: { updates: [{ elementId: 344900, param: "Comments", value: "Updated from MCP" }] }
      }
    ]
  },
  "params.set_where": {
    args: {
      where: "object (required)",
      set: "array (required) -> [{ param|bip|guid:string, value:any }, ...]",
      allowTypeParams: "boolean?"
    },
    examples: [
      {
        action: "params.set_where",
        args: {
          where: { categories: ["OST_Walls"] },
          set: [{ param: "Comments", value: "QA Checked" }],
          allowTypeParams: false
        }
      }
    ]
  }
};

function inferHintForAction(action) {
  const a = String(action || "").trim();
  if (!a) return { args: {}, examples: [{ action: "", args: {} }] };

  if (a.endsWith(".list") || a === "view.active") {
    return { args: {}, examples: [{ action: a, args: {} }] };
  }

  if (a.startsWith("qto.")) {
    const out = {
      args: {
        groupBy: "string[]?",
        includeIds: "boolean?"
      },
      examples: [
        { action: a, args: {} },
        { action: a, args: { groupBy: ["type"] } }
      ]
    };
    if (a.endsWith(".count")) {
      out.notes = ["Count endpoint: usually returns grouped counts, not quantities."];
    }
    return out;
  }

  if (a.startsWith("view.category.")) {
    return {
      args: {
        categories: "string[]",
        viewId: "int?",
        forceDetachTemplate: "boolean?"
      },
      examples: [{ action: a, args: { categories: ["Floors"] } }]
    };
  }

  if (a.startsWith("view.") || a.startsWith("views.")) {
    return {
      args: { viewId: "int?" },
      examples: [{ action: a, args: {} }]
    };
  }

  if (a.startsWith("params.")) {
    return {
      args: { "varies": "See action-specific hints. Use /mcp.flat for tolerant payloads." },
      examples: [{ action: a, args: {} }]
    };
  }

  if (a.includes(".create") || a.includes(".place") || a.includes(".set")) {
    return {
      args: { "varies": "Action-specific object payload" },
      examples: [{ action: a, args: {} }]
    };
  }

  return { args: {}, examples: [{ action: a, args: {} }] };
}

function buildHintsForActions(actions = []) {
  const out = {};
  for (const action of actions || []) {
    out[action] = ACTION_HINTS[action] || inferHintForAction(action);
  }
  return out;
}

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

function maybeCoerceBool(key, val) {
  if (!/(include|force|visible|halftone|structural|combine|exportLinks|convertElementProperties|useSelection|projectionLines|surfaceSolid|placeOnlyEnclosed|roundVsRect|dryRun|allowTypeParams|flipHand|flipFacing)$/i.test(key)) {
    return val;
  }
  if (typeof val === "boolean") return val;
  if (typeof val === "number") return val !== 0;
  if (typeof val === "string") {
    const low = val.trim().toLowerCase();
    if (low === "true" || low === "1" || low === "yes") return true;
    if (low === "false" || low === "0" || low === "no") return false;
  }
  return val;
}

function toArray(val) {
  if (Array.isArray(val)) return val;
  if (val == null) return [];
  if (typeof val === "string") {
    const s = val.trim();
    if (!s) return [];
    if (s.includes(",")) return s.split(",").map(x => x.trim()).filter(Boolean);
    return [s];
  }
  return [val];
}

function normalizeArgsAliases(action, args) {
  const src = (args && typeof args === "object") ? args : {};
  const out = {};

  for (const [rawK, rawV] of Object.entries(src)) {
    let k = toCamelCaseKey(rawK);

    if (k === "groupby") k = "groupBy";
    if (k === "includeRows" || k === "rows" || k === "includeDetails") k = "includeIds";
    if (k === "category") k = "categories";
    if (k === "view") k = "viewId";
    if (k === "id" && (action === "element.info" || action === "view.apply_template")) {
      k = action === "element.info" ? "elementId" : "templateId";
    }

    // Combinar categorías si vienen múltiples llaves equivalentes
    if (k === "categories" && out.categories != null) {
      out.categories = [...toArray(out.categories), ...toArray(rawV)];
      continue;
    }
    out[k] = rawV;
  }

  // Atajos para color en override
  if (action === "view.category.override_color") {
    if (out.color == null) {
      if (Array.isArray(out.rgb) && out.rgb.length >= 3) {
        out.color = { r: out.rgb[0], g: out.rgb[1], b: out.rgb[2] };
      } else if (out.rgb && typeof out.rgb === "object") {
        out.color = out.rgb;
      } else if (out.r != null && out.g != null && out.b != null) {
        out.color = { r: out.r, g: out.g, b: out.b };
      } else if (out.red != null && out.green != null && out.blue != null) {
        out.color = { r: out.red, g: out.green, b: out.blue };
      } else if (typeof out.hex === "string") {
        out.color = out.hex;
      }
    }
    if (out.categories != null) out.categories = toArray(out.categories);
  }

  if (action.startsWith("qto.") && out.groupBy != null) {
    out.groupBy = toArray(out.groupBy);
  }

  if (out.categories != null && !Array.isArray(out.categories)) {
    out.categories = toArray(out.categories);
  }

  return out;
}

function coerceArgsForAction(action, rawArgs) {
  if (!rawArgs || typeof rawArgs !== "object") return rawArgs || {};
  const normalizedIn = normalizeArgsAliases(action, rawArgs);
  const out = {};
  for (const [k, v] of Object.entries(normalizedIn)) {
    const camel = toCamelCaseKey(k);
    const scalar = maybeCoerceScalar(camel, v);
    const coerced = maybeCoerceBool(camel, scalar);

    if (/(^|.*)(viewIds|levelIds|typeIds|templateIds|elementIds|categoryIds)$/i.test(camel)) {
      const arr = toArray(coerced)
        .map(x => typeof x === "string" ? parseInt(x, 10) : x)
        .filter(x => !(typeof x === "number" && Number.isNaN(x)));
      out[camel] = arr;
      continue;
    }

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
  const actions = j.data.actions || [];
  const hints = buildHintsForActions(actions);
  res.json({
    ok: true,
    data: {
      actions,
      hints,
      hintsMeta: {
        hintedActions: Object.keys(hints).length,
        generatedAt: new Date().toISOString()
      }
    }
  });
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
    info: { title: "TAD-GPT Gateway", version: "1.3.0" },
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
          summary: "List available bridge actions and runtime hints",
          description: "Call this first. Use `data.hints[action]` to build valid args payloads for `/mcp`.",
          responses: {
            "200": {
              description: "OK",
              content: {
                "application/json": {
                  schema: {
                    type: "object",
                    properties: {
                      ok: { type: "boolean" },
                      data: {
                        type: "object",
                        properties: {
                          actions: { type: "array", items: { type: "string" } },
                          hints: { type: "object", additionalProperties: true },
                          hintsMeta: {
                            type: "object",
                            properties: {
                              hintedActions: { type: "integer" },
                              generatedAt: { type: "string" }
                            }
                          }
                        }
                      }
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
          summary: "Generic MCP Proxy (tolerant JSON)",
          description: "Tip: call /actions first and then send args matching `data.hints[action].args`.",
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
