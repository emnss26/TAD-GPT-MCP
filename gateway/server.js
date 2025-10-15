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

// --- Auth simple del gateway ---
app.use((req, res, next) => {
  // Permite /health sin auth
  if (req.path === "/health") return next();

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
    if (!r.ok) {
      return { ok: false, message: `Bridge HTTP ${r.status}`, detail: json || null };
    }
    return json;
  } catch (err) {
    const msg = err?.name === "AbortError" ? "Bridge timeout" : (err?.message || "Bridge error");
    return { ok: false, message: msg };
  } finally {
    cancel();
  }
}

// --- Healthcheck ---
app.get("/health", (_req, res) => {
  res.json({ ok: true, bridge: BRIDGE_URL });
});

// ====== ENDPOINTS CÓMODOS ======
// QTO — Pipes
app.post("/qto/mep/pipes", async (req, res) => {
  res.json(await callBridge("qto.mep.pipes", req.body || {}));
});

// QTO — Ducts
app.post("/qto/mep/ducts", async (req, res) => {
  res.json(await callBridge("qto.mep.ducts", req.body || {}));
});

// Query — Levels
app.post("/query/levels", async (_req, res) => {
  res.json(await callBridge("levels.list", {}));
});

// QA — Apply view templates
app.post("/qa/apply-view-templates", async (req, res) => {
  res.json(await callBridge("qa.fix.apply_view_templates", req.body || {}));
});

// Export — NWC
app.post("/export/nwc", async (req, res) => {
  res.json(await callBridge("export.nwc", req.body || {}));
});

// Structure — Wall foundation
app.post("/struct/foundation/wall", async (req, res) => {
  res.json(await callBridge("struct.foundation.wall.create", req.body || {}));
});

// ====== PROXY GENÉRICO ======
// Body: { action: "qto.walls", args: { ... } }
app.post("/mcp", async (req, res) => {
  const { action, args } = req.body || {};
  if (!action) return res.status(400).json({ ok: false, message: "Missing 'action'." });
  res.json(await callBridge(action, args || {}));
});

// --- Start ---
app.listen(PORT, () => {
  console.log(`Gateway listening on http://localhost:${PORT}`);
});
