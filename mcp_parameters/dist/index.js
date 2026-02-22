import "dotenv/config";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { postRevit } from "./bridge.js";
const server = new McpServer({
    name: "mcp-params",
    version: "1.0.0",
});
const asText = (obj) => ({
    content: [{ type: "text", text: JSON.stringify(obj, null, 2) }],
});
// ------------------------ Helpers ------------------------
function parseMaybeJson(v) {
    if (typeof v !== "string")
        return v;
    try {
        return JSON.parse(v);
    }
    catch {
        return v;
    }
}
function toBoolNum(v) {
    if (typeof v !== "string")
        return v;
    const s = v.trim().toLowerCase();
    if (s === "true")
        return true;
    if (s === "false")
        return false;
    if (!isNaN(Number(v)) && v !== "")
        return Number(v);
    return v;
}
function looksLikeGuid(s) {
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(s);
}
function looksLikeBip(s) {
    return /^[A-Z0-9_]+$/.test(s);
}
function pickToken(obj) {
    const cand = obj?.guid ??
        obj?.bip ??
        obj?.param ??
        obj?.parameter ??
        obj?.paramName ??
        obj?.parameterName;
    if (!cand || typeof cand !== "string")
        return null;
    if (looksLikeGuid(cand))
        return { kind: "guid", value: cand };
    if (looksLikeBip(cand))
        return { kind: "bip", value: cand };
    return { kind: "param", value: cand };
}
function expandElementIds(one, many, useSelection) {
    const ids = [];
    if (Array.isArray(many) && many.length)
        ids.push(...many);
    if (typeof one === "number")
        ids.push(one);
    if (useSelection || one === 0)
        ids.push(0); // 0 = selección actual (ParamsSet lo respalda)
    return ids;
}
// ------------------------ Normalizadores ------------------------
function normalizeParamsGetArgs(raw) {
    const args = raw ?? {};
    const elementIds = Array.isArray(args.elementIds)
        ? args.elementIds
        : typeof args.elementId === "number"
            ? [args.elementId]
            : [];
    let paramNames = parseMaybeJson(args.paramNames);
    if (!Array.isArray(paramNames)) {
        const single = args.param ?? args.parameter ?? null;
        paramNames = single ? [String(single)] : [];
    }
    const includeValueString = args.includeValueString === undefined ? true : !!args.includeValueString;
    return { elementIds, paramNames, includeValueString };
}
function normalizeParamsSetArgs(raw) {
    const args = raw ?? {};
    let updates = parseMaybeJson(args.updates);
    // Construir updates si vienen como single-call
    if (!Array.isArray(updates)) {
        const token = pickToken(args);
        const ids = expandElementIds(args.elementId, args.elementIds, args.useSelection);
        if (token && ids.length && "value" in args) {
            updates = ids.map((id) => {
                const entry = {
                    elementId: id,
                    value: toBoolNum(args.value),
                };
                if (token.kind === "guid")
                    entry.guid = token.value;
                else if (token.kind === "bip")
                    entry.bip = token.value;
                else
                    entry.param = token.value;
                return entry;
            });
        }
    }
    if (!Array.isArray(updates))
        updates = [];
    const norm = updates.flatMap((u) => {
        if (!u || typeof u !== "object")
            return [];
        const ids = expandElementIds(u.elementId, u.elementIds, u.useSelection);
        const token = pickToken(u);
        const val = toBoolNum(u.value);
        if (!ids.length || (!token && !u.param && !u.guid && !u.bip))
            return [];
        return ids.map((id) => {
            const entry = { elementId: id, value: val };
            if (u.guid)
                entry.guid = u.guid;
            if (u.bip)
                entry.bip = u.bip;
            if (u.param || u.parameter || u.paramName || u.parameterName) {
                entry.param = u.param ?? u.parameter ?? u.paramName ?? u.parameterName;
            }
            else if (token) {
                if (token.kind === "guid")
                    entry.guid = token.value;
                else if (token.kind === "bip")
                    entry.bip = token.value;
                else
                    entry.param = token.value;
            }
            for (const k of Object.keys(u)) {
                if (!(k in entry) && !["elementIds", "useSelection"].includes(k))
                    entry[k] = u[k];
            }
            return entry;
        });
    });
    return { updates: norm };
}
function normalizeParamsSetWhereArgs(raw) {
    const args = raw ?? {};
    const where = parseMaybeJson(args.where) ?? {};
    let setArr = parseMaybeJson(args.set);
    if (!Array.isArray(setArr))
        setArr = [];
    setArr = setArr
        .filter((s) => s && typeof s === "object")
        .map((s) => {
        const token = pickToken(s);
        const entry = { value: toBoolNum(s.value) };
        if (s.guid)
            entry.guid = s.guid;
        if (s.bip)
            entry.bip = s.bip;
        if (s.param || s.parameter || s.paramName || s.parameterName) {
            entry.param = s.param ?? s.parameter ?? s.paramName ?? s.parameterName;
        }
        else if (token) {
            if (token.kind === "guid")
                entry.guid = token.value;
            else if (token.kind === "bip")
                entry.bip = token.value;
            else
                entry.param = token.value;
        }
        for (const k of Object.keys(s))
            if (!(k in entry))
                entry[k] = s[k];
        return entry;
    });
    const out = { where, set: setArr };
    // C# soporta allowTypeParams (default true). Lo propagamos si llega.
    if (typeof args.allowTypeParams === "boolean") {
        out.allowTypeParams = args.allowTypeParams;
    }
    if (typeof args.dryRun === "boolean") {
        out.dryRun = !!args.dryRun; // será ignorado por C#, pero útil para logging o futuras extensiones
    }
    return out;
}
// ------------------------ Zod shapes (laxos) ------------------------
const PARAMS_GET_SHAPE = {
    elementId: z.number().int().optional(),
    elementIds: z.array(z.number().int()).optional(),
    // single-call:
    param: z.string().optional(),
    parameter: z.string().optional(),
    // multi:
    paramNames: z.array(z.string()).optional(),
    includeValueString: z.boolean().optional(),
};
const PARAMS_SET_SHAPE = {
    updates: z.union([z.string(), z.array(z.any())]).optional(),
    elementId: z.number().int().optional(),
    elementIds: z.array(z.number().int()).optional(),
    useSelection: z.boolean().optional(),
    param: z.string().optional(),
    parameter: z.string().optional(),
    paramName: z.string().optional(),
    parameterName: z.string().optional(),
    bip: z.string().optional(),
    guid: z.string().optional(),
    value: z.any().optional(),
};
const PARAMS_SET_WHERE_SHAPE = {
    where: z.union([z.any(), z.string()]).optional(),
    set: z.union([z.string(), z.array(z.any())]).optional(),
    allowTypeParams: z.boolean().optional(), // NUEVO
    dryRun: z.boolean().optional(),
};
const PARAMS_BULK_SHAPE = {
    updates: z.union([z.string(), z.array(z.any())]).optional(),
};
const PARAMS_LIST_SHAPE = {
    elementIds: z.array(z.number().int()),
    includeType: z.boolean().optional(), // default true en C#
};
// ------------------------ Tools ------------------------
server.registerTool("params_get", {
    description: "Lee parámetros de elementos. Ej: {elementIds:[344900], paramNames:['Comments'], includeValueString:true}",
    inputSchema: PARAMS_GET_SHAPE,
}, async (args) => {
    try {
        const normalized = normalizeParamsGetArgs(args);
        const res = await postRevit("params.get", normalized);
        return asText(res);
    }
    catch (e) {
        return asText({ ok: false, error: String(e?.message ?? e) });
    }
});
server.registerTool("params_set", {
    description: "Actualiza parámetros. Ej: {updates:[{elementId:344900,param:'Comments',value:'texto'}]}",
    inputSchema: PARAMS_SET_SHAPE,
}, async (args) => {
    try {
        console.error("=== params_set ===");
        console.error("RAW ARGS:", JSON.stringify(args, null, 2));
        const normalized = normalizeParamsSetArgs(args);
        console.error("NORMALIZED:", JSON.stringify(normalized, null, 2));
        const result = await postRevit("params.set", normalized);
        console.error("RESULT:", JSON.stringify(result, null, 2));
        return asText(result);
    }
    catch (e) {
        console.error("ERROR params_set:", e);
        return asText({ ok: false, error: String(e?.message ?? e) });
    }
});
server.registerTool("params_set_where", {
    description: "Actualiza parámetros con filtros. Ej: {where:{categories:['OST_Walls']}, set:[{param:'Comments',value:'ok'}], allowTypeParams:false}",
    inputSchema: PARAMS_SET_WHERE_SHAPE,
}, async (args) => {
    try {
        console.error("=== params_set_where ===");
        console.error("RAW ARGS:", JSON.stringify(args, null, 2));
        const normalized = normalizeParamsSetWhereArgs(args);
        console.error("NORMALIZED:", JSON.stringify(normalized, null, 2));
        const result = await postRevit("params.set_where", normalized);
        console.error("RESULT:", JSON.stringify(result, null, 2));
        return asText(result);
    }
    catch (e) {
        console.error("ERROR params_set_where:", e);
        return asText({ ok: false, error: String(e?.message ?? e) });
    }
});
server.registerTool("params_bulk_from_table", {
    description: "Actualiza parámetros desde tabla. {updates:[{elementId,param,value}]}",
    inputSchema: PARAMS_BULK_SHAPE,
}, async (args) => {
    try {
        const normalized = { updates: parseMaybeJson(args.updates) ?? [] };
        const res = await postRevit("params.bulk_from_table", normalized);
        return asText(res);
    }
    catch (e) {
        return asText({ ok: false, error: String(e?.message ?? e) });
    }
});
// NUEVO: listado de parámetros (instancia / tipo)
server.registerTool("params_list", {
    description: "Lista parámetros de elementos (instancia y/o tipo). Ej: {elementIds:[344900], includeType:true}",
    inputSchema: PARAMS_LIST_SHAPE,
}, async (args) => {
    try {
        const res = await postRevit("params.list", args);
        return asText(res);
    }
    catch (e) {
        return asText({ ok: false, error: String(e?.message ?? e) });
    }
});
// ------------------------ Transport ------------------------
const transport = new StdioServerTransport();
await server.connect(transport);
