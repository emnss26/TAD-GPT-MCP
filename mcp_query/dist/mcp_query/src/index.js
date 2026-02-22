import "dotenv/config";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { postRevit } from "./bridge.js";
import { snapshotContext } from "../../../TAD-GPT-MCP/mcp_common/src/context.js";
const server = new McpServer({ name: "mcp-query", version: "1.0.0" });
const asText = (obj) => ({
    content: [{ type: "text", text: JSON.stringify(obj, null, 2) }],
});
// Opcional: helper para dejar claro que pasamos “shape” crudo
const asInputShape = (shape) => shape;
// Shape vacío reutilizable
const NoArgs = asInputShape({});
/* levels.list */
server.registerTool("levels_list", { title: "List levels", description: "Devuelve todos los niveles del documento.", inputSchema: NoArgs }, async () => asText(await postRevit("levels.list", {})));
/* walltypes.list */
server.registerTool("walltypes_list", { title: "List wall types", description: "Tipos de muro disponibles.", inputSchema: NoArgs }, async () => asText(await postRevit("walltypes.list", {})));
/* view.active */
server.registerTool("activeview_info", { title: "Active view info", description: "Info de la vista activa.", inputSchema: NoArgs }, async () => asText(await postRevit("view.active", {})));
/* views.list */
server.registerTool("views_list", { title: "List views", description: "Lista de vistas no-template.", inputSchema: NoArgs }, async () => asText(await postRevit("views.list", {})));
/* schedules.list */
server.registerTool("schedules_list", { title: "List schedules", description: "Lista de schedules.", inputSchema: NoArgs }, async () => asText(await postRevit("schedules.list", {})));
/* materials.list */
server.registerTool("materials_list", { title: "List materials", description: "Materiales del documento.", inputSchema: NoArgs }, async () => asText(await postRevit("materials.list", {})));
/* categories.list */
server.registerTool("categories_list", { title: "List categories", description: "Categorías (incluye BuiltInCategory).", inputSchema: NoArgs }, async () => asText(await postRevit("categories.list", {})));
/* families.types.list */
server.registerTool("families_types_list", { title: "List family types", description: "FamilySymbol (familia, tipo, categoría).", inputSchema: NoArgs }, async () => asText(await postRevit("families.types.list", {})));
/* links.list */
server.registerTool("links_list", { title: "List Revit links", description: "Instancias de Revit Link (id, nombre, pinned).", inputSchema: NoArgs }, async () => asText(await postRevit("links.list", {})));
/* imports.list */
server.registerTool("imports_list", { title: "List CAD imports", description: "ImportInstance en el documento.", inputSchema: NoArgs }, async () => asText(await postRevit("imports.list", {})));
/* worksets.list */
server.registerTool("worksets_list", { title: "List worksets", description: "Worksets (id, nombre, tipo).", inputSchema: NoArgs }, async () => asText(await postRevit("worksets.list", {})));
/* textnotes.find */
server.registerTool("textnotes_find", { title: "Find text notes in active view", description: "TextNotes (id, texto) en la vista activa.", inputSchema: NoArgs }, async () => asText(await postRevit("textnotes.find", {})));
/* ducttypes.list */
server.registerTool("ducttypes_list", { title: "List duct types", description: "Tipos de ducto (MEP).", inputSchema: NoArgs }, async () => asText(await postRevit("ducttypes.list", {})));
/* pipetypes.list */
server.registerTool("pipetypes_list", { title: "List pipe types", description: "Tipos de tubería (MEP).", inputSchema: NoArgs }, async () => asText(await postRevit("pipetypes.list", {})));
/* cabletraytypes.list */
server.registerTool("cabletraytypes_list", { title: "List cable tray types", description: "Tipos de charola.", inputSchema: NoArgs }, async () => asText(await postRevit("cabletraytypes.list", {})));
/* selection.info */
const SelectionInfoShape = asInputShape({
    includeParameters: z.boolean().optional(),
    topNParams: z.number().int().min(1).max(1000).optional(),
});
const SelectionInfoSchema = z.object({
    includeParameters: z.boolean().optional(),
    topNParams: z.number().int().min(1).max(1000).optional(),
});
server.registerTool("selection_info", { title: "Selection info", description: "Información de los elementos seleccionados.", inputSchema: SelectionInfoShape }, async (rawArgs) => asText(await postRevit("selection.info", SelectionInfoSchema.parse(rawArgs))));
/* element.info */
const ElementInfoShape = asInputShape({
    elementId: z.number().int(),
    includeParameters: z.boolean().optional(),
    topNParams: z.number().int().min(1).max(1000).optional(),
});
const ElementInfoSchema = z.object({
    elementId: z.number().int(),
    includeParameters: z.boolean().optional(),
    topNParams: z.number().int().min(1).max(1000).optional(),
});
server.registerTool("element_info", { title: "Element info by id", description: "Información detallada por Id.", inputSchema: ElementInfoShape }, async (rawArgs) => asText(await postRevit("element.info", ElementInfoSchema.parse(rawArgs))));
/* context snapshot */
const ContextSnapshotShape = asInputShape({
    cacheSec: z.number().int().min(1).max(3600).optional(),
});
const ContextSnapshotSchema = z.object({
    cacheSec: z.number().int().min(1).max(3600).optional(),
});
server.registerTool("query_context_snapshot", { title: "Revit Context Snapshot", description: "Vista activa, niveles, grids, tipos, worksets, selección, etc.", inputSchema: ContextSnapshotShape }, async (rawArgs) => {
    const { cacheSec = 30 } = ContextSnapshotSchema.parse(rawArgs);
    const snap = await snapshotContext(postRevit, { cacheSec });
    return asText(snap);
});
const transport = new StdioServerTransport();
await server.connect(transport);
