// mcp_qa/src/index.ts
import "dotenv/config";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { postRevit } from "./bridge.js";
// Server MCP
const server = new McpServer({
    name: "mcp-qa",
    version: "1.0.0",
});
// Helper: devolver JSON pretty como texto
const asText = (obj) => ({
    content: [{ type: "text", text: JSON.stringify(obj, null, 2) }],
});
// Shape sin argumentos
const NoArgs = {};
/* =========================================
   qa.fix.pin_all_links
   ========================================= */
server.registerTool("qa_pin_all_links", {
    title: "Pin all Revit links",
    description: "Pone pin a todos los Revit Link Instances del modelo.",
    inputSchema: NoArgs,
}, async () => asText(await postRevit("qa.fix.pin_all_links", {})));
/* =========================================
   qa.fix.delete_imports
   ========================================= */
server.registerTool("qa_delete_imports", {
    title: "Delete all CAD imports",
    description: "Elimina todos los ImportInstance (CAD) del documento.",
    inputSchema: NoArgs,
}, async () => asText(await postRevit("qa.fix.delete_imports", {})));
/* =========================================
   qa.fix.apply_view_templates
   - soporta autoPickFirst y devuelve lista si hay varias plantillas
   ========================================= */
const ApplyViewTemplatesShape = {
    templateName: z.string().optional(),
    templateId: z.number().int().optional(),
    onlyWithoutTemplate: z.boolean().optional(), // default true en C#
    viewIds: z.array(z.number().int()).optional(),
    autoPickFirst: z.boolean().optional(),
};
const ApplyViewTemplatesSchema = z.object(ApplyViewTemplatesShape);
server.registerTool("qa_apply_view_templates", {
    title: "Apply view template to views",
    description: "Aplica una plantilla a vistas. Si no se especifica plantilla y hay varias, devuelve la lista de opciones.",
    inputSchema: ApplyViewTemplatesShape,
}, async (rawArgs) => {
    const args = ApplyViewTemplatesSchema.parse(rawArgs);
    return asText(await postRevit("qa.fix.apply_view_templates", args));
});
/* =========================================
   qa.fix.remove_textnotes
   ========================================= */
const RemoveTextNotesShape = {
    viewId: z.number().int().optional(),
};
const RemoveTextNotesSchema = z.object(RemoveTextNotesShape);
server.registerTool("qa_remove_text_notes", {
    title: "Remove text notes",
    description: "Elimina todas las TextNotes de la vista activa o de la vista indicada por viewId.",
    inputSchema: RemoveTextNotesShape,
}, async (rawArgs) => {
    const args = RemoveTextNotesSchema.parse(rawArgs);
    return asText(await postRevit("qa.fix.remove_textnotes", args));
});
/* =========================================
   qa.fix.delete_unused_types
   ========================================= */
server.registerTool("qa_delete_unused_types", {
    title: "Delete unused element types",
    description: "Intenta borrar ElementTypes no usados (excluye ViewFamilyType y TitleBlocks).",
    inputSchema: NoArgs,
}, async () => asText(await postRevit("qa.fix.delete_unused_types", {})));
/* =========================================
   qa.fix.rename_views
   ========================================= */
const RenameViewsShape = {
    prefix: z.string().optional(),
    find: z.string().optional(),
    replace: z.string().optional(),
    viewIds: z.array(z.number().int()).optional(),
};
const RenameViewsSchema = z.object(RenameViewsShape);
server.registerTool("qa_rename_views", {
    title: "Rename views",
    description: "Renombra vistas (find/replace y/o prefix). Si no se pasan viewIds, aplica a todas las vistas no plantilla.",
    inputSchema: RenameViewsShape,
}, async (rawArgs) => {
    const args = RenameViewsSchema.parse(rawArgs);
    return asText(await postRevit("qa.fix.rename_views", args));
});
/* =========================================
   NUEVO: qa.fix.unhide_all_in_view
   ========================================= */
const UnhideAllInViewShape = {
    viewId: z.number().int().optional(),
    unhideCategories: z.boolean().optional(), // default true en C#
    clearTemporaryHideIsolate: z.boolean().optional(), // default true en C#
};
const UnhideAllInViewSchema = z.object(UnhideAllInViewShape);
server.registerTool("qa_unhide_all_in_view", {
    title: "Unhide all in view",
    description: "Desoculta elementos/categorías y limpia Temporary Hide/Isolate en la vista activa o indicada.",
    inputSchema: UnhideAllInViewShape,
}, async (rawArgs) => {
    const args = UnhideAllInViewSchema.parse(rawArgs);
    return asText(await postRevit("qa.fix.unhide_all_in_view", args));
});
/* =========================================
   NUEVO: qa.fix.delete_unused_view_templates
   ========================================= */
server.registerTool("qa_delete_unused_view_templates", {
    title: "Delete unused view templates",
    description: "Elimina plantillas de vista no referenciadas por ninguna vista.",
    inputSchema: NoArgs,
}, async () => asText(await postRevit("qa.fix.delete_unused_view_templates", {})));
/* =========================================
   NUEVO: qa.fix.delete_unused_view_filters
   ========================================= */
server.registerTool("qa_delete_unused_view_filters", {
    title: "Delete unused view filters",
    description: "Elimina filtros (ParameterFilterElement) no usados por ninguna vista/plantilla.",
    inputSchema: NoArgs,
}, async () => asText(await postRevit("qa.fix.delete_unused_view_filters", {})));
// stdio transport
const transport = new StdioServerTransport();
await server.connect(transport);
