import "dotenv/config";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { postRevit } from "./bridge.js";
// Crea el servidor MCP
const server = new McpServer({
    name: "mcp-arch",
    version: "1.0.0",
});
// Helper: empaqueta cualquier objeto como texto (JSON pretty)
const asText = (obj) => ({
    content: [{ type: "text", text: JSON.stringify(obj, null, 2) }],
});
// ===== Shapes / Schemas reutilizables =====
const Pt2Shape = { x: z.number(), y: z.number() };
const Pt2Schema = z.object(Pt2Shape);
// ===== wall.create =====
const WallCreateShape = {
    level: z.string().optional(),
    wallType: z.string().optional(),
    start: z.object(Pt2Shape),
    end: z.object(Pt2Shape),
    height_m: z.number().optional(),
    structural: z.boolean().optional(),
};
const WallCreateSchema = z.object(WallCreateShape);
server.registerTool("arch_wall_create", {
    title: "Create Wall",
    description: "Crea un muro recto con tipo/nivel opcional",
    inputSchema: WallCreateShape,
}, async (args) => {
    const result = await postRevit("wall.create", args);
    return asText(result);
});
// ===== level.create =====
const LevelCreateShape = {
    elevation_m: z.number(),
    name: z.string().optional(),
};
const LevelCreateSchema = z.object(LevelCreateShape);
server.registerTool("arch_level_create", {
    title: "Create Level",
    description: "Crea un nivel a cierta elevación (m)",
    inputSchema: LevelCreateShape,
}, async (args) => {
    const result = await postRevit("level.create", args);
    return asText(result);
});
// ===== grid.create =====
const GridCreateShape = {
    start: z.object(Pt2Shape),
    end: z.object(Pt2Shape),
    name: z.string().optional(),
};
const GridCreateSchema = z.object(GridCreateShape);
server.registerTool("arch_grid_create", {
    title: "Create Grid",
    description: "Crea una retícula",
    inputSchema: GridCreateShape,
}, async (args) => {
    const result = await postRevit("grid.create", args);
    return asText(result);
});
// ===== floor.create =====
const FloorCreateShape = {
    level: z.string().optional(),
    floorType: z.string().optional(),
    profile: z.array(z.object(Pt2Shape)).min(3),
};
const FloorCreateSchema = z.object(FloorCreateShape);
server.registerTool("arch_floor_create", {
    title: "Create Floor",
    description: "Crea un piso por contorno en un nivel",
    inputSchema: FloorCreateShape,
}, async (args) => {
    const result = await postRevit("floor.create", args);
    return asText(result);
});
// ===== ceiling.create (ahora funcional) =====
const CeilingCreateShape = {
    level: z.string().optional(),
    ceilingType: z.string().optional(),
    profile: z.array(z.object(Pt2Shape)).min(3),
    baseOffset_m: z.number().optional(), // altura sobre el nivel
};
const CeilingCreateSchema = z.object(CeilingCreateShape);
server.registerTool("arch_ceiling_create", {
    title: "Create Ceiling (Footprint/Sketch)",
    description: "Crea un plafón por contorno en un nivel (Revit 2022+).",
    inputSchema: CeilingCreateShape,
}, async (args) => {
    const result = await postRevit("ceiling.create", args);
    return asText(result);
});
// ===== door.place =====
const DoorPlaceShape = {
    hostWallId: z.number().int().optional(),
    level: z.string().optional(),
    familySymbol: z.string().optional(),
    point: z.object(Pt2Shape).optional(),
    offset_m: z.number().optional(),
    offsetAlong_m: z.number().optional(),
    alongNormalized: z.number().min(0).max(1).optional(),
    flipHand: z.boolean().optional(),
    flipFacing: z.boolean().optional(),
};
const DoorPlaceSchema = z.object(DoorPlaceShape);
server.registerTool("arch_door_place", {
    title: "Place Door",
    description: "Coloca una puerta. hostWallId/point pueden omitirse; el bridge resuelve por selección o muro cercano. Soporta offsetAlong_m / alongNormalized.",
    inputSchema: DoorPlaceShape,
}, async (args) => {
    const result = await postRevit("door.place", args);
    return asText(result);
});
// ===== window.place =====
const WindowPlaceShape = DoorPlaceShape;
const WindowPlaceSchema = DoorPlaceSchema;
server.registerTool("arch_window_place", {
    title: "Place Window",
    description: "Coloca una ventana. hostWallId/point pueden omitirse; el bridge resuelve por selección o muro cercano. Soporta offsetAlong_m / alongNormalized.",
    inputSchema: WindowPlaceShape,
}, async (args) => {
    const result = await postRevit("window.place", args);
    return asText(result);
});
// ===== rooms.create_on_levels =====
const RoomsCreateOnLevelsShape = {
    levelNames: z.array(z.string()).optional(),
    placeOnlyEnclosed: z.boolean().optional(),
};
const RoomsCreateOnLevelsSchema = z.object(RoomsCreateOnLevelsShape);
server.registerTool("arch_rooms_create_on_levels", {
    title: "Create Rooms on Levels",
    description: "Crea habitaciones automáticamente en los niveles indicados (o en todos si no se especifican).",
    inputSchema: RoomsCreateOnLevelsShape,
}, async (args) => {
    const result = await postRevit("rooms.create_on_levels", args);
    return asText(result);
});
// ===== floors.from_rooms =====
const FloorsFromRoomsShape = {
    roomIds: z.array(z.number().int()).min(1),
    floorType: z.string().optional(),
    baseOffset_m: z.number().optional(),
};
const FloorsFromRoomsSchema = z.object(FloorsFromRoomsShape);
server.registerTool("arch_floors_from_rooms", {
    title: "Create Floors from Rooms",
    description: "Crea pisos siguiendo el borde de las habitaciones. Acepta floorType y baseOffset_m.",
    inputSchema: FloorsFromRoomsShape,
}, async (args) => {
    const result = await postRevit("floors.from_rooms", args);
    return asText(result);
});
// ===== ceilings.from_rooms (ahora funcional) =====
const CeilingsFromRoomsShape = {
    roomIds: z.array(z.number().int()).min(1),
    ceilingType: z.string().optional(),
    baseOffset_m: z.number().optional(),
    useFinishBoundaries: z.boolean().optional(),
};
const CeilingsFromRoomsSchema = z.object(CeilingsFromRoomsShape);
server.registerTool("arch_ceilings_from_rooms", {
    title: "Create Ceilings from Rooms",
    description: "Crea plafones a partir de habitaciones (Revit 2022+).",
    inputSchema: CeilingsFromRoomsShape,
}, async (args) => {
    const result = await postRevit("ceilings.from_rooms", args);
    return asText(result);
});
// ===== roof.create_footprint =====
const RoofCreateFootprintShape = {
    level: z.string(), // requerido por el bridge
    roofType: z.string().optional(),
    profile: z.array(z.object(Pt2Shape)).min(3),
    slope: z.number().optional(), // grados
};
const RoofCreateFootprintSchema = z.object(RoofCreateFootprintShape);
server.registerTool("arch_roof_create_footprint", {
    title: "Create Roof (Footprint)",
    description: "Crea una cubierta por huella en un nivel, con perfil cerrado y pendiente opcional (grados).",
    inputSchema: RoofCreateFootprintShape,
}, async (args) => {
    const result = await postRevit("roof.create_footprint", args);
    return asText(result);
});
// ===== families.load =====
const FamilyLoadShape = {
    path: z.string(), // ruta al .rfa
    overwriteExisting: z.boolean().optional(),
};
const FamilyLoadSchema = z.object(FamilyLoadShape);
server.registerTool("arch_family_load", {
    title: "Load Family (RFA)",
    description: "Carga una familia .rfa al proyecto.",
    inputSchema: FamilyLoadShape,
}, async (args) => {
    const result = await postRevit("family.load", args);
    return asText(result);
});
// ===== family.place (non-hosted) =====
const FamilyPlaceShape = {
    familySymbol: z.string(), // "Familia: Tipo" o "Tipo"
    point: z.object(Pt2Shape), // XY metros
    level: z.string().optional(),
};
const FamilyPlaceSchema = z.object(FamilyPlaceShape);
server.registerTool("arch_family_place", {
    title: "Place Family (Non-Hosted)",
    description: "Coloca una instancia de familia no-hosted (p.ej. Generic Model). Para hosteadas usa herramientas específicas (puertas/ventanas).",
    inputSchema: FamilyPlaceShape,
}, async (args) => {
    const result = await postRevit("family.place", args);
    return asText(result);
});
// ===== Railing (stub) =====
const RailingCreateShape = {
    level: z.string().optional(),
    railingType: z.string().optional(),
    path: z.array(z.object(Pt2Shape)).min(2),
};
const RailingCreateSchema = z.object(RailingCreateShape);
server.registerTool("arch_railing_create", {
    title: "Create Railing (stub)",
    description: "Barandal independiente a lo largo de un path (pendiente de implementación).",
    inputSchema: RailingCreateShape,
}, async (args) => {
    const result = await postRevit("railing.create", args);
    return asText(result);
});
// ===== Stairs (stub) =====
const StairCreateShape = {
    baseLevel: z.string(),
    topLevel: z.string(),
    runWidth_m: z.number().optional(),
    riserHeight_m: z.number().optional(),
    treadDepth_m: z.number().optional(),
};
const StairCreateSchema = z.object(StairCreateShape);
server.registerTool("arch_stair_create", {
    title: "Create Stair (stub)",
    description: "Tramo recto entre niveles (epic, pendiente de implementación).",
    inputSchema: StairCreateShape,
}, async (args) => {
    const result = await postRevit("stair.create", args);
    return asText(result);
});
// ===== Ramp (stub) =====
const RampCreateShape = {
    baseLevel: z.string().optional(),
    topLevel: z.string().optional(),
    slopePercent: z.number().optional(),
    path: z.array(z.object(Pt2Shape)).min(2).optional(),
};
const RampCreateSchema = z.object(RampCreateShape);
server.registerTool("arch_ramp_create", {
    title: "Create Ramp (stub)",
    description: "Rampa por path o por niveles (pendiente de implementación).",
    inputSchema: RampCreateShape,
}, async (args) => {
    const result = await postRevit("ramp.create", args);
    return asText(result);
});
// Arrancar stdio
const transport = new StdioServerTransport();
await server.connect(transport);
