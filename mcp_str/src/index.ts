// mcp_str/src/index.ts
import "dotenv/config";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { postRevit } from "./bridge.js";

// Servidor MCP
const server = new McpServer({
  name: "mcp-str",
  version: "1.0.0",
});

// Respuesta como texto (JSON pretty)
const asText = (obj: unknown) => ({
  content: [{ type: "text" as const, text: JSON.stringify(obj, null, 2) }],
});

/* ============================
   struct.beam.create
   ============================ */
const BeamCreateShape = {
  level: z.string().optional(),          // el bridge resuelve activo si falta
  familyType: z.string().optional(),     // "Family: Type" o solo type
  elevation_m: z.number().optional(),    // en el bridge default = 3.0
  start: z.object({ x: z.number(), y: z.number() }),
  end: z.object({ x: z.number(), y: z.number() }),
};
const BeamCreateSchema = z.object(BeamCreateShape);

server.registerTool(
  "struct_beam_create",
  {
    title: "Create Structural Beam",
    description: "Crea una viga entre dos puntos XY en un nivel (opcional).",
    inputSchema: BeamCreateShape,
  },
  async (args: z.infer<typeof BeamCreateSchema>) => {
    const result = await postRevit("struct.beam.create", args);
    return asText(result);
  }
);

/* ============================
   struct.column.create
   ============================ */
const ColumnCreateShape = {
  level: z.string().optional(),
  familyType: z.string().optional(),     // "Family: Type" o solo type
  elevation_m: z.number().optional(),    // base offset; default 0.0 en bridge
  point: z.object({ x: z.number(), y: z.number() }),
};
const ColumnCreateSchema = z.object(ColumnCreateShape);

server.registerTool(
  "struct_column_create",
  {
    title: "Create Structural Column",
    description: "Crea una columna estructural en un punto XY (nivel opcional).",
    inputSchema: ColumnCreateShape,
  },
  async (args: z.infer<typeof ColumnCreateSchema>) => {
    const result = await postRevit("struct.column.create", args);
    return asText(result);
  }
);

/* ============================
   struct.floor.create
   ============================ */
const SFloorCreateShape = {
  level: z.string().optional(),
  floorType: z.string().optional(),
  profile: z.array(z.object({ x: z.number(), y: z.number() })).min(3), // polígono cerrado
};
const SFloorCreateSchema = z.object(SFloorCreateShape);

server.registerTool(
  "struct_floor_create",
  {
    title: "Create Structural Floor",
    description: "Crea un piso estructural por contorno en un nivel.",
    inputSchema: SFloorCreateShape,
  },
  async (args: z.infer<typeof SFloorCreateSchema>) => {
    const result = await postRevit("struct.floor.create", args);
    return asText(result);
  }
);

/* ============================
   struct.columns.place_on_grid
   ============================ */
const ColumnsPlaceOnGridShape = {
  baseLevel: z.string(),                    // requerido
  topLevel: z.string().optional(),          // opcional (usa baseLevel si falta)
  familyType: z.string(),                   // requerido ("Family: Type" o solo Type)

  gridX: z.array(z.string()).optional(),    // p.ej. ["A-E"] o ["A","B","C"]
  gridY: z.array(z.string()).optional(),    // p.ej. ["1-8"] o ["1","2","3"]
  gridNames: z.array(z.string()).optional(),// alternativa: lista plana (el server separa X/Y por dirección)

  baseOffset_m: z.number().optional(),      // default 0
  topOffset_m: z.number().optional(),       // default 0

  onlyIntersectionsInsideActiveCrop: z.boolean().optional(), // filtra por crop de vista activa
  tolerance_m: z.number().optional(),       // default 0.05
  skipIfColumnExistsNearby: z.boolean().optional(),

  worksetName: z.string().optional(),
  pinned: z.boolean().optional(),

  orientationRelativeTo: z.enum(["X", "Y", "None"]).optional(), // rotación respecto a Z
};
const ColumnsPlaceOnGridSchema = z.object(ColumnsPlaceOnGridShape);

server.registerTool(
  "struct_columns_place_on_grid",
  {
    title: "Place Columns on Grid",
    description:
      "Coloca columnas estructurales en intersecciones de ejes (rangos A-E / 1-8 o gridNames).",
    inputSchema: ColumnsPlaceOnGridShape,
  },
  async (args: z.infer<typeof ColumnsPlaceOnGridSchema>) => {
    const result = await postRevit("struct.columns.place_on_grid", args);
    return asText(result);
  }
);

/* ============================
   struct.foundation.isolated.create
   ============================ */
const FoundationIsolatedShape = {
  level: z.string(),                                     // requerido
  familyType: z.string().optional(),                     // "Family: Type" o solo Type
  point: z.object({ x: z.number(), y: z.number() }),     // XY en m
  baseOffset_m: z.number().optional(),
  pinned: z.boolean().optional(),
};
const FoundationIsolatedSchema = z.object(FoundationIsolatedShape);

server.registerTool(
  "struct_foundation_isolated_create",
  {
    title: "Create Isolated Foundation",
    description: "Crea una zapata aislada (Structural Foundation) en un punto XY.",
    inputSchema: FoundationIsolatedShape,
  },
  async (args: z.infer<typeof FoundationIsolatedSchema>) => {
    const result = await postRevit("struct.foundation.isolated.create", args);
    return asText(result);
  }
);

/* ============================
   struct.foundation.wall.create
   ============================ */
const WallFoundationShape = {
  wallId: z.number().int(),              // requerido
  typeName: z.string().optional(),       // nombre de WallFoundationType
  baseOffset_m: z.number().optional(),
};
const WallFoundationSchema = z.object(WallFoundationShape);

server.registerTool(
  "struct_foundation_wall_create",
  {
    title: "Create Wall Foundation",
    description: "Crea una cimentación continua bajo un muro existente.",
    inputSchema: WallFoundationShape,
  },
  async (args: z.infer<typeof WallFoundationSchema>) => {
    const result = await postRevit("struct.foundation.wall.create", args);
    return asText(result);
  }
);

/* ============================
   struct.beamsystem.create
   ============================ */
const BeamSystemShape = {
  level: z.string(),                                       // requerido
  profile: z.array(z.object({ x: z.number(), y: z.number() })).min(3),
  beamType: z.string().optional(),                         // "Family: Type" o solo Type
  direction: z.enum(["X", "Y"]).optional(),
  spacing_m: z.number().optional(),
  is3D: z.boolean().optional(),
};
const BeamSystemSchema = z.object(BeamSystemShape);

server.registerTool(
  "struct_beamsystem_create",
  {
    title: "Create Beam System",
    description: "Crea un sistema de vigas por contorno; dirección X/Y opcional.",
    inputSchema: BeamSystemShape,
  },
  async (args: z.infer<typeof BeamSystemSchema>) => {
    const result = await postRevit("struct.beamsystem.create", args);
    return asText(result);
  }
);

/* ============================
   struct.rebar.add_straight_on_beam
   ============================ */
const RebarOnBeamShape = {
  hostId: z.number().int(),              // Id de viga (Structural Framing)
  barType: z.string().optional(),        // nombre de RebarBarType
  number: z.number().int().optional(),   // default 1
  spacing_m: z.number().optional(),      // si number>1 y se quiere espaciar
  zOffset_m: z.number().optional(),      // offset vertical simple
};
const RebarOnBeamSchema = z.object(RebarOnBeamShape);

server.registerTool(
  "struct_rebar_add_straight_on_beam",
  {
    title: "Add Straight Rebar On Beam",
    description: "Coloca barras rectas a lo largo del eje de una viga.",
    inputSchema: RebarOnBeamShape,
  },
  async (args: z.infer<typeof RebarOnBeamSchema>) => {
    const result = await postRevit("struct.rebar.add_straight_on_beam", args);
    return asText(result);
  }
);

// stdio
const transport = new StdioServerTransport();
await server.connect(transport);
