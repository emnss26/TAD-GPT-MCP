import "dotenv/config";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { postRevit } from "./bridge.js";

const server = new McpServer({ name: "mcp-qto", version: "1.0.0" });
const asText = (obj: unknown) => ({ content: [{ type: "text" as const, text: JSON.stringify(obj, null, 2) }] });
const asInputShape = <T extends Record<string, unknown>>(shape: T) => shape as unknown as Record<string, any>;

/* ========================== Walls ========================== */
const WallsShape = {
  groupBy: z.array(z.enum(["type", "level", "phase"])).optional(),
  includeIds: z.boolean().optional(),
  filter: z.object({
    typeIds: z.array(z.number()).optional(),
    typeNames: z.array(z.string()).optional(),
    nameRegex: z.string().optional(),
    levels: z.array(z.string()).optional(),
    phase: z.string().optional(),
    useSelection: z.boolean().optional(),
  }).optional(),
};
server.registerTool(
  "qto_walls",
  { title: "QTO Walls", description: "Metrajes de muros por agrupación.", inputSchema: asInputShape(WallsShape) },
  async (rawArgs) => {
    const args = z.object(WallsShape).parse(rawArgs);
    return asText(await postRevit("qto.walls", args));
  }
);

/* ========================== Floors ========================== */
const FloorsShape = { groupBy: z.array(z.enum(["type", "level"])).optional(), includeIds: z.boolean().optional() };
server.registerTool(
  "qto_floors",
  { title: "QTO Floors", description: "Metrajes de pisos por agrupación.", inputSchema: asInputShape(FloorsShape) },
  async (rawArgs) => {
    const args = z.object(FloorsShape).parse(rawArgs);
    return asText(await postRevit("qto.floors", args));
  }
);

/* ========================== Ceilings (plafones) ========================== */
const CeilingsShape = { groupBy: z.array(z.enum(["type", "level"])).optional(), includeIds: z.boolean().optional() };
server.registerTool(
  "qto_ceilings",
  { title: "QTO Ceilings (Plafones)", description: "Áreas de plafones por agrupación.", inputSchema: asInputShape(CeilingsShape) },
  async (rawArgs) => {
    const args = z.object(CeilingsShape).parse(rawArgs);
    return asText(await postRevit("qto.ceilings", args));
  }
);

/* ========================== Railings (barandales) ========================== */
const RailingsShape = { groupBy: z.array(z.enum(["type", "level"])).optional(), includeIds: z.boolean().optional() };
server.registerTool(
  "qto_railings",
  { title: "QTO Railings", description: "Metros lineales de barandales por agrupación.", inputSchema: asInputShape(RailingsShape) },
  async (rawArgs) => {
    const args = z.object(RailingsShape).parse(rawArgs);
    return asText(await postRevit("qto.railings", args));
  }
);

/* ========================== Families (conteo) ========================== */
const FamiliesCountShape = {
  groupBy: z.array(z.enum(["category", "family", "type", "level"])).optional(),
  includeIds: z.boolean().optional(),
  categories: z.array(z.string()).optional(), // OST_* o sin OST_
};
server.registerTool(
  "qto_families_count",
  { title: "QTO Families Count", description: "Conteo de instancias por categoría/familia/tipo/level.", inputSchema: asInputShape(FamiliesCountShape) },
  async (rawArgs) => {
    const args = z.object(FamiliesCountShape).parse(rawArgs);
    return asText(await postRevit("qto.families.count", args));
  }
);

/* ========================== Struct Beams ========================== */
const StructBeamsShape = {
  groupBy: z.array(z.enum(["type", "level"])).optional(),
  includeIds: z.boolean().optional(),
  includeLength: z.boolean().optional(),
  includeVolume: z.boolean().optional(),
};
server.registerTool(
  "qto_struct_beams",
  { title: "QTO Structural Beams", description: "m.l. y/o m³ de vigas por agrupación.", inputSchema: asInputShape(StructBeamsShape) },
  async (rawArgs) => {
    const args = z.object(StructBeamsShape).parse(rawArgs);
    return asText(await postRevit("qto.struct.beams", args));
  }
);

/* ========================== Struct Columns ========================== */
const StructColumnsShape = {
  groupBy: z.array(z.enum(["type", "level"])).optional(),
  includeIds: z.boolean().optional(),
  includeLength: z.boolean().optional(),
  includeVolume: z.boolean().optional(),
};
server.registerTool(
  "qto_struct_columns",
  { title: "QTO Structural Columns", description: "m.l. y/o m³ de columnas por agrupación.", inputSchema: asInputShape(StructColumnsShape) },
  async (rawArgs) => {
    const args = z.object(StructColumnsShape).parse(rawArgs);
    return asText(await postRevit("qto.struct.columns", args));
  }
);

/* ========================== Struct Foundations ========================== */
const StructFoundationsShape = { groupBy: z.array(z.enum(["type", "level"])).optional(), includeIds: z.boolean().optional() };
server.registerTool(
  "qto_struct_foundations",
  { title: "QTO Structural Foundations", description: "Conteo y m³ de cimentaciones por agrupación.", inputSchema: asInputShape(StructFoundationsShape) },
  async (rawArgs) => {
    const args = z.object(StructFoundationsShape).parse(rawArgs);
    return asText(await postRevit("qto.struct.foundations", args));
  }
);

/* ========================== Struct Concrete (combo rápido/legacy) ========================== */
const StructConcreteShape = {
  includeBeams: z.boolean().optional(),
  includeColumns: z.boolean().optional(),
  includeFoundation: z.boolean().optional(),
  groupBy: z.array(z.enum(["type", "level"])).optional(),
  includeIds: z.boolean().optional(),
};
server.registerTool(
  "qto_struct_concrete",
  { title: "QTO Structural Concrete (Combo)", description: "Volúmenes y m.l. (vigas) por agrupación.", inputSchema: asInputShape(StructConcreteShape) },
  async (rawArgs) => {
    const args = z.object(StructConcreteShape).parse(rawArgs);
    return asText(await postRevit("qto.struct.concrete", args));
  }
);

/* ========================== MEP Pipes ========================== */
const PipesShape = {
  groupBy: z.array(z.enum(["system", "type", "level"])).optional(),
  diameterBucketsMm: z.array(z.number()).optional(),
  includeIds: z.boolean().optional(),
};
server.registerTool(
  "qto_mep_pipes",
  { title: "QTO MEP Pipes", description: "m.l. totales y por buckets de diámetro (mm).", inputSchema: asInputShape(PipesShape) },
  async (rawArgs) => {
    const args = z.object(PipesShape).parse(rawArgs);
    return asText(await postRevit("qto.mep.pipes", args));
  }
);

/* ========================== MEP Ducts ========================== */
const DuctsShape = {
  groupBy: z.array(z.enum(["system", "type", "level"])).optional(),
  roundVsRect: z.boolean().optional(),
  includeIds: z.boolean().optional(),
};
server.registerTool(
  "qto_mep_ducts",
  { title: "QTO MEP Ducts", description: "m.l. y área superficial estimada. Puede separar redondos/rectangulares.", inputSchema: asInputShape(DuctsShape) },
  async (rawArgs) => {
    const args = z.object(DuctsShape).parse(rawArgs);
    return asText(await postRevit("qto.mep.ducts", args));
  }
);

/* ========================== Cable Trays ========================== */
const CableTraysShape = { groupBy: z.array(z.enum(["type", "level"])).optional(), includeIds: z.boolean().optional() };
server.registerTool(
  "qto_mep_cabletrays",
  { title: "QTO Cable Trays", description: "Metros lineales de cable trays por agrupación.", inputSchema: asInputShape(CableTraysShape) },
  async (rawArgs) => {
    const args = z.object(CableTraysShape).parse(rawArgs);
    return asText(await postRevit("qto.mep.cabletrays", args));
  }
);

/* ========================== Conduits ========================== */
const ConduitsShape = { groupBy: z.array(z.enum(["type", "level"])).optional(), includeIds: z.boolean().optional() };
server.registerTool(
  "qto_mep_conduits",
  { title: "QTO Conduits", description: "Metros lineales de conduits por agrupación.", inputSchema: asInputShape(ConduitsShape) },
  async (rawArgs) => {
    const args = z.object(ConduitsShape).parse(rawArgs);
    return asText(await postRevit("qto.mep.conduits", args));
  }
);

/* ========================== Wrappers de conteo por categoría ========================== */
/* Forma común (no exponemos 'categories' porque se fija en el bridge) */
const CountShape = { groupBy: z.array(z.enum(["category", "family", "type", "level"])).optional(), includeIds: z.boolean().optional() };

server.registerTool(
  "qto_mep_duct_fittings_count",
  { title: "QTO Duct Fittings (Count)", description: "Conteo de Duct Fittings.", inputSchema: asInputShape(CountShape) },
  async (rawArgs) => asText(await postRevit("qto.mep.duct_fittings.count", z.object(CountShape).parse(rawArgs)))
);

server.registerTool(
  "qto_mep_air_terminals_count",
  { title: "QTO Air Terminals (Count)", description: "Conteo de Air Terminals.", inputSchema: asInputShape(CountShape) },
  async (rawArgs) => asText(await postRevit("qto.mep.air_terminals.count", z.object(CountShape).parse(rawArgs)))
);

server.registerTool(
  "qto_mep_mechanical_equip_count",
  { title: "QTO Mechanical Equipment (Count)", description: "Conteo de Mechanical Equipment.", inputSchema: asInputShape(CountShape) },
  async (rawArgs) => asText(await postRevit("qto.mep.mechanical_equip.count", z.object(CountShape).parse(rawArgs)))
);

server.registerTool(
  "qto_mep_pipe_fittings_count",
  { title: "QTO Pipe Fittings (Count)", description: "Conteo de Pipe Fittings.", inputSchema: asInputShape(CountShape) },
  async (rawArgs) => asText(await postRevit("qto.mep.pipe_fittings.count", z.object(CountShape).parse(rawArgs)))
);

server.registerTool(
  "qto_mep_pipe_accessories_count",
  { title: "QTO Pipe Accessories (Count)", description: "Conteo de Pipe Accessories.", inputSchema: asInputShape(CountShape) },
  async (rawArgs) => asText(await postRevit("qto.mep.pipe_accessories.count", z.object(CountShape).parse(rawArgs)))
);

server.registerTool(
  "qto_mep_plumbing_equipment_count",
  { title: "QTO Plumbing Equipment (Count)", description: "Conteo de Plumbing Equipment.", inputSchema: asInputShape(CountShape) },
  async (rawArgs) => asText(await postRevit("qto.mep.plumbing_equipment.count", z.object(CountShape).parse(rawArgs)))
);

server.registerTool(
  "qto_mep_plumbing_fixtures_count",
  { title: "QTO Plumbing Fixtures (Count)", description: "Conteo de Plumbing Fixtures.", inputSchema: asInputShape(CountShape) },
  async (rawArgs) => asText(await postRevit("qto.mep.plumbing_fixtures.count", z.object(CountShape).parse(rawArgs)))
);

server.registerTool(
  "qto_electrical_equipment_count",
  { title: "QTO Electrical Equipment (Count)", description: "Conteo de Electrical Equipment.", inputSchema: asInputShape(CountShape) },
  async (rawArgs) => asText(await postRevit("qto.electrical.equipment.count", z.object(CountShape).parse(rawArgs)))
);

server.registerTool(
  "qto_electrical_lighting_count",
  { title: "QTO Lighting Fixtures (Count)", description: "Conteo de Lighting Fixtures.", inputSchema: asInputShape(CountShape) },
  async (rawArgs) => asText(await postRevit("qto.electrical.lighting.count", z.object(CountShape).parse(rawArgs)))
);

server.registerTool(
  "qto_electrical_devices_count",
  { title: "QTO Electrical Devices (Count)", description: "Conteo de Electrical Fixtures / Lighting Devices.", inputSchema: asInputShape(CountShape) },
  async (rawArgs) => asText(await postRevit("qto.electrical.devices.count", z.object(CountShape).parse(rawArgs)))
);

/* ========================== Boot ========================== */
const transport = new StdioServerTransport();
await server.connect(transport);
