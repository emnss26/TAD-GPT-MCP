import "dotenv/config";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { postRevit } from "./bridge.js";

// Server MCP
const server = new McpServer({
  name: "mcp-mep",
  version: "1.0.0",
});

// Helpers
const asText = (obj: unknown) => ({
  content: [{ type: "text" as const, text: JSON.stringify(obj, null, 2) }],
});

// Tipado que espera el SDK para inputSchema
const asInputShape = <T extends Record<string, unknown>>(shape: T) =>
  shape as unknown as Record<string, any>;

// PT2 común
const PtShape = z.object({ x: z.number(), y: z.number() });

/* =========================================
   mep.duct.create
========================================= */
const DuctCreateShape = {
  level: z.string().optional(),
  systemType: z.string().optional(), // nombre o clasificación (SupplyAir, ReturnAir, etc)
  ductType: z.string().optional(),
  elevation_m: z.number().optional(), // por defecto en bridge: 2.7
  start: PtShape,
  end: PtShape,
  width_mm: z.number().optional(), // rectangular
  height_mm: z.number().optional(), // rectangular
  diameter_mm: z.number().optional(), // redondo
};
const DuctCreateSchema = z.object(DuctCreateShape);

server.registerTool(
  "mep_duct_create",
  {
    title: "Create Duct",
    description:
      "Crea un ducto entre dos puntos XY (m). Soporta systemType/ductType y dimensiones opcionales.",
    inputSchema: asInputShape(DuctCreateShape),
  },
  async (rawArgs) => {
    const args = DuctCreateSchema.parse(rawArgs);
    const result = await postRevit("mep.duct.create", args);
    return asText(result);
  }
);

/* =========================================
   mep.duct.connect
========================================= */
const DuctConnectShape = {
  aId: z.number().int(),
  bId: z.number().int(),
  mode: z.enum(["auto", "elbow", "transition"]).optional(),
};
const DuctConnectSchema = z.object(DuctConnectShape);

server.registerTool(
  "mep_duct_connect",
  {
    title: "Connect Ducts/Elements",
    description:
      "Conecta dos elementos MEP (ductos o familias con conectores). mode: auto | elbow | transition.",
    inputSchema: asInputShape(DuctConnectShape),
  },
  async (rawArgs) => {
    const args = DuctConnectSchema.parse(rawArgs);
    const result = await postRevit("mep.duct.connect", args);
    return asText(result);
  }
);

/* =========================================
   mep.duct.add_insulation
========================================= */
const DuctAddInsulationShape = {
  ductIds: z.array(z.number().int()).min(1),
  typeName: z.string().optional(),
  thickness_mm: z.number().optional(),
};
const DuctAddInsulationSchema = z.object(DuctAddInsulationShape);

server.registerTool(
  "mep_duct_add_insulation",
  {
    title: "Add Duct Insulation (Bulk)",
    description:
      "Agrega aislamiento a ductos; typeName opcional; thickness_mm opcional.",
    inputSchema: asInputShape(DuctAddInsulationShape),
  },
  async (rawArgs) => {
    const args = DuctAddInsulationSchema.parse(rawArgs);
    const result = await postRevit("mep.duct.add_insulation", args);
    return asText(result);
  }
);

/* =========================================
   mep.duct.add_lining
========================================= */
const DuctAddLiningShape = {
  ductIds: z.array(z.number().int()).min(1),
  typeName: z.string().optional(),
  thickness_mm: z.number().optional(),
};
const DuctAddLiningSchema = z.object(DuctAddLiningShape);

server.registerTool(
  "mep_duct_add_lining",
  {
    title: "Add Duct Lining (Bulk)",
    description:
      "Agrega recubrimiento interno (lining) a ductos; typeName/thickness_mm opcionales.",
    inputSchema: asInputShape(DuctAddLiningShape),
  },
  async (rawArgs) => {
    const args = DuctAddLiningSchema.parse(rawArgs);
    const result = await postRevit("mep.duct.add_lining", args);
    return asText(result);
  }
);

/* =========================================
   mep.duct.set_offset
========================================= */
const DuctSetOffsetShape = {
  ductIds: z.array(z.number().int()).min(1),
  offset_m: z.number(),
};
const DuctSetOffsetSchema = z.object(DuctSetOffsetShape);

server.registerTool(
  "mep_duct_set_offset",
  {
    title: "Set Duct Offset (Bulk)",
    description: "Ajusta el offset (m) de múltiples ductos.",
    inputSchema: asInputShape(DuctSetOffsetShape),
  },
  async (rawArgs) => {
    const args = DuctSetOffsetSchema.parse(rawArgs);
    const result = await postRevit("mep.duct.set_offset", args);
    return asText(result);
  }
);

/* =========================================
   mep.pipe.create
========================================= */
const PipeCreateShape = {
  level: z.string().optional(),
  systemType: z.string().optional(), // PipingSystemType (nombre)
  pipeType: z.string().optional(),
  elevation_m: z.number().optional(), // por defecto en bridge: 2.5
  start: PtShape,
  end: PtShape,
  diameter_mm: z.number().optional(),
};
const PipeCreateSchema = z.object(PipeCreateShape);

server.registerTool(
  "mep_pipe_create",
  {
    title: "Create Pipe",
    description:
      "Crea una tubería entre dos puntos XY (m). systemType/pipeType opcionales; diámetro opcional.",
    inputSchema: asInputShape(PipeCreateShape),
  },
  async (rawArgs) => {
    const args = PipeCreateSchema.parse(rawArgs);
    const result = await postRevit("mep.pipe.create", args);
    return asText(result);
  }
);

/* =========================================
   mep.pipe.connect
========================================= */
const PipeConnectShape = {
  aId: z.number().int(),
  bId: z.number().int(),
  mode: z.enum(["auto", "elbow", "transition"]).optional(),
};
const PipeConnectSchema = z.object(PipeConnectShape);

server.registerTool(
  "mep_pipe_connect",
  {
    title: "Connect Pipes/Elements",
    description:
      "Conecta dos elementos (pipes o familias con conectores). mode: auto | elbow | transition.",
    inputSchema: asInputShape(PipeConnectShape),
  },
  async (rawArgs) => {
    const args = PipeConnectSchema.parse(rawArgs);
    const result = await postRevit("mep.pipe.connect", args);
    return asText(result);
  }
);

/* =========================================
   mep.pipe.add_insulation
========================================= */
const PipeAddInsulationShape = {
  pipeIds: z.array(z.number().int()).min(1),
  typeName: z.string().optional(),
  thickness_mm: z.number().optional(),
};
const PipeAddInsulationSchema = z.object(PipeAddInsulationShape);

server.registerTool(
  "mep_pipe_add_insulation",
  {
    title: "Add Pipe Insulation (Bulk)",
    description:
      "Agrega aislamiento a pipes; typeName opcional; thickness_mm opcional.",
    inputSchema: asInputShape(PipeAddInsulationShape),
  },
  async (rawArgs) => {
    const args = PipeAddInsulationSchema.parse(rawArgs);
    const result = await postRevit("mep.pipe.add_insulation", args);
    return asText(result);
  }
);

/* =========================================
   mep.pipe.set_offset
========================================= */
const PipeSetOffsetShape = {
  pipeIds: z.array(z.number().int()).min(1),
  offset_m: z.number(),
};
const PipeSetOffsetSchema = z.object(PipeSetOffsetShape);

server.registerTool(
  "mep_pipe_set_offset",
  {
    title: "Set Pipe Offset (Bulk)",
    description: "Ajusta el offset (m) de múltiples pipes.",
    inputSchema: asInputShape(PipeSetOffsetShape),
  },
  async (rawArgs) => {
    const args = PipeSetOffsetSchema.parse(rawArgs);
    const result = await postRevit("mep.pipe.set_offset", args);
    return asText(result);
  }
);

/* =========================================
   mep.conduit.create
========================================= */
const ConduitCreateShape = {
  level: z.string().optional(),
  conduitType: z.string().optional(),
  elevation_m: z.number().optional(), // por defecto en bridge: 2.4
  start: PtShape,
  end: PtShape,
  diameter_mm: z.number().optional(),
};
const ConduitCreateSchema = z.object(ConduitCreateShape);

server.registerTool(
  "mep_conduit_create",
  {
    title: "Create Conduit",
    description: "Crea un conduit entre dos puntos XY (m). Tipo y diámetro opcionales.",
    inputSchema: asInputShape(ConduitCreateShape),
  },
  async (rawArgs) => {
    const args = ConduitCreateSchema.parse(rawArgs);
    const result = await postRevit("mep.conduit.create", args);
    return asText(result);
  }
);

/* =========================================
   mep.conduit.set_offset
========================================= */
const ConduitSetOffsetShape = {
  conduitIds: z.array(z.number().int()).min(1),
  offset_m: z.number(),
};
const ConduitSetOffsetSchema = z.object(ConduitSetOffsetShape);

server.registerTool(
  "mep_conduit_set_offset",
  {
    title: "Set Conduit Offset (Bulk)",
    description: "Ajusta el offset (m) de múltiples conduits.",
    inputSchema: asInputShape(ConduitSetOffsetShape),
  },
  async (rawArgs) => {
    const args = ConduitSetOffsetSchema.parse(rawArgs);
    const result = await postRevit("mep.conduit.set_offset", args);
    return asText(result);
  }
);

/* =========================================
   mep.conduit.set_diameter
========================================= */
const ConduitSetDiameterShape = {
  conduitIds: z.array(z.number().int()).min(1),
  diameter_mm: z.number(),
};
const ConduitSetDiameterSchema = z.object(ConduitSetDiameterShape);

server.registerTool(
  "mep_conduit_set_diameter",
  {
    title: "Set Conduit Diameter (Bulk)",
    description: "Define el diámetro (mm) para múltiples conduits.",
    inputSchema: asInputShape(ConduitSetDiameterShape),
  },
  async (rawArgs) => {
    const args = ConduitSetDiameterSchema.parse(rawArgs);
    const result = await postRevit("mep.conduit.set_diameter", args);
    return asText(result);
  }
);

/* =========================================
   mep.cabletray.create
========================================= */
const CableTrayCreateShape = {
  level: z.string().optional(),
  cableTrayType: z.string().optional(),
  elevation_m: z.number().optional(), // por defecto en bridge: 2.7
  start: PtShape,
  end: PtShape,
};
const CableTrayCreateSchema = z.object(CableTrayCreateShape);

server.registerTool(
  "mep_cabletray_create",
  {
    title: "Create Cable Tray",
    description: "Crea una charola entre dos puntos XY (m). Tipo opcional.",
    inputSchema: asInputShape(CableTrayCreateShape),
  },
  async (rawArgs) => {
    const args = CableTrayCreateSchema.parse(rawArgs);
    const result = await postRevit("mep.cabletray.create", args);
    return asText(result);
  }
);

/* =========================================
   mep.cabletray.set_size
========================================= */
const CableTraySetSizeShape = {
  trayIds: z.array(z.number().int()).min(1),
  width_mm: z.number().optional(),
  height_mm: z.number().optional(),
};
const CableTraySetSizeSchema = z.object(CableTraySetSizeShape);

server.registerTool(
  "mep_cabletray_set_size",
  {
    title: "Set Cable Tray Size (Bulk)",
    description: "Ajusta ancho/alto (mm) de charolas.",
    inputSchema: asInputShape(CableTraySetSizeShape),
  },
  async (rawArgs) => {
    const args = CableTraySetSizeSchema.parse(rawArgs);
    const result = await postRevit("mep.cabletray.set_size", args);
    return asText(result);
  }
);

/* =========================================
   mep.cabletray.set_offset
========================================= */
const CableTraySetOffsetShape = {
  trayIds: z.array(z.number().int()).min(1),
  offset_m: z.number(),
};
const CableTraySetOffsetSchema = z.object(CableTraySetOffsetShape);

server.registerTool(
  "mep_cabletray_set_offset",
  {
    title: "Set Cable Tray Offset (Bulk)",
    description: "Ajusta el offset (m) de charolas.",
    inputSchema: asInputShape(CableTraySetOffsetShape),
  },
  async (rawArgs) => {
    const args = CableTraySetOffsetSchema.parse(rawArgs);
    const result = await postRevit("mep.cabletray.set_offset", args);
    return asText(result);
  }
);

// stdio
const transport = new StdioServerTransport();
await server.connect(transport);