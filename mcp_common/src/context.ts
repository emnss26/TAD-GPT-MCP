import { z } from "zod";

export const ContextOpts = z.object({
  cacheSec: z.number().int().min(1).max(3600).default(30),
});
type Ctx = z.infer<typeof ContextOpts>;

let _cache: { data: any; exp: number } | null = null;

export async function snapshotContext(
  postRevit: (action: string, args: any) => Promise<any>,
  opts?: Partial<Ctx>
) {
  const { cacheSec } = ContextOpts.parse(opts ?? {});
  const now = Date.now();
  if (_cache && _cache.exp > now) return _cache.data;

  const [activeView, levels, wallTypes, floorTypes, matList, famTypes, selection] =
    await Promise.all([
      postRevit("view.active", {}),
      postRevit("levels.list", {}),
      postRevit("walltypes.list", {}),
      postRevit("floortypes.list", {}).catch(() => ({ items: [] })), // si no existe, devuelve vacío
      postRevit("materials.list", {}),
      postRevit("families.types.list", {}),
      postRevit("selection.info", { includeParameters: false }).catch(() => ({ items: [] })),
    ]);

  const data = {
    activeView, levels,
    catalogs: {
      walls: wallTypes,
      floors: floorTypes,
      materials: matList,
      families: famTypes,
    },
    selection,
    ts: new Date().toISOString(),
    ttlSec: cacheSec,
  };

  _cache = { data, exp: now + cacheSec * 1000 };
  return data;
}