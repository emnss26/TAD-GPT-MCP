export function applyTemplate<T = any>(obj: T, scope: any): T {
  if (obj == null) return obj;
  if (typeof obj === "string") {
    return obj.replace(/\{\$(.+?)\}/g, (_, expr) => {
      const parts = expr.split("."); // e.g. vars.level, item.start.x
      let cur: any = scope;
      for (const p of parts) { if (cur == null) return ""; cur = cur[p]; }
      return cur ?? "";
    }) as any;
  }
  if (Array.isArray(obj)) return obj.map(v => applyTemplate(v, scope)) as any;
  if (typeof obj === "object") {
    const out: any = {};
    for (const k of Object.keys(obj as any)) out[k] = applyTemplate((obj as any)[k], scope);
    return out;
  }
  return obj;
}