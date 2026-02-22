export function applyTemplate(obj, scope) {
    if (obj == null)
        return obj;
    if (typeof obj === "string") {
        return obj.replace(/\{\$(.+?)\}/g, (_, expr) => {
            const parts = expr.split("."); // e.g. vars.level, item.start.x
            let cur = scope;
            for (const p of parts) {
                if (cur == null)
                    return "";
                cur = cur[p];
            }
            return cur ?? "";
        });
    }
    if (Array.isArray(obj))
        return obj.map(v => applyTemplate(v, scope));
    if (typeof obj === "object") {
        const out = {};
        for (const k of Object.keys(obj))
            out[k] = applyTemplate(obj[k], scope);
        return out;
    }
    return obj;
}
