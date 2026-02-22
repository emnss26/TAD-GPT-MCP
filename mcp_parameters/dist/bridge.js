import "dotenv/config";
const BASE_URL = process.env.REVIT_BRIDGE_URL ?? "http://127.0.0.1:55234/mcp";
export async function postRevit(action, args) {
    const res = await fetch(BASE_URL, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ action, args }),
    });
    const data = (await res.json());
    if (!res.ok || !data?.ok) {
        const msg = data?.message ?? res.statusText;
        throw new Error(msg);
    }
    return data.data ?? undefined;
}
