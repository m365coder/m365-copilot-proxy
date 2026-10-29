// Which backend region serves this account? Opens handshake-only Chathub connections (no chat message, so
// no turn and no thread) and prints `x-bepartition` / `x-calculatedbetarget` from each upgrade response:
// first with a new ConversationId each, then each with its own `X-RoutingParameter-SessionKey`.
// On 2026-10-29 every account always landed in one region; the key chose the machine within it (§47 F114).
// Usage: node scripts/region-probe.mjs [N]   (N connections per kind, default 4)
import { pathToFileURL } from "node:url";
import { getToken, decodeJwt } from "../packages/core/dist/index.mjs";

const ROOT = pathToFileURL(process.cwd()).href;
const wsMod = await import(`${ROOT}/node_modules/.pnpm/ws@8.20.0/node_modules/ws/wrapper.mjs`);
const WebSocket = wsMod.default ?? wsMod.WebSocket;
const N = Number(process.argv[2] ?? 4);

const token = await getToken();
const { oid, tid } = decodeJwt(token);

function connect(sessionKey) {
  return new Promise((resolve) => {
    const id = crypto.randomUUID();
    const params = new URLSearchParams({
      chatsessionid: id, clientrequestid: id, "X-SessionId": crypto.randomUUID(), ConversationId: crypto.randomUUID(),
      access_token: token, source: '"officeweb"', product: "Office", agentHost: "Bizchat.FullScreen",
      licenseType: "Starter", agent: "web", scenario: "OfficeWebIncludedCopilot",
    });
    const headers = { Origin: "https://m365.cloud.microsoft", "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:148.0) Gecko/20100101 Firefox/148.0" };
    if (sessionKey) headers["X-RoutingParameter-SessionKey"] = sessionKey;
    const ws = new WebSocket(`wss://substrate.office.com/m365Copilot/Chathub/${oid}@${tid}?${params}`, { headers });
    let target = "?";
    ws.on("upgrade", (res) => { target = `${res.headers["x-bepartition"]} ${String(res.headers["x-calculatedbetarget"]).replace(".SUBSTRATE.COSMIC.OFFICE.NET", "")}`; });
    ws.on("open", () => ws.close());
    ws.on("error", (e) => resolve(`error ${e.message}`));
    ws.on("close", () => resolve(target));
  });
}

const regions = new Map();
for (const [label, key] of [["new ConversationId", false], ["own session key", true]]) {
  for (let i = 0; i < N; i++) {
    const t = await connect(key ? crypto.randomUUID() : null);
    console.log(`${label} ${i + 1}: ${t}`);
    const region = t.split(" ")[0];
    regions.set(region, (regions.get(region) ?? 0) + 1);
  }
}
console.log(`regions: ${[...regions].map(([r, n]) => `${r} ×${n}`).join(", ")}`);
process.exit(0);
