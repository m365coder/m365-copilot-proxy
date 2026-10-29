// Compare two scripts/webclient/driver.mjs captures (two accounts): (1) the web client's own Chathub request,
// ids ignored; (2) the ECS configs it downloaded; (3) the flags, model picker and enabled model providers in the
// server-rendered page (save it in each capture with `return h.html("chat-page")`).
// Usage: node scripts/webclient/diff.mjs <OUT dir A> <OUT dir B>
import { existsSync, readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";

const [A, B] = process.argv.slice(2);
const RS = "\x1e";
const ID = /token|sessionid|requestid|conversationid|X-SessionId|XRoutingParameterSessionKey|clientCorrelationId|traceId|clientSessionId|chatsessionid|clientrequestid|^text$|createdAt|timestamp/i;
const MODEL = /opus|claude|anthropic|sonnet|gpt.?6|_sol|forum|council|critique|modelselect|modelpicker/i;

function diffMaps(label, a, b) {
  const keys = [...new Set([...a.keys(), ...b.keys()])].sort();
  let n = 0;
  for (const k of keys) {
    const va = a.get(k) ?? "(absent)", vb = b.get(k) ?? "(absent)";
    if (va !== vb) { n++; console.log(`  ${k}\n      A=${String(va).slice(0, 300)}\n      B=${String(vb).slice(0, 300)}`); }
  }
  console.log(`  ${label}: ${n} of ${keys.length} differ`);
}
const flatten = (o, p = "", out = new Map()) => {
  if (o && typeof o === "object" && !Array.isArray(o)) { for (const [k, v] of Object.entries(o)) if (!ID.test(k)) flatten(v, p ? `${p}.${k}` : k, out); }
  else out.set(p, JSON.stringify(o));
  return out;
};

// (1) the client's own request: prefer the routed log (the frame before any tone rewrite).
function request(dir) {
  const routed = join(dir, "ws-routed.ndjson");
  let url, frame;
  if (existsSync(routed)) {
    const lines = readFileSync(routed, "utf8").trim().split("\n").map((l) => JSON.parse(l)).reverse();
    url = lines.find((l) => l.dir === "url")?.p;
    frame = lines.find((l) => l.dir === "sent-original")?.p;
  } else if (existsSync(join(dir, "ws-frames.ndjson"))) {
    url = readFileSync(join(dir, "websockets.txt"), "utf8").trim().split("\n").reverse().find((l) => /Chathub/i.test(l))?.split(" ")[2];
    frame = readFileSync(join(dir, "ws-frames.ndjson"), "utf8").trim().split("\n").map((l) => JSON.parse(l)).reverse().find((l) => l.dir === "sent" && l.p.includes('"target":"chat"'))?.p;
  }
  if (!url || !frame) return null;
  const params = new Map();
  for (const [k, v] of new URL(url.replace(/^wss/, "https")).searchParams) if (!ID.test(k)) params.set(k, v);
  let args = null;
  for (const f of frame.split(RS)) { try { const j = JSON.parse(f); if (j.type === 4 && j.target === "chat") args = j.arguments[0]; } catch {} }
  return { params, fields: flatten(args ?? {}) };
}
console.log("== (1) the web client's own Chathub request");
const ra = request(A), rb = request(B);
if (ra && rb) {
  const va = (ra.params.get("variants") ?? "").split(","), vb = (rb.params.get("variants") ?? "").split(",");
  ra.params.delete("variants"); rb.params.delete("variants");
  diffMaps("URL parameters other than variants", ra.params, rb.params);
  console.log(`  variants: A ${va.length}, B ${vb.length}\n    only A: ${va.filter((v) => !vb.includes(v)).join(", ") || "-"}\n    only B: ${vb.filter((v) => !va.includes(v)).join(", ") || "-"}`);
  diffMaps("chat frame fields", ra.fields, rb.fields);
} else console.log("  (no chat request in one of the captures)");

// (2) ECS configs.
function ecs(dir) {
  const out = new Map();
  const bodies = join(dir, "bodies");
  if (!existsSync(bodies)) return out;
  for (const f of readdirSync(bodies).sort()) {
    const s = readFileSync(join(bodies, f), "utf8");
    const url = s.split("\n")[0];
    if (!/config\.edge\.skype\.com|ecs\.office\.com|config\/v1\//i.test(url)) continue;
    const name = url.match(/PartnerId=([^&]+)/)?.[1] ?? url.match(/config\/v1\/([^/?]+)/)?.[1] ?? f;
    try { flatten(JSON.parse(s.slice(s.indexOf("\n\n") + 2)), name, out); } catch {}
  }
  for (const k of [...out.keys()]) if (/Headers\.(Expires|ETag)|\.ETag$/i.test(k)) out.delete(k);
  return out;
}
console.log("\n== (2) ECS configs");
diffMaps("ECS keys", ecs(A), ecs(B));

// (3) the server-rendered page.
function page(dir) {
  const f = join(dir, "chat-page.html");
  if (!existsSync(f)) return null;
  const s = readFileSync(f, "utf8").replace(/\\u0022/g, '"').replace(/\\"/g, '"');
  const flags = new Map();
  for (const m of s.matchAll(/"([A-Za-z_][A-Za-z0-9_.\-]{2,100})":(true|false)\b/g)) {
    const prev = flags.get(m[1]);
    flags.set(m[1], prev === undefined || prev === m[2] ? m[2] : `${prev}|${m[2]}`);
  }
  const providers = [...new Set([...s.matchAll(/"enabledModels":\[([^\]]*)\]/g)].map((m) => [...m[1].matchAll(/"displayName":"([^"]+)"/g)].map((x) => x[1]).join(", ")))];
  const options = new Set([...s.matchAll(/"id":"([^"]+)","type":"item","menuItemTitle":"([^"]+)"/g)].map((m) => `${m[2]} (${m[1]})`));
  return { flags, providers, options };
}
console.log("\n== (3) the server-rendered page");
const pa = page(A), pb = page(B);
if (pa && pb) {
  console.log(`  model picker\n      A=${[...pa.options].sort().join("; ")}\n      B=${[...pb.options].sort().join("; ")}`);
  console.log(`  enabled model providers\n      A=${pa.providers.join(" / ")}\n      B=${pb.providers.join(" / ")}`);
  const model = (m) => new Map([...m].filter(([k]) => MODEL.test(k)));
  console.log("  model-related flags:");
  diffMaps("model-related flags", model(pa.flags), model(pb.flags));
  for (const [k, v] of [...model(pa.flags)].sort()) console.log(`      ${k} = ${v}`);
  console.log("  all boolean flags:");
  diffMaps("boolean flags", pa.flags, pb.flags);
} else console.log("  (no chat-page.html in one of the captures)");
