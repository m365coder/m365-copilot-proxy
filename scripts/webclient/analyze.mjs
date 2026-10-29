// Summarise one scripts/webclient/driver.mjs capture: the Chathub URL parameters the web client sent, its
// chat frames (the client's own and, under h.routeTone, the one that left), each turn's final result, and the
// saved bodies that mention models, Claude, Opus, Anthropic, tones or flights.
// Usage: node scripts/webclient/analyze.mjs <OUT dir> [--bodies]
import { existsSync, readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";

const dir = process.argv[2];
const showBodies = process.argv.includes("--bodies");
const RS = "\x1e";
const ID_PARAMS = /token|sessionid|requestid|conversationid|X-SessionId/i;

function params(url) {
  const u = new URL(url.replace(/^wss/, "https"));
  console.log(`  ${u.pathname.replace(/[0-9a-f-]{36}@[0-9a-f-]{36}/, "<oid>@<tid>")}`);
  for (const [k, v] of u.searchParams) if (!ID_PARAMS.test(k)) console.log(`    ${k} = ${v.slice(0, 600)}`);
}

function frames(file, label) {
  if (!existsSync(file)) return;
  console.log(`\n== ${label}`);
  for (const line of readFileSync(file, "utf8").trim().split("\n")) {
    const { t, id, dir: d, p } = JSON.parse(line);
    if (d === "url") { console.log(`${t} ${id} connection`); params(p); continue; }
    for (const f of String(p).split(RS)) {
      let j; try { j = JSON.parse(f); } catch { continue; }
      if (d.startsWith("sent") && j.type === 4 && j.target === "chat") {
        const a = j.arguments?.[0] ?? {};
        const { message, ...rest } = a;
        console.log(`${t} ${id} ${d === "sent-original" ? "CLIENT'S OWN FRAME" : "SENT"} tone=${a.tone} text=${JSON.stringify(message?.text ?? "").slice(0, 100)}`);
        if (d !== "sent" || !file.endsWith("ws-routed.ndjson")) {
          console.log(`    ${JSON.stringify(rest).slice(0, 2500)}`);
          if (message) { const { text, ...m } = message; console.log(`    message: ${JSON.stringify(m).slice(0, 1200)}`); }
        }
      }
      if (d === "recv" && j.type === 2) {
        const it = Array.isArray(j.item) ? j.item[0] : j.item;
        const bots = (it?.messages ?? []).filter((m) => m.author === "bot");
        console.log(`${t} ${id} FINAL result=${JSON.stringify(it?.result ?? {}).slice(0, 300)} turnState=${it?.turnState} origins=${bots.map((m) => m.contentOrigin ?? m.messageType).join(",")}`);
        for (const m of bots) if (m.text && !m.messageType) console.log(`    bot (${m.contentOrigin}): ${JSON.stringify(m.text).slice(0, 300)}`);
      }
    }
  }
}

if (existsSync(join(dir, "websockets.txt"))) {
  console.log("== Chathub connections");
  for (const line of readFileSync(join(dir, "websockets.txt"), "utf8").trim().split("\n")) {
    const [t, id, url] = line.split(" ");
    if (/Chathub/i.test(url)) { console.log(`${t} ${id}`); params(url); }
  }
}
frames(join(dir, "ws-frames.ndjson"), "chat frames (as the page saw them)");
frames(join(dir, "ws-routed.ndjson"), "routed chat frames (h.routeTone)");

const bodies = join(dir, "bodies");
if (existsSync(bodies)) {
  console.log("\n== bodies mentioning models / Claude / Opus / Anthropic / tones / flights");
  const RE = /opus|claude|anthropic|"tone|modelSelector|availableModelSelectionOptions|enabledModels|Gpt_5|Claude_|flight/i;
  for (const f of readdirSync(bodies).sort()) {
    const s = readFileSync(join(bodies, f), "utf8");
    const hits = (s.match(new RegExp(RE.source, "gi")) ?? []).length;
    if (!hits) continue;
    console.log(`${f} (${hits} hits, ${s.length} chars) ${s.split("\n")[0].slice(0, 200)}`);
    if (!showBodies) continue;
    const seen = new Set();
    for (const m of s.matchAll(new RegExp(`.{0,120}(${RE.source}).{0,160}`, "gi"))) {
      const x = m[0].replace(/\s+/g, " ");
      if (!seen.has(x)) { seen.add(x); console.log(`    …${x}…`); }
      if (seen.size >= 25) break;
    }
  }
}
