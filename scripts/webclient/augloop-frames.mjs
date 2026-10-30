// The AugLoop frames of one scripts/webclient/driver.mjs capture (run with WS_FRAMES='Chathub|augloop'): how an
// Office app's client talks to its agent, for docs/m365-copilot-api.md §13.
//
// Usage: node scripts/webclient/augloop-frames.mjs <OUT dir> [--from N] [--grep RE]   one line per frame: its
//          number, time, socket, direction, message type, annotation type or signal path, and the first body's keys
//        node scripts/webclient/augloop-frames.mjs <OUT dir> --show N [--width W]       frame N as JSON, strings
//          cut to W characters (default 400) and token fields left out
//        node scripts/webclient/augloop-frames.mjs <OUT dir> --flights [--grep RE]     the session's flights
// The capture names the account (document URLs, the user's name in the agent's context): keep it private.
import { readFileSync } from "node:fs";
import { join } from "node:path";

const [dir, ...args] = process.argv.slice(2);
if (!dir) { console.log("usage: augloop-frames.mjs <OUT dir> [--from N] [--grep RE] | --show N [--width W] | --flights [--grep RE]"); process.exit(2); }
const opt = (name) => { const i = args.indexOf(name); return i >= 0 ? args[i + 1] : undefined; };
const sockets = new Set(readFileSync(join(dir, "websockets.txt"), "utf8").split("\n").filter((l) => /augloop/i.test(l)).map((l) => l.split(" ")[1]));
const frames = readFileSync(join(dir, "ws-frames.ndjson"), "utf8").trim().split("\n").map((l, i) => ({ n: i + 1, ...JSON.parse(l) })).filter((f) => sockets.has(f.id));
const parse = (f) => { try { return JSON.parse(f.p); } catch { return null; } };

function cut(v, width) {
  if (typeof v === "string") return v.length > width ? `${v.slice(0, width)}…(${v.length})` : v;
  if (!v || typeof v !== "object") return v;
  if (Array.isArray(v)) return v.map((x) => cut(x, width));
  return Object.fromEntries(Object.entries(v).map(([k, x]) => [k, /token$/i.test(k) && typeof x === "string" && x.length > 40 ? "<token>" : cut(x, width)]));
}

if (opt("--show")) {
  const f = frames.find((x) => x.n === Number(opt("--show")));
  if (!f) { console.log("no such AugLoop frame"); process.exit(1); }
  console.log(`${f.t} ${f.id} ${f.dir}`);
  console.log(JSON.stringify(cut(parse(f) ?? f.p, Number(opt("--width") ?? 400)), null, 1));
} else if (args.includes("--flights")) {
  const init = frames.map(parse).find((m) => m?.H_?.T_ === "AugLoop_Session_Protocol_SessionInitMessage");
  const re = new RegExp(opt("--grep") ?? ".", "i");
  for (const fl of String(init?.clientMetadata?.flights ?? "").split(";").filter((x) => x && re.test(x))) console.log(fl);
} else {
  const from = Number(opt("--from") ?? 0);
  const re = opt("--grep") ? new RegExp(opt("--grep"), "i") : null;
  for (const f of frames.filter((x) => x.n >= from)) {
    const m = parse(f);
    if (!m) { if (!re && f.p !== "~") console.log(`${f.n} ${f.t.slice(11, 19)} ${f.id} ${f.dir} ${f.p.slice(0, 60)}`); continue; }
    const type = (m.H_?.T_ ?? "").replace("AugLoop_Session_Protocol_", "");
    if (type === "KeepAlive" || (type === "Response" && !m.error)) continue;
    const ops = m.ops ?? [];
    const what = m.annotationType ?? ops.map((o) => o.parentPath?.join("/")).filter(Boolean).join(" ") ?? "";
    const body = ops[0]?.items?.[0]?.body ?? m.content ?? null;
    const keys = body && typeof body === "object" ? Object.keys(body).filter((k) => k !== "H_").slice(0, 12).join(",") : "";
    const line = `${f.n} ${f.t.slice(11, 19)} ${f.id} ${f.dir} ${type} ${what}${m.error ? ` error=${m.code}:${String(m.error).slice(0, 80)}` : ""} ${f.p.length}b${keys ? ` {${keys}}` : ""}`;
    if (!re || re.test(line)) console.log(line);
  }
}
