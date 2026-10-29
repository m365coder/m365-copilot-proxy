// One chat turn on one tone with every frame kept, to compare what two accounts get for the same request
// (a live one and a dead one, §47). Prints the final item without its messages; the frames go to
// ~/.cache/tone-frames/<time>-<LABEL>.ndjson. Costs one turn.
// Env: TONE (default Claude_Opus), SCENARIO (default OfficeWebIncludedCopilot), LICENSE, PROMPT, LABEL,
//      VARIANTS (replaces the WS `variants` list; a file's contents: VARIANTS="$(cat file)"),
//      OPTIONS_SETS (comma list), SAVE_HISTORY=1 (no disableMemory), NO_PLUGINS=1.
// Usage: M365_DEBUG=1 node scripts/tone-frames-probe.mjs
import { mkdirSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { getToken, decodeJwt } from "../packages/core/dist/index.mjs";
import { oneTurn } from "./_probe-chat.mjs";

const LABEL = process.env.LABEL ?? "probe";
const tone = process.env.TONE ?? "Claude_Opus";
const scenario = process.env.SCENARIO ?? "OfficeWebIncludedCopilot";
const licenseType = process.env.LICENSE ?? (scenario === "OfficeWebPaidCopilot" ? "Premium" : "Starter");
const text = process.env.PROMPT ??
  'Reply with only your model name and version, a semicolon, the full company that made you, a semicolon, and your knowledge cutoff date (YYYY-mm) or "unknown".';

const token = await getToken();
const frames = [];
const opts = { token, claims: decodeJwt(token), agentId: null, tone, scenario, licenseType, text, timeoutMs: 90000, onFrame: (f) => frames.push(f) };
if (process.env.VARIANTS !== undefined) opts.variants = process.env.VARIANTS;
if (process.env.OPTIONS_SETS) opts.optionsSets = process.env.OPTIONS_SETS.split(",").filter(Boolean);
if (process.env.SAVE_HISTORY === "1") opts.temporaryChat = false;
if (process.env.NO_PLUGINS === "1") opts.plugins = [];

const r = await oneTurn(opts);
const dir = join(homedir(), ".cache", "tone-frames");
mkdirSync(dir, { recursive: true });
const file = join(dir, `${new Date().toISOString().replace(/[:.]/g, "-")}-${LABEL}.ndjson`);
writeFileSync(file, frames.map((f) => JSON.stringify(f)).join("\n") + "\n");

const final = frames.find((f) => f.type === 2);
const item = final ? (Array.isArray(final.item) ? final.item[0] : final.item) : null;
const { messages, ...finalItem } = item ?? {};
console.log(JSON.stringify({
  label: LABEL, tone, scenario, licenseType,
  variants: opts.variants ? opts.variants.split(",").filter(Boolean).length : "default",
  contentOrigin: r.contentOrigin, reply: r.fullText.slice(0, 300), elapsedMs: r.elapsedMs, frameCount: r.frameCount, error: r.error,
  finalItem, file,
}, null, 1));
process.exit(0);
