// RE probe: how long does `OboThrottled` last (hypotheses §45)? It came on the
// tool-agent path only, on every account within 39 minutes on Oct 26, at lower
// agent-turn rates than earlier days without one, so our own turn counts don't
// predict it. This watches the agent path and times each episode.
//
// Every WATCH_S (120) seconds it sends one tool-less turn WITH the tool agent.
// When one comes back `Throttled` / `OboThrottled` it retries the agent path
// after FIRST_RETRY_S (30), then every RETRY_S (60) seconds, until a turn
// answers, and sends an agent-less turn at the start of the episode and every
// AGENTLESS_EVERY retries, to see whether the agent-less path is throttled too.
// Each episode's length is first OboThrottled → first answered agent turn.
//
// Usage: M365_NO_INTERACTIVE=1 M365_DEBUG=1 node scripts/obo-watch.mjs
//   env: HOURS (6)  EPISODES (3: stop after this many timed episodes)  TONE
//        (Gpt_5_5_Chat)  OUT (scripts/obo-watch-out/<timestamp>.jsonl)
// Cost: 30 turns an hour while watching, up to ~70 during an episode, each in a
// fresh conversation. It stops at the first `PerUserThrottled` turn. With
// M365_DEBUG=1 its turns go to debug-probes.log, so M365_AVOID_THROTTLING's
// bucket counts them (_probe-chat.mjs); run nothing else on the account.
import { appendFileSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { getToken, decodeJwt, getOrCreateAgent } from "../packages/core/dist/index.mjs";
import { oneTurn } from "./_probe-chat.mjs";

const num = (v, d) => (v !== undefined && Number.isFinite(Number(v)) ? Number(v) : d);
const HOURS = num(process.env.HOURS, 6);
const EPISODES = num(process.env.EPISODES, 3);
const WATCH_S = num(process.env.WATCH_S, 120);
const FIRST_RETRY_S = num(process.env.FIRST_RETRY_S, 30);
const RETRY_S = num(process.env.RETRY_S, 60);
const AGENTLESS_EVERY = num(process.env.AGENTLESS_EVERY, 5);
const TONE = process.env.TONE || "Gpt_5_5_Chat";
const OUT = process.env.OUT || join("scripts", "obo-watch-out", `${new Date().toISOString().replace(/[:.]/g, "-")}.jsonl`);
mkdirSync(dirname(OUT), { recursive: true });

const sleep = (s) => new Promise((r) => setTimeout(r, s * 1000));
const claims = decodeJwt(await getToken());
const agentId = await getOrCreateAgent();
if (!agentId) { console.error("[obo] no agent"); process.exit(1); }

/** One tool-less turn; `result` is the final item's, read from the frames. */
async function probe(withAgent) {
  let result = null;
  const t = Date.now();
  // A token lasts about an hour: get it per probe (getToken refreshes it silently).
  const token = await getToken();
  const r = await oneTurn({
    token, claims, agentId: withAgent ? agentId : null, tone: TONE, timeoutMs: 90000,
    text: "Reply with exactly the single word: pong",
    onFrame: (f) => {
      const it = f.type === 2 ? (Array.isArray(f.item) ? f.item[0] : f.item) : null;
      if (it?.result) result = { value: it.result.value, errorCode: it.result.errorCode ?? null };
    },
  });
  const row = { t: new Date(t).toISOString(), path: withAgent ? "agent" : "agent-less", value: result?.value ?? (r.error ? "error" : null),
    errorCode: result?.errorCode ?? null, ms: r.elapsedMs, reply: (r.fullText || "").trim().slice(0, 40), error: r.error };
  appendFileSync(OUT, JSON.stringify(row) + "\n");
  console.log(`[obo] ${row.t} ${row.path.padEnd(10)} ${String(row.value).padEnd(10)} ${row.errorCode ?? ""} ${row.ms}ms ${JSON.stringify(row.reply)}`);
  if (row.errorCode === "PerUserThrottled") { console.log("[obo] PerUserThrottled — stopping"); process.exit(3); }
  return row;
}
const obo = (row) => row.value === "Throttled" && row.errorCode === "OboThrottled";

const end = Date.now() + HOURS * 3600e3;
let episodes = 0;
console.log(`[obo] watching the agent path (${TONE}) every ${WATCH_S}s for ${HOURS}h → ${OUT}`);
while (Date.now() < end && episodes < EPISODES) {
  const first = await probe(true);
  if (obo(first)) {
    const since = Date.parse(first.t);
    await probe(false);
    let retries = 0;
    let wait = FIRST_RETRY_S;
    for (;;) {
      await sleep(wait);
      wait = RETRY_S;
      retries++;
      const r = await probe(true);
      if (retries % AGENTLESS_EVERY === 0) await probe(false);
      if (!obo(r)) {
        const s = Math.round((Date.parse(r.t) - since) / 1000);
        episodes++;
        const summary = { episode: episodes, since: first.t, answered: r.t, seconds: s, oboTurns: retries, endedWith: r.value };
        appendFileSync(OUT, JSON.stringify(summary) + "\n");
        console.log(`[obo] EPISODE ${episodes}: OboThrottled from ${first.t} until ${r.t} (${s}s, ${retries} retries; ended with ${r.value})`);
        break;
      }
    }
  }
  await sleep(WATCH_S);
}
console.log(`[obo] done: ${episodes} episode(s) timed`);
