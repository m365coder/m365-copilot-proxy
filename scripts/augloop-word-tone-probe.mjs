// RE probe: which model IDs Word's agent takes as `selectedModel` (docs/hypotheses.md §49). One turn per ID,
// through `augloop-word:<modelId>[:<reasoningEffort>]` on the proxy's own path (Chat only), plus a made-up
// control ID. Word's model contract doesn't check IDs (F131), and nothing on the client names the model that
// answered (F133), so read each row against the control:
//   refused          → an error or a failed turn, where the control's differs from it or the control is refused too
//   adjusted         → the server's `ModelSelectionAdjusted` annotation: it changed the model
//   taken            → a normal answer; only evidence if the control was refused or adjusted
// The answer's self-ID says at most the family.
//
// Usage: node scripts/augloop-word-tone-probe.mjs [MODELID[:LEVEL] …]
//   default: the IDs the proxy's augloop-word-* models send. LEVEL is a GPT reasoning level (`GPT56:low`).
//   env: PROMPT, M365_SAVE_HISTORY=1 (keep the chats), M365_DEBUG=1, M365_DUMP_FRAMES=1
// Cost: one Word turn per ID and one for the control. Build first (`pnpm build`).
import { AUGLOOP_MODEL_IDS, wordAgentModelOf } from "../packages/core/dist/index.mjs";
import { outcomeOf, runCell, writeResults } from "./_augloop-tone.mjs";

const CONTROL = "NotAModel123";
const args = process.argv.slice(2);
const ids = args.length
  ? args
  : [...new Set(AUGLOOP_MODEL_IDS.map(wordAgentModelOf).filter(Boolean).map((m) => (m.reasoningEffort ? `${m.modelId}:${m.reasoningEffort}` : m.modelId)))];

const rows = [];
for (const id of [...ids, CONTROL]) {
  const row = { id, ...(await runCell("word", `augloop-word:${id}`)) };
  rows.push(row);
  console.log(`[tone] ${id.padEnd(30)} ${outcomeOf(row).padEnd(24)} ${row.ms}ms${row.adjusted ? ` ADJUSTED=${JSON.stringify(row.adjusted).slice(0, 160)}` : ""} ${JSON.stringify(row.text.slice(0, 90))}`);
}

const control = rows.at(-1);
const taken = (r) => !r.error && r.result === "Success" && !r.adjusted;
console.log("\n[tone] === VERDICT ===");
console.log(`[tone] control ${CONTROL}: ${control.adjusted ? "adjusted" : outcomeOf(control)}${taken(control) ? " — the server takes any ID, so a taken ID proves nothing by itself" : ""}`);
for (const r of rows.slice(0, -1)) {
  const v = r.adjusted ? `adjusted by the server: ${JSON.stringify(r.adjusted).slice(0, 200)}`
    : !taken(r) ? `refused: ${outcomeOf(r)}`
    : taken(control) ? "taken, as the control was: not evidence on its own"
    : "taken, where the control wasn't";
  console.log(`  ${r.id}: ${v}`);
}
console.log(`[tone] out: ${writeResults("word", rows)}`);
process.exit(0);
