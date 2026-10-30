// RE probe: which Excel agent slots (`agentModelType`) serve which model (docs/hypotheses.md §48 F120). One
// turn per slot, through `augloop-excel:<slot>` on the proxy's own path (Chat only), plus a made-up control
// slot. A slot the proxy knows gets the flights Excel's config names for it (its model and metering units);
// any other gets none, so the server serves its own default for it. The served model is read from the agent
// state the server asks the client to save (F121), not from the answer. A slot that serves what the control
// serves is a default, not a model of its own.
//
// Usage: node scripts/augloop-excel-tone-probe.mjs [SLOT …]
//   default: the slots the proxy's augloop-excel-* models use
//   env: PROMPT, M365_SAVE_HISTORY=1 (keep the chats), M365_DEBUG=1, M365_DUMP_FRAMES=1
// Cost: one Excel turn per slot and one for the control; a Claude Opus slot's turn counts toward its daily
// limit (Opus 5.5: 40 a day, F124). Build first (`pnpm build`).
import { AUGLOOP_MODEL_IDS, augLoopModelOf } from "../packages/core/dist/index.mjs";
import { outcomeOf, runCell, writeResults } from "./_augloop-tone.mjs";

const CONTROL = "NotASlot1";
const args = process.argv.slice(2);
const slots = args.length ? args : [...new Set(AUGLOOP_MODEL_IDS.map(augLoopModelOf).filter(Boolean).map((m) => m.modelType))];

const rows = [];
for (const slot of [...slots, CONTROL]) {
  const row = { slot, ...(await runCell("excel", `augloop-excel:${slot}`)) };
  // The origin is `AugLoop/<models the saved state names>`, or `AugLoop` when no state named one.
  row.served = row.origin?.startsWith("AugLoop/") ? row.origin.slice("AugLoop/".length) : null;
  rows.push(row);
  console.log(`[tone] ${slot.padEnd(22)} ${outcomeOf(row).padEnd(24)} served=${row.served ?? "?"} ${row.ms}ms ${JSON.stringify(row.text.slice(0, 80))}`);
}

const control = rows.at(-1);
console.log("\n[tone] === VERDICT ===");
console.log(`[tone] control ${CONTROL}: ${control.served ? `served ${control.served}, the server's default for a slot without a model` : outcomeOf(control)}`);
for (const r of rows.slice(0, -1)) {
  const v = r.error ? `error: ${r.error.slice(0, 120)}`
    : !r.served ? `no served model in the saved state (${outcomeOf(r)})`
    : control.served && r.served === control.served ? `served ${r.served}, as the control did: a default`
    : `serves ${r.served}`;
  console.log(`  ${r.slot}: ${v}`);
}
console.log(`[tone] out: ${writeResults("excel", rows)}`);
process.exit(0);
