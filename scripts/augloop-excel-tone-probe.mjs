// RE probe: which Excel agent slots (`agentModelType`) serve which model (docs/hypotheses.md §48 F120). One
// turn per slot, through `augloop-excel:<slot>` on the proxy's own path (Chat only), the made-up control slot
// last. A slot the proxy knows gets the flights Excel's config names for it (its model and metering units);
// any other gets none, so the server serves its own default for it. The served model is read from the agent
// state the server asks the client to save (F121), not from the answer. A slot that serves what the control
// serves is a default, not a model of its own.
//
// Usage: node scripts/augloop-excel-tone-probe.mjs [SLOT …]
//   without arguments, the SLOTS below; with them, those slots and the control
//   env: PROMPT (default below), M365_SAVE_HISTORY=1 (keep the chats), M365_DEBUG=1, M365_DUMP_FRAMES=1
// Output: the terminal, and scripts/augloop-excel-tone-out/<time>/ (git-ignored): results.json, each slot's
// AugLoop frames (`NN-<slot>.frames.jsonl`) and the agent state the served model is read from
// (`NN-<slot>.state.zip`, when the server sent one). Keep it private: the state carries the user's directory
// details.
// Cost: one Excel turn per slot and one for the control; a Claude Opus slot's turn counts toward its daily
// limit (Opus 5.5: 40 a day, F124). Build first (`pnpm build`).
import { outDir, outcomeOf, runCell, slugOf, writeResults } from "./_augloop-tone.mjs";

const PROMPT = process.env.PROMPT || "In one sentence: which AI model are you, and which company made you?";

/**
 * Excel agent slots (2026-10-30). The made-up control is last.
 * @type {{slot: string, note: string}[]}
 */
const SLOTS = [
  // The proxy's augloop-excel-* models: each slot gets the flights Excel for the web's config names for it.
  { slot: "ClaudeSlot6", note: "Opus 5.5 (augloop-excel-claude-opus-5.5): counts toward its 40 a day (F124)" },
  { slot: "ClaudeSlot7", note: "Sonnet 5.5 (augloop-excel-claude-sonnet-5.5)" },
  { slot: "ClaudeSlot4", note: "Opus 5 (augloop-excel-claude-opus-5)" },
  { slot: "Gpt5Slot4", note: "GPT-6 Sol (augloop-excel-gpt-6-sol)" },
  { slot: "Gpt5Slot6", note: "GPT-6.1 Sol (augloop-excel-gpt-6.1-sol)" },

  // A slot the proxy has no flights for: the server's default for it.
  { slot: "ClaudeSonnetSlot5", note: "the reference project's Sonnet slot: served GPT-5.6 Sol without its flights (F120)" },

  { slot: "NotASlot1", note: "CONTROL: made-up" },
];

const args = process.argv.slice(2);
const cells = args.length ? [...args.map((slot) => ({ slot, note: "" })), SLOTS.at(-1)] : SLOTS;

const OUT = outDir("excel");
const rows = [];
for (const [i, cell] of cells.entries()) {
  const row = { ...cell, ...(await runCell("excel", `augloop-excel:${cell.slot}`, PROMPT, OUT, slugOf(i, cell.slot))) };
  // The origin is `AugLoop/<models the saved state names>`, or `AugLoop` when no state named one.
  row.served = row.origin?.startsWith("AugLoop/") ? row.origin.slice("AugLoop/".length) : null;
  rows.push(row);
  console.log(`[tone] ${cell.slot.padEnd(22)} ${outcomeOf(row).padEnd(24)} served=${row.served ?? "?"} ${row.ms}ms ${JSON.stringify(row.text.slice(0, 80))}`);
}

const control = rows.at(-1);
console.log("\n[tone] === VERDICT ===");
console.log(`[tone] control ${control.slot}: ${control.served ? `served ${control.served}, the server's default for a slot without a model` : outcomeOf(control)}`);
for (const r of rows.slice(0, -1)) {
  const v = r.error ? `error: ${r.error.slice(0, 120)}`
    : !r.served ? `no served model in the saved state (${outcomeOf(r)})`
    : control.served && r.served === control.served ? `served ${r.served}, as the control did: a default`
    : `serves ${r.served}`;
  console.log(`  ${r.slot}: ${v}${r.note ? ` — ${r.note}` : ""}`);
}
writeResults(OUT, PROMPT, rows);
console.log(`[tone] out: ${OUT}`);
process.exit(0);
