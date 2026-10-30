// RE probe: which model IDs Word's agent takes as `selectedModel` (docs/hypotheses.md §49). One turn per ID,
// through `augloop-word:<modelId>[:<reasoningEffort>]` on the proxy's own path (Chat only), the made-up
// control last. Word's model contract doesn't check IDs (F131), and nothing on the client names the model
// that answered (F133), so read each row against the control:
//   refused          → an error or a failed turn
//   adjusted         → the server's `ModelSelectionAdjusted` annotation: it changed the model
//   taken            → a normal answer; only evidence if the control was refused or adjusted
// The control is taken too (F137): an unknown ID gets an OpenAI model without a word. So read the answer: the
// self-ID says the family, and the Claude models write a straight apostrophe ("I'm"), the OpenAI ones a
// typographic one ("I’m").
//
// Usage: node scripts/augloop-word-tone-probe.mjs [MODELID[:LEVEL] …]
//   without arguments, the IDS below; with them, those IDs and the control. LEVEL is a reasoning level.
//   env: PROMPT (default below), M365_SAVE_HISTORY=1 (keep the chats), M365_DEBUG=1, M365_DUMP_FRAMES=1
// Cost: one Word turn per ID and one for the control. Build first (`pnpm build`).
import { outcomeOf, runCell, writeResults } from "./_augloop-tone.mjs";

const PROMPT = process.env.PROMPT || "In one sentence: which AI model are you, and which company made you?";

/**
 * Model IDs, as Word's model table names them (`omniAgent.modelAlias`, and `reasoningEffort` after a
 * colon), 2026-10-30. The made-up control is last.
 * @type {{id: string, note: string}[]}
 */
const IDS = [
  // What Word's client sends for this account: the proxy's augloop-word-* models (F131).
  { id: "anthropic-claude-sonnet-5-5", note: "Sonnet 5.5 (augloop-word-claude-sonnet-5.5): answered as Claude (F137)" },
  { id: "Opus55", note: "Opus 5.5 (augloop-word-claude-opus-5.5): answered as Claude (F137)" },
  { id: "GPT56:none", note: "GPT-5.6 Sol Quick Response (augloop-word-gpt-5.6-sol-quick)" },
  { id: "GPT56:low", note: "GPT-5.6 Sol Think Deeper (augloop-word-gpt-5.6-sol-think-deeper)" },
  { id: "GPT6-sol:low", note: "GPT-6.1 Sol (augloop-word-gpt-6.1-sol)" },

  // Other aliases in Word's model table, which its client sends under settings this account's config
  // doesn't turn on.
  { id: "Sonnet55", note: "Sonnet 5.5 without EnableWordModelSonnet55Omni: answered as the control did (F137)" },
  { id: "gpt61sol:low", note: "GPT-6.1 Sol with EnableWordOmniModelGpt61Sol" },
  { id: "claude-fable-5", note: "Claude Fable 5, behind EnableWordModelFable" },
  { id: "gpt6astra:low", note: "GPT-6 Astra, behind EnableWordModelGpt6Astra and EnableWordOmniModelGpt6Astra" },
  { id: "xai-grok-4.7:low", note: "Grok 4.7, behind EnableWordModelXai and EnableWordOmniModelXai" },

  { id: "NotAModel123", note: "CONTROL: made-up" },
];

const args = process.argv.slice(2);
const cells = args.length ? [...args.map((id) => ({ id, note: "" })), IDS.at(-1)] : IDS;

const rows = [];
for (const cell of cells) {
  const row = { ...cell, ...(await runCell("word", `augloop-word:${cell.id}`, PROMPT)) };
  rows.push(row);
  console.log(`[tone] ${cell.id.padEnd(30)} ${outcomeOf(row).padEnd(24)} ${row.ms}ms${row.adjusted ? ` ADJUSTED=${JSON.stringify(row.adjusted).slice(0, 160)}` : ""} ${JSON.stringify(row.text.slice(0, 90))}`);
}

const control = rows.at(-1);
const taken = (r) => !r.error && r.result === "Success" && !r.adjusted;
console.log("\n[tone] === VERDICT ===");
console.log(`[tone] control ${control.id}: ${control.adjusted ? "adjusted" : outcomeOf(control)}${taken(control) ? " — the server takes any ID, so a taken ID proves nothing by itself" : ""}`);
for (const r of rows.slice(0, -1)) {
  const v = r.adjusted ? `adjusted by the server: ${JSON.stringify(r.adjusted).slice(0, 200)}`
    : !taken(r) ? `refused: ${outcomeOf(r)}`
    : taken(control) ? "taken, as the control was: not evidence on its own"
    : "taken, where the control wasn't";
  console.log(`  ${r.id}: ${v}${r.note ? ` — ${r.note}` : ""}`);
}
console.log(`[tone] out: ${writeResults("word", PROMPT, rows)}`);
process.exit(0);
