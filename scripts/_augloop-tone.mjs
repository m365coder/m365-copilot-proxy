// Shared runner for the AugLoop tone probes (augloop-excel-tone-probe.mjs, augloop-word-tone-probe.mjs): one
// new conversation and one turn per cell, through the core's conversation for the app's agent (the proxy's
// own path: Chat only, the configured sensitivity label, and temporary chat unless M365_SAVE_HISTORY=1).
// Each probe adds a made-up control cell, so that a cell can be read against what the server does with a
// model it doesn't have. Results go to scripts/tone-out/augloop-<app>-<time>/results.json.
import { mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { AugLoopConversation, WordAgentConversation, deleteAgentChats, getAugLoopToken, getGraphToken, getToken } from "../packages/core/dist/index.mjs";

export const PROMPT = process.env.PROMPT || "In one sentence: which AI model are you, and which company made you?";

/** One turn of PROMPT with `model` (a proxy model ID, raw forms included) in a new conversation. */
export async function runCell(app, model) {
  const temporary = process.env.M365_SAVE_HISTORY !== "1";
  const opts = {
    getToken: getAugLoopToken,
    getGraphToken,
    ...(temporary ? { deleteChats: async (ids) => deleteAgentChats(await getToken(), ids, app) } : {}),
  };
  const conv = app === "word" ? new WordAgentConversation(opts) : new AugLoopConversation(opts);
  const row = { model, result: null, errorCode: null, error: null, origin: null, adjusted: null, text: "", ms: 0 };
  const t0 = Date.now();
  try {
    const stream = await conv.chat(PROMPT, model);
    for await (const _ of stream) { /* drain */ }
    row.result = stream.result?.value ?? null;
    row.errorCode = stream.result?.errorCode ?? null;
    row.origin = stream.contentOrigin;
    row.text = stream.fullText.slice(0, 300);
    if (app === "word") row.adjusted = conv.lastModelAdjustment;
  } catch (err) {
    row.error = String(err?.message ?? err).slice(0, 300);
  } finally {
    row.ms = Date.now() - t0;
    conv.close();
    await conv.historySettled();
  }
  return row;
}

/** Write the rows; returns the directory. */
export function writeResults(app, rows) {
  const out = join(process.cwd(), "scripts", "tone-out", `augloop-${app}-${new Date().toISOString().replace(/[:.]/g, "-")}`);
  mkdirSync(out, { recursive: true });
  writeFileSync(join(out, "results.json"), JSON.stringify({ prompt: PROMPT, rows }, null, 2));
  return out;
}

/** The row's outcome in a few words: the result, or the error. */
export function outcomeOf(row) {
  if (row.error) return `error: ${row.error.slice(0, 120)}`;
  return `${row.result ?? "no result"}${row.errorCode ? ` (${row.errorCode})` : ""}`;
}
