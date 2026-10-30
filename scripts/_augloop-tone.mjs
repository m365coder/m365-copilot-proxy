// Shared runner for the AugLoop tone probes (augloop-excel-tone-probe.mjs, augloop-word-tone-probe.mjs): one
// new conversation and one turn per cell, through the core's conversation for the app's agent (the proxy's
// own path: Chat only, the configured sensitivity label, and temporary chat unless M365_SAVE_HISTORY=1).
// Each probe has its own list and prompt, its made-up control last, so that a cell can be read against what
// the server does with a model it doesn't have.
//
// A run's output goes to scripts/augloop-<app>-tone-out/<time>/ (git-ignored): results.json, and per cell the
// files the probe downloaded and read: the cell's AugLoop frames, `NN-<cell>.frames.jsonl`, redacted as the
// core's frame dumps are (no tokens, no saved state), and for Excel the agent's saved state,
// `NN-<cell>.state.zip`, when the server sent one. They name the account (the saved state carries the user's
// directory details): keep them private.
import { mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { AugLoopConversation, WordAgentConversation, deleteAgentChats, getAugLoopToken, getGraphToken, getToken } from "../packages/core/dist/index.mjs";

/** A new output directory for one run: scripts/augloop-<app>-tone-out/<time>/. */
export function outDir(app) {
  const out = join(process.cwd(), "scripts", `augloop-${app}-tone-out`, new Date().toISOString().replace(/[:.]/g, "-"));
  mkdirSync(out, { recursive: true });
  return out;
}

/** A cell's file-name prefix: its place in the list and its name. */
export function slugOf(i, name) {
  return `${String(i).padStart(2, "0")}-${name.replace(/[^A-Za-z0-9._-]+/g, "_")}`;
}

/** One turn of `prompt` with `model` (a proxy model ID, raw forms included) in a new conversation; its files go to `out`. */
export async function runCell(app, model, prompt, out, slug) {
  const temporary = process.env.M365_SAVE_HISTORY !== "1";
  const frames = [];
  const opts = {
    getToken: getAugLoopToken,
    getGraphToken,
    ...(temporary ? { deleteChats: async (ids) => deleteAgentChats(await getToken(), ids, app) } : {}),
    onFrame: (dir, frame) => {
      let parsed = frame;
      try { parsed = JSON.parse(frame); } catch { /* keep the text */ }
      frames.push({ t: new Date().toISOString(), dir, frame: parsed });
    },
  };
  const conv = app === "word" ? new WordAgentConversation(opts) : new AugLoopConversation(opts);
  const row = { model, result: null, errorCode: null, error: null, origin: null, adjusted: null, text: "", ms: 0, files: [] };
  const t0 = Date.now();
  try {
    const stream = await conv.chat(prompt, model);
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
  if (frames.length) {
    const file = `${slug}.frames.jsonl`;
    writeFileSync(join(out, file), frames.map((f) => JSON.stringify(f)).join("\n") + "\n");
    row.files.push(file);
  }
  const state = app === "excel" ? conv.lastSavedState : null;
  if (state) {
    const file = `${slug}.state.zip`;
    writeFileSync(join(out, file), Buffer.from(state, "base64"));
    row.files.push(file);
  }
  return row;
}

/** Write results.json: the prompt and the rows. */
export function writeResults(out, prompt, rows) {
  writeFileSync(join(out, "results.json"), JSON.stringify({ prompt, rows }, null, 2));
}

/** The row's outcome in a few words: the result, or the error. */
export function outcomeOf(row) {
  if (row.error) return `error: ${row.error.slice(0, 120)}`;
  return `${row.result ?? "no result"}${row.errorCode ? ` (${row.errorCode})` : ""}`;
}
