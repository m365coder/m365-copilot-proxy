// RE probe: the Excel agent's models over AugLoop (docs/hypotheses.md §48). For each model
// it sends the turns of TURNS in one conversation through the core's AugLoopConversation
// (the proxy's own path: Entra ID token, Chat only, the configured sensitivity label) and
// reports each answer, the turn's result, and the model the agent's saved state names.
// That record is the server's own, and it is what to read: the models' self-IDs are often
// wrong (F121).
//
// Usage: node scripts/augloop-probe.mjs [MODEL[,MODEL…]]   (default augloop-claude-opus-5.5)
//   MODEL: an AugLoop model ID (augloop-claude-opus-5.5, augloop-claude-sonnet-5.5, …) or
//          augloop:<agentModelType> for any slot (augloop:ClaudeSlot6, augloop:ClaudeSonnetSlot5)
//   env: TURNS="first|||second" (default: one short question), M365_AUGLOOP_LABEL,
//        M365_DUMP_FRAMES=1 (frames under ~/.config/opencode-m365/frames/augloop-*.ndjson),
//        RECONNECT_BEFORE=<turn>: before that turn, let the chat's deletion run and close the
//        connection, so the turn goes over a new one and sends the whole transcript (F129),
//        M365_SAVE_HISTORY=1: keep the chats in Copilot history. Without it, as in the proxy,
//        each chat is deleted from history 5 s after its last turn, and the probe waits for that.
// Cost: one AugLoop turn and one new conversation per model, plus a turn per extra TURNS
// entry. Build first (`pnpm build`); with M365_DEBUG=1 the turns are logged where
// M365_AVOID_THROTTLING's bucket reads them.
import { AugLoopConversation, deleteExcelChats, getAugLoopToken, getGraphToken, getToken } from "../packages/core/dist/index.mjs";

const MODELS = (process.argv[2] || "augloop-claude-opus-5.5").split(",").map((s) => s.trim()).filter(Boolean);
const TURNS = (process.env.TURNS || "In one sentence: what kind of assistant are you, and what can you do here?").split("|||");
const RECONNECT_BEFORE = process.env.RECONNECT_BEFORE === undefined ? -1 : Number(process.env.RECONNECT_BEFORE);
const temporary = process.env.M365_SAVE_HISTORY !== "1";

for (const model of MODELS) {
  const conv = new AugLoopConversation({
    getToken: getAugLoopToken,
    getGraphToken,
    ...(temporary ? { deleteChats: async (ids) => deleteExcelChats(await getToken(), ids) } : {}),
  });
  console.log(`\n=== ${model}`);
  let throttled = false;
  try {
    for (const [i, text] of TURNS.entries()) {
      if (i === RECONNECT_BEFORE) {
        await conv.historySettled();
        conv.close();
        console.log(`[turn ${i}] new connection`);
      }
      const t0 = Date.now();
      const stream = await conv.chat(text, model);
      let answer = "";
      for await (const delta of stream) answer += delta;
      const ms = Date.now() - t0;
      console.log(`[turn ${i}] ${ms} ms, result=${stream.result?.value ?? "-"}${stream.result?.errorCode ? ` (${stream.result.errorCode})` : ""}, origin=${stream.contentOrigin}, reasoning=${(stream.reasoning ?? "").length} chars`);
      console.log(`[turn ${i}] ${JSON.stringify(stream.fullText.slice(0, 600))}`);
      if (stream.result?.value === "Throttled") { console.log("THROTTLED — stopping"); throttled = true; break; }
      if (stream.result?.value && stream.result.value !== "Success") break;
    }
  } catch (err) {
    console.log(`[error] ${err.message}`);
  } finally {
    conv.close();
    await conv.historySettled();
  }
  if (throttled) process.exit(3);
}
process.exit(0);
