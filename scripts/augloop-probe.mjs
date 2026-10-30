// RE probe: the Office agents' models over AugLoop (docs/hypotheses.md §48 for Excel, §49 for
// Word). For each model it sends the turns of TURNS in one conversation through the core's
// conversation for that app's agent (the proxy's own path: Entra ID token, Chat only, the
// configured sensitivity label) and reports each answer and the turn's result. For Excel the
// origin names the model the agent's saved state names: that record is the server's own, and it
// is what to read, as the models' self-IDs are often wrong (F121). Word's agent keeps no such
// record on the client; its origin is the model the proxy asked for.
//
// Usage: node scripts/augloop-probe.mjs [MODEL[,MODEL…]]   (default augloop-excel-claude-opus-5.5)
//   MODEL: an AugLoop model ID (augloop-excel-claude-opus-5.5, augloop-word-claude-sonnet-5.5, …) or
//          augloop-excel:<agentModelType> for any Excel slot (augloop-excel:ClaudeSlot6)
//   env: TURNS="first|||second" (default: one short question), M365_AUGLOOP_LABEL,
//        M365_DUMP_FRAMES=1 (frames under ~/.config/opencode-m365/frames/augloop-*.ndjson),
//        SETTLE_BEFORE=<turn>: before that turn, let the chat's deletion run (temporary chat),
//        and send the turn over the same connection,
//        RECONNECT_BEFORE=<turn>: before that turn, let the chat's deletion run and close the
//        connection, so the turn goes over a new one; after a deletion the Excel agent's turn
//        sends the whole transcript (F129),
//        M365_SAVE_HISTORY=1: keep the chats in Copilot history. Without it, as in the proxy,
//        each chat is deleted from history 5 s after its last turn, and the probe waits for that.
// Cost: one AugLoop turn and one new conversation per model, plus a turn per extra TURNS
// entry. Build first (`pnpm build`); with M365_DEBUG=1 the turns are logged where
// M365_AVOID_THROTTLING's bucket reads them.
import { AugLoopConversation, WordAgentConversation, augLoopAppOf, deleteAgentChats, getAugLoopToken, getGraphToken, getToken } from "../packages/core/dist/index.mjs";

const MODELS = (process.argv[2] || "augloop-excel-claude-opus-5.5").split(",").map((s) => s.trim()).filter(Boolean);
const TURNS = (process.env.TURNS || "In one sentence: what kind of assistant are you, and what can you do here?").split("|||");
const RECONNECT_BEFORE = process.env.RECONNECT_BEFORE === undefined ? -1 : Number(process.env.RECONNECT_BEFORE);
const SETTLE_BEFORE = process.env.SETTLE_BEFORE === undefined ? -1 : Number(process.env.SETTLE_BEFORE);
const temporary = process.env.M365_SAVE_HISTORY !== "1";

for (const model of MODELS) {
  const app = augLoopAppOf(model);
  if (!app) { console.log(`\n=== ${model}: not an AugLoop model`); continue; }
  const opts = {
    getToken: getAugLoopToken,
    getGraphToken,
    ...(temporary ? { deleteChats: async (ids) => deleteAgentChats(await getToken(), ids, app) } : {}),
  };
  const conv = app === "word" ? new WordAgentConversation(opts) : new AugLoopConversation(opts);
  console.log(`\n=== ${model}`);
  let throttled = false;
  try {
    for (const [i, text] of TURNS.entries()) {
      if (i === SETTLE_BEFORE) {
        await conv.historySettled();
        console.log(`[turn ${i}] after the chat's deletion, same connection`);
      }
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
      console.log(`[turn ${i}] ${ms} ms, result=${stream.result?.value ?? "-"}${stream.result?.errorCode ? ` (${stream.result.errorCode})` : ""}, origin=${stream.contentOrigin}, reasoning=${(stream.reasoning ?? "").length} chars, cid=${conv.serverConversationId || "-"}`);
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
