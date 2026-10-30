// Ask Word's agent's model contract about model entries (docs/hypotheses.md §49 F131): the request Word
// for the web's client sends when its Copilot pane opens (`OfficeAgentGetModelContractSignal`), and the
// server's answer, per entry. No model turn is sent.
//
// Usage: node scripts/augloop-word-contract.mjs [MODELID[:LEVEL] …]
//   Without arguments, the entries of the proxy's augloop-word-* models. Each argument is an entry:
//   a model ID and, for a GPT model, a reasoning level (`GPT56:low`); without one, `:default`.
// Build first (`pnpm build`).
import { randomUUID } from "node:crypto";
import { AUGLOOP_MODEL_IDS, AugLoopConnection, WORD_AGENT_FLIGHTS, WORD_CLIENT, getAugLoopToken, wordAgentModelOf } from "../packages/core/dist/index.mjs";

const OA = "AugLoop_OfficeAgent_";
const H = (T, B) => ({ T_: T, B_: B });
const args = process.argv.slice(2);
const wanted = args.length
  ? args.map((a) => { const [modelId, level] = a.split(":"); return { modelId, reasoningEffort: level }; })
  : AUGLOOP_MODEL_IDS.map(wordAgentModelOf).filter(Boolean);
const entries = wanted.map((m) => ({
  modelId: m.modelId,
  entryId: `${m.modelId}:${m.reasoningEffort ?? "default"}`,
  displayName: m.modelId,
  buttonLabel: m.modelId,
  family: /claude|opus|sonnet|anthropic|fable/i.test(m.modelId) ? "Claude" : "GPT",
  availabilityState: "available",
  isDefault: false,
  ...(m.reasoningEffort ? { defaultReasoningLevel: m.reasoningEffort } : {}),
}));

const client = { ...WORD_CLIENT, activated: [...WORD_CLIENT.activated, `${OA}OfficeAgentGetModelContractAnnotation`] };
const c = await AugLoopConnection.open(WORD_AGENT_FLIGHTS.map((f) => `${f};`).join(""), await getAugLoopToken(), undefined, client);
const requestId = randomUUID();
const answer = new Promise((resolve, reject) => {
  const timer = setTimeout(() => reject(new Error("no contract answer in 20 s")), 20_000);
  c.listen((m) => {
    if (m.annotationType !== `${OA}OfficeAgentGetModelContractAnnotation`) return;
    const b = m.ops?.[0]?.items?.[0]?.body ?? {};
    if (b.requestId === requestId) { clearTimeout(timer); resolve(b); }
  });
});
c.send({
  cv: c.nextCv(),
  ops: [{
    parentPath: ["Signal", `${OA}OfficeAgentGetModelContractSignal`],
    items: [{ id: randomUUID(), body: { appName: "word", requestId, models: entries, H_: H(`${OA}OfficeAgentGetModelContractSignal`, ["AugLoop_MicrosoftAiPlugins_PluginFunctionInputSignal", "AugLoop_Signals_Signal"]) }, contextId: c.nextContextId() }],
    H_: H("AugLoop_Signals_SignalOperation", ["AugLoop_Core_Operation"]),
  }],
  H_: H("AugLoop_Session_Protocol_SyncMessage", ["AugLoop_Session_Protocol_Message"]),
  messageId: c.nextMessageId(),
});
try {
  const b = await answer;
  console.log(`success=${b.success} selectorEnabled=${b.selectorEnabled}`);
  const back = new Map((b.models ?? []).map((m) => [m.entryId, m]));
  for (const e of entries) {
    const m = back.get(e.entryId);
    console.log(`${e.entryId.padEnd(36)} ${m ? `${m.availabilityState}${m.isDefault ? " (default)" : ""}` : "not in the answer"}`);
  }
  for (const [id, m] of back) if (!entries.some((e) => e.entryId === id)) console.log(`${id.padEnd(36)} ${m.availabilityState} (added by the server)`);
} catch (err) {
  console.log(`[error] ${err.message}`);
} finally {
  c.close();
}
process.exit(0);
