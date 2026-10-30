// AugLoop model slots — the Excel agent's `agentModelType` (docs §13 of the API doc).
//
// Excel's Copilot (Agent Mode) picks its model by a slot name, `ClaudeSlot6`,
// `Gpt5Slot4`, …, and the client's flights say which model each slot is: the
// server reads `…ExcelAgent.<slot>ModelId` (and the container agent's
// `…OfficeAgent.OfficeAgent.<slot>ModelId`) from `clientMetadata.flights`.
// Without them it serves a default model for the slot: `ClaudeSlot6` was
// claude-opus-4.5, not Opus 5.5 (hypotheses §48). So each slot here carries the
// flights Excel for the web receives from its config service for that slot,
// copied as sent, metering units included: the proxy sends what the first-party
// client sends and nothing it doesn't, so that M365 meters these turns as it
// meters Excel's. The slot names come from the client's enum (`ClaudeSlot1`–`10`,
// `Gpt5Slot1`–`10`, `MAISlot1`–`4`, …); a slot only serves what its flights name.
//
// Which model served a turn is in the agent state the server asks the client to
// save (`selectedModel` and each message's `model` in events.jsonl); that, not a
// self-ID, is what the table below was checked against (scripts/augloop-probe.mjs).

/** One model reachable through the Excel agent. */
export interface AugLoopModel {
  /** The Excel agent's `agentModelType`. */
  modelType: string;
  /** Flights for this slot, as Excel's own client sends them (`name:value`). */
  flights: readonly string[];
}

const FA = "Microsoft.Office.Excel.FA000000124.AugLoop";
const XL = "Microsoft.Office.Excel.AugLoop";

/** The flights for a slot: its model for the Excel agent and for the container agent, and its metering units. */
function slotFlights(prefix: string, slot: string, modelId: string, metering?: { daily: string; weekly: string }): string[] {
  return [
    `${prefix}.ExcelAgent.${slot}ModelId:${modelId}`,
    `${prefix}.OfficeAgent.OfficeAgent.${slot}ModelId:${modelId}`,
    ...(metering ? [`${XL}.DailyMeteringUnit.${slot}:${metering.daily}`, `${XL}.WeeklyMeteringUnit.${slot}:${metering.weekly}`] : []),
  ];
}

/**
 * The slots Excel for the web offers in its model picker (`…AgentMode.<slot>Enabled:True`),
 * with the model and metering units its config names for each (2026-10-30).
 */
const SLOTS: Readonly<Record<string, AugLoopModel>> = {
  ClaudeSlot4: { modelType: "ClaudeSlot4", flights: slotFlights(FA, "ClaudeSlot4", "custom-substrate-anthropic-claude-opus-5", { daily: "ClaudeOpusQueryDailyExcel", weekly: "ClaudeOpusQueryWeeklyExcel" }) },
  ClaudeSlot6: { modelType: "ClaudeSlot6", flights: slotFlights(FA, "ClaudeSlot6", "custom-substrate-anthropic-claude-opus-5-5", { daily: "ClaudeOpus55QueryDailyExcel", weekly: "ClaudeOpus55QueryWeeklyExcel" }) },
  ClaudeSlot7: { modelType: "ClaudeSlot7", flights: slotFlights(FA, "ClaudeSlot7", "custom-substrate-anthropic-claude-sonnet-5-5", { daily: "ClaudeSonnet55DailyExcel", weekly: "ClaudeSonnet55WeeklyExcel" }) },
  Gpt5Slot4: { modelType: "Gpt5Slot4", flights: slotFlights(XL, "Gpt5Slot4", "custom-substrate-gpt-6-sol") },
  Gpt5Slot6: { modelType: "Gpt5Slot6", flights: slotFlights(XL, "Gpt5Slot6", "custom-substrate-gpt-61-sol", { daily: "GPT61SolQueryDailyExcel", weekly: "GPT61SolQueryWeeklyExcel" }) },
};

/** Model IDs → slot. Named for the model the slot serves, with an `augloop-` prefix: the same models on Chathub are other routes. */
const AUGLOOP_MODELS: Readonly<Record<string, string>> = {
  "augloop-claude-opus-5.5": "ClaudeSlot6",
  "augloop-claude-sonnet-5.5": "ClaudeSlot7",
  "augloop-claude-opus-5": "ClaudeSlot4",
  "augloop-gpt-6-sol": "Gpt5Slot4",
  "augloop-gpt-6.1-sol": "Gpt5Slot6",
};

/** `augloop:<agentModelType>` sends any slot as it is: a slot in SLOTS gets its flights, any other none. */
const RAW_PREFIX = "augloop:";

/** The AugLoop model behind a model ID, or null when the ID isn't an AugLoop one. */
export function augLoopModelOf(model: string): AugLoopModel | null {
  const slot = AUGLOOP_MODELS[model];
  if (slot) return SLOTS[slot];
  if (model.startsWith(RAW_PREFIX)) {
    const modelType = model.slice(RAW_PREFIX.length).trim();
    if (!/^[A-Za-z][A-Za-z0-9]*$/.test(modelType)) return null;
    return SLOTS[modelType] ?? { modelType, flights: [] };
  }
  return null;
}

/** Whether a model ID goes to the Excel agent over AugLoop. */
export function isAugLoopModel(model: string): boolean {
  return augLoopModelOf(model) !== null;
}

/** Whether this slot is a Claude one (it needs `isAnthropicAvailable` on the account). */
export function isClaudeAugLoopModel(model: string): boolean {
  return /^Claude/i.test(augLoopModelOf(model)?.modelType ?? "");
}

/** The AugLoop model IDs /v1/models advertises (not the `augloop:<slot>` form). */
export const AUGLOOP_MODEL_IDS: readonly string[] = Object.keys(AUGLOOP_MODELS);
