// AugLoop models — the Office apps' Copilot agents and the models they reach (docs §13 of the API doc).
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
// Which model served an Excel turn is in the agent state the server asks the
// client to save (`selectedModel` and each message's `model` in events.jsonl);
// that, not a self-ID, is what the table below was checked against
// (scripts/augloop-probe.mjs).
//
// Word's Copilot runs on the shared OfficeAgent workflow and names its model
// directly: the chat signal's AppContext carries `selectedModel`, a `modelId`
// from the model table in Word's client, and for a GPT entry its reasoning level
// as `reasoningEffort` (hypotheses §49). Word's model contract with the server
// echoes whatever it is sent (F131), so it says nothing about which IDs exist.

export type AugLoopApp = "excel" | "word";

/** One model reachable through the Excel agent. */
export interface AugLoopModel {
  /** The Excel agent's `agentModelType`. */
  modelType: string;
  /** Flights for this slot, as Excel's own client sends them (`name:value`). */
  flights: readonly string[];
}

/** One model reachable through Word's agent. */
export interface WordAgentModel {
  /** The entry's `modelId` in Word's model table, sent as the AppContext's `selectedModel`. */
  modelId: string;
  /** The entry's reasoning level, sent as `reasoningEffort` (GPT entries only). */
  reasoningEffort?: string;
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

/** Excel model IDs → slot. Named for the model the slot serves: the same models on Chathub are other routes. */
const EXCEL_MODELS: Readonly<Record<string, string>> = {
  "augloop-excel-claude-opus-5.5": "ClaudeSlot6",
  "augloop-excel-claude-sonnet-5.5": "ClaudeSlot7",
  "augloop-excel-claude-opus-5": "ClaudeSlot4",
  "augloop-excel-gpt-6-sol": "Gpt5Slot4",
  "augloop-excel-gpt-6.1-sol": "Gpt5Slot6",
};

/**
 * Word model IDs → the entries Word for the web's client sends for this account,
 * named as its picker names them (2026-10-30). The client's table also has models
 * behind gates this account's config doesn't open (Claude Fable, Grok, GPT-6
 * Astra), and other aliases for the same model under other gates (`Sonnet55`
 * where `anthropic-claude-sonnet-5-5` is sent here); the proxy advertises only
 * what the client sends.
 */
const WORD_MODELS: Readonly<Record<string, WordAgentModel>> = {
  "augloop-word-claude-sonnet-5.5": { modelId: "anthropic-claude-sonnet-5-5" },
  "augloop-word-claude-opus-5.5": { modelId: "Opus55" },
  "augloop-word-gpt-5.6-sol-quick": { modelId: "GPT56", reasoningEffort: "none" },
  "augloop-word-gpt-5.6-sol-think-deeper": { modelId: "GPT56", reasoningEffort: "low" },
  "augloop-word-gpt-6.1-sol": { modelId: "GPT6-sol", reasoningEffort: "low" },
};

/**
 * The OfficeAgent flights Word for the web sends in its session (2026-10-30),
 * copied as sent except `EnableGpt6Sol`, which the proxy turns on; they include
 * the ones that enable and name Word's models (`EnableSonnet55`,
 * `AgentRuntime.Sonnet55Model`, …). Word's session sends no metering flight.
 */
export const WORD_AGENT_FLIGHTS: readonly string[] = [
  "Microsoft.Office.WordOnline.Augloop.OfficeAgent.AgentRuntime.Sonnet55Model:True",
  "Microsoft.Office.WordOnline.Augloop.OfficeAgent.AgentRuntime.Gpt6SolModelAlias:False",
  "Microsoft.Office.WordOnline.Augloop.OfficeAgent.EnableSonnet55:True",
  "Microsoft.Office.WordOnline.Augloop.OfficeAgent.NormalizeSonnet55DisabledThinking:True",
  // Word's session for this account sends `False`; the proxy turns it on (hypotheses §49 F138).
  "Microsoft.Office.WordOnline.Augloop.OfficeAgent.EnableGpt6Sol:True",
  "Microsoft.Office.WordOnline.Augloop.OfficeAgent.Word.EnableSearchAccuracy:True",
  "Microsoft.Office.AugLoop.OfficeAgent.EnableExtensibilityPluginReliabilityTelemetry:True",
  "Microsoft.Office.AugLoop.OfficeAgent.EnableWorkspaceId:True",
  "Microsoft.Office.AugLoop.OfficeAgent.EnableExcelAgentModeUsage:True",
  "Microsoft.Office.AugLoop.OfficeAgent.XlsxJs.BrandKitEnabled:True",
  "Microsoft.Office.Augloop.OfficeAgent.ConnectorFetch.EnableReturnedTransientErrorRetry:True",
  "Microsoft.Office.Augloop.OfficeAgent.EnablePluginSkillNamespacing:True",
  "Microsoft.Office.Augloop.OfficeAgent.EnableConversationWrite:True",
  "Microsoft.Office.Augloop.OfficeAgent.PptAgent.RepairPhantomContentTypeOverrides:True",
  "Microsoft.Office.Augloop.OfficeAgent.EnableLargerVQAThumbnails:True",
  "Microsoft.Office.AugLoop.OfficeAgent.EnableExtensibilityLicenseGate:True",
  "Microsoft.Office.AugLoop.OfficeAgent.EnableExtensibilityParallelGetCapabilitiesInitialization:True",
  "Microsoft.Office.WordOnline.FA000000125.AugLoop.OfficeAgent.EnablePPCVaultListConnectors:True",
];

/** `augloop-excel:<agentModelType>` sends any Excel slot as it is: a slot in SLOTS gets its flights, any other none. */
const RAW_PREFIX = "augloop-excel:";

/** The Excel agent's model behind a model ID, or null when the ID isn't an Excel agent one. */
export function augLoopModelOf(model: string): AugLoopModel | null {
  const slot = EXCEL_MODELS[model];
  if (slot) return SLOTS[slot];
  if (model.startsWith(RAW_PREFIX)) {
    const modelType = model.slice(RAW_PREFIX.length).trim();
    if (!/^[A-Za-z][A-Za-z0-9]*$/.test(modelType)) return null;
    return SLOTS[modelType] ?? { modelType, flights: [] };
  }
  return null;
}

/**
 * `augloop-word:<modelId>[:<reasoningEffort>]` sends any model ID as `selectedModel`, as the Excel form
 * sends any slot: for probing which model IDs Word's agent takes (scripts/augloop-word-tone-probe.mjs).
 */
const WORD_RAW_PREFIX = "augloop-word:";

/** Word's agent's model behind a model ID, or null when the ID isn't a Word agent one. */
export function wordAgentModelOf(model: string): WordAgentModel | null {
  const known = WORD_MODELS[model];
  if (known) return known;
  if (model.startsWith(WORD_RAW_PREFIX)) {
    const [modelId, reasoningEffort, ...rest] = model.slice(WORD_RAW_PREFIX.length).trim().split(":");
    if (rest.length || !/^[A-Za-z0-9][A-Za-z0-9._-]*$/.test(modelId ?? "")) return null;
    if (reasoningEffort !== undefined && !/^[a-z]+$/.test(reasoningEffort)) return null;
    return reasoningEffort ? { modelId, reasoningEffort } : { modelId };
  }
  return null;
}

/** Which Office app's agent a model ID goes to, or null when it isn't an AugLoop model. */
export function augLoopAppOf(model: string): AugLoopApp | null {
  if (augLoopModelOf(model)) return "excel";
  if (wordAgentModelOf(model)) return "word";
  return null;
}

/** Whether a model ID goes to an Office app's agent over AugLoop. */
export function isAugLoopModel(model: string): boolean {
  return augLoopAppOf(model) !== null;
}

/** Whether this Excel slot is a Claude one (it needs `isAnthropicAvailable` on the account). */
export function isClaudeAugLoopModel(model: string): boolean {
  return /^Claude/i.test(augLoopModelOf(model)?.modelType ?? "");
}

/** The AugLoop model IDs /v1/models advertises (not the `augloop-excel:<slot>` and `augloop-word:<modelId>` forms). */
export const AUGLOOP_MODEL_IDS: readonly string[] = [...Object.keys(EXCEL_MODELS), ...Object.keys(WORD_MODELS)];
