export {
  getToken,
  getTokenSilent,
  getTokenForScope,
  getImageArtifactToken,
  getAugLoopToken,
  getGraphToken,
  loginAutomated,
  loadSecrets,
  forceReauth,
} from "./auth.js";

export {
  AugLoopConversation,
  AugLoopConnection,
  AugLoopTurnStream,
  applyOutput,
  buildChatSignal,
  ChatDeleter,
  chatDeletionsSettled,
  deleteExcelChats,
  HISTORY_DELETE_DELAY_MS,
  turnMessages,
  outcomeOfStatus,
  augLoopLabelSetting,
  pickSensitivityLabel,
  resolveSensitivityLabel,
  resetSensitivityLabel,
  lastAgentTurnOf,
  readZipEntry,
  redactFrame,
  type AugLoopConversationOptions,
  type AugLoopOutcome,
  type SensitivityLabel,
  type GraphSensitivityLabel,
} from "./augloop.js";

export {
  augLoopModelOf,
  isAugLoopModel,
  isClaudeAugLoopModel,
  AUGLOOP_MODEL_IDS,
  type AugLoopModel,
} from "./augloop-models.js";

export {
  generateImage,
  fetchImageBytes,
  buildImagePrompt,
  classifyImageFailure,
  ImageGenerationError,
  type GeneratedImage,
  type GenerateImageOptions,
  type ImageOrientation,
  type ImageStyle,
  type ImageGenFailureReason,
} from "./image.js";

export { uploadImage, imageUrls, type ImageAnnotation } from "./image-input.js";

export {
  noteRequestOutcome,
  awaitDegradationBackoff,
  isDegradationBackoff,
  createBackoffController,
  type BackoffController,
  type BackoffOptions,
} from "./auth-recovery.js";

export { getOrCreateAgent, getEnvironmentUrl, environmentUrlFromName } from "./agent.js";

export {
  decodeJwt,
  getToneForModel,
  getAvailableModels,
  getScenarioForTone,
  getScenarioForModel,
  isSonnet5Model,
  isSonnet55Model,
  isClaudeAutoPaidModel,
  isIncludedOpusModel,
  isIncludedSonnetModel,
  toneUsesToolAgent,
  toolRequestUsesAgent,
  isPremiumOnlyModel,
  prefersPaidScenario,
  hasPaidSeat,
  notePaidSeat,
  resetPaidSeat,
  isLicenceRefusal,
  couldBeLicenceRefusalPrefix,
  usesPaidScenario,
  requestUsesAgent,
  meteredBudgetOf,
  priorityAccessFallbackModel,
  PRIORITY_ACCESS_FALLBACK_ENV,
  noteAgentRouteDead,
  noteAgentRouteAlive,
  isAgentRouteAlive,
  noteAgentlessRouteDead,
  noteAgentlessRouteAlive,
  isAgentlessRouteAlive,
  shouldResendToollessWithAgent,
  resetAgentRoutes,
  AGENTLESS_TOOL_TONES,
  AGENT_CLAUDE_TONES,
  PREMIUM_ONLY_AGENT_TONES,
  PAID_SCENARIO_TONES,
  PAID_SCENARIO_MODELS,
  INCLUDED_SCENARIO_MODELS,
  type ScenarioRouting,
  type CopilotStream,
  type CapturedImage,
} from "./copilot.js";

export {
  parsePriorityAccessExhaustion,
  couldBePriorityAccessPrefix,
  secondsUntilReset,
  parseMetering,
  meteredAllowance,
  METERED_BUDGETS,
  priorityAccessExhaustionOf,
  notePriorityAccessExhausted,
  activePriorityAccessExhaustion,
  resetPriorityAccessState,
  type MeteredBudget,
  type PriorityAccessExhaustion,
  type PriorityAccessWindow,
} from "./priority-access.js";

export {
  currentFramingVariant,
  defaultFramingForTone,
  defaultFramingForModel,
  transcriptStyleForVariant,
  FRAMING_VARIANT_NAMES,
  longHeredocAsWrite,
  type TranscriptStyle,
} from "./fenced.js";

export {
  CopilotSession,
  type CopilotSessionOptions,
  type ChatTurnOptions,
  type NativeActionConfig,
} from "./session.js";

export {
  parseActionConfirmation,
  buildResumeInvokeAction,
  shouldAutoConfirm,
  buildNativeActionPrompt,
  NATIVE_ACTION_INSTRUCTIONS,
  ACTION_ALLOWED_MESSAGE_TYPES,
  ACTION_CONFIRM_MESSAGE_TYPES,
  type ActionConfirmation,
} from "./native-actions.js";

export {
  ModelSession,
  type ModelSessionOptions,
} from "./model.js";

export {
  formatMessages,
  formatToolDefinitions,
  formatToolChoiceInstruction,
  getMessageContent,
  parseToolCalls,
  looksLikeConfabulation,
  looksLikeHallucinatedCompletion,
  looksLikeRemoteArtifactCompletion,
  truncateAtFabricatedToolResponse,
  fabricatedToolResponseAt,
  textAfterFirstToolCall,
  isProseDocument,
  formatToolResponse,
  type Message,
  type ToolDef,
  type ToolFunction,
  type ToolChoice,
  type ParsedToolCall,
  type ParseResult,
} from "./tools.js";

export { createLogger, trunc, LOG_PATH, FRAME_DIR, fileLoggingEnabled } from "./log.js";
export {
  budgetConfig,
  parseEvents,
  findLogs,
  readLogEvents,
  replay,
  decide,
  budgetStatus,
  proxyBudgetStatus,
  awaitTurnBudget,
  recordTurn,
  recordThrottle,
  noteOboThrottled,
  noteOboPathAnswered,
  oboPauseMs,
  awaitOboPause,
  resetTurnBudget,
  TurnBudgetHold,
  type BudgetConfig,
  type BudgetEvent,
  type BudgetState,
  type BudgetDecision,
  type AwaitTurnBudgetOptions,
} from "./turn-budget.js";
