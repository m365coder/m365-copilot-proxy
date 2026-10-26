import { describe, expect, it, afterEach } from "vitest";
import {
  AGENT_CLAUDE_TONES,
  AGENTLESS_TOOL_TONES,
  couldBeLicenceRefusalPrefix,
  getAvailableModels,
  getScenarioForModel,
  hasPaidSeat,
  isLicenceRefusal,
  notePaidSeat,
  prefersPaidScenario,
  resetPaidSeat,
  getScenarioForTone,
  getToneForModel,
  INCLUDED_SCENARIO_MODELS,
  isIncludedOpusModel,
  isIncludedSonnetModel,
  meteredBudgetOf,
  priorityAccessFallbackModel,
  isSonnet5Model,
  isSonnet55Model,
  isAgentRouteAlive,
  isClaudeAutoPaidModel,
  isAgentlessRouteAlive,
  isPremiumOnlyModel,
  noteAgentlessRouteAlive,
  noteAgentlessRouteDead,
  noteAgentRouteAlive,
  noteAgentRouteDead,
  PAID_SCENARIO_TONES,
  PREMIUM_ONLY_AGENT_TONES,
  requestUsesAgent,
  resetAgentRoutes,
  toneUsesToolAgent,
  toolRequestUsesAgent,
} from "./copilot.js";

// The seat a test learned (or set) must not route the next test's request.
afterEach(() => resetPaidSeat());

const INCLUDED = { scenario: "OfficeWebIncludedCopilot", licenseType: "Starter" };
const PAID = { scenario: "OfficeWebPaidCopilot", licenseType: "Premium" };

describe("Sonnet routing — one tone, two models", () => {
  // `Claude_Sonnet` is Sonnet 4.6 on the included scenario and Sonnet 5.5 on
  // the paid one; `Claude_Sonnet_5` is Sonnet 5 on the paid one and Sonnet 4.6
  // on the included one (tone-probe 2026-10-06), so the model ID has to pick.
  afterEach(() => {
    delete process.env.M365_SCENARIO;
  });

  it("sends claude-sonnet-5.5 to the shared tone under the paid scenario", () => {
    expect(getToneForModel("claude-sonnet-5.5")).toBe("Claude_Sonnet");
    expect(getScenarioForModel("claude-sonnet-5.5")).toEqual(PAID);
    expect(getAvailableModels()).toContain("claude-sonnet-5.5");
  });

  it("sends claude-sonnet-5 to its own tone, which is paid", () => {
    expect(getToneForModel("claude-sonnet-5")).toBe("Claude_Sonnet_5");
    expect(getScenarioForModel("claude-sonnet-5")).toEqual(PAID);
    expect(getScenarioForTone("Claude_Sonnet_5")).toEqual(PAID);
    expect(getAvailableModels()).toContain("claude-sonnet-5");
  });

  it("keeps claude-sonnet-4.6 (and its legacy -4.5) on the included scenario everywhere: Sonnet 4.6, off Sonnet 5.5's budget", () => {
    for (const id of ["claude-sonnet-4.5", "claude-sonnet-4.6", "claude-sonnet-4-6-20260101", "claude-sonnet-4-5-20250929"]) {
      expect(getToneForModel(id)).toBe("Claude_Sonnet");
      expect(isIncludedSonnetModel(id)).toBe(true);
      notePaidSeat(true);
      expect(getScenarioForModel(id)).toEqual(INCLUDED);
      expect(isSonnet55Model(id)).toBe(false);
      resetPaidSeat();
    }
  });

  it("serves claude-sonnet as Sonnet 5.5 on a premium account and Sonnet 4.6 on any other (docs §41)", () => {
    for (const id of ["claude", "claude-sonnet"]) {
      expect(getScenarioForModel(id)).toEqual(PAID); // tried while the seat is unknown
      expect(isSonnet55Model(id)).toBe(true);
      expect(meteredBudgetOf(id)).toBe("sonnet-5.5");
      notePaidSeat(false);
      expect(getScenarioForModel(id)).toEqual(INCLUDED);
      expect(isSonnet55Model(id)).toBe(false);
      expect(meteredBudgetOf(id)).toBeNull();
      resetPaidSeat();
    }
  });

  it("does NOT move the Claude_Sonnet tone itself onto the paid scenario", () => {
    // That would silently turn `claude-sonnet` into Sonnet 5.5 as well.
    expect(getScenarioForTone("Claude_Sonnet")).toEqual(INCLUDED);
  });

  it("routes unmapped Sonnet 5.5 strings a client may send to Sonnet 5.5, not 4.6 or 5", () => {
    for (const id of ["claude-sonnet-5.5[1m]", "claude-sonnet-5-5-20261001", "Claude-Sonnet-5.5", "claude-sonnet-5_5"]) {
      expect(getToneForModel(id)).toBe("Claude_Sonnet");
      expect(isSonnet55Model(id)).toBe(true);
      expect(isSonnet5Model(id)).toBe(false);
      expect(getScenarioForModel(id)).toEqual(PAID);
    }
  });

  it("routes unmapped Sonnet 5 strings a client may send to Sonnet 5, not 4.6 or 5.5", () => {
    for (const id of ["claude-sonnet-5[1m]", "claude-sonnet-5-20260115", "Claude-Sonnet-5"]) {
      expect(getToneForModel(id)).toBe("Claude_Sonnet_5");
      expect(isSonnet5Model(id)).toBe(true);
      expect(isSonnet55Model(id)).toBe(false);
      expect(getScenarioForModel(id)).toEqual(PAID);
    }
  });

  it("does not mistake older Sonnet names for Sonnet 5 or 5.5", () => {
    // Without a paid seat; on a premium account the shared tone is Sonnet 5.5 anyway.
    notePaidSeat(false);
    for (const id of ["claude-sonnet-4-5-20250929", "claude-3-5-sonnet", "claude-sonnet-4.5", "claude-sonnet-50", "claude-sonnet-5.50", "claude-sonnet-5.1"]) {
      expect(isSonnet5Model(id)).toBe(false);
      expect(isSonnet55Model(id)).toBe(false);
      expect(getScenarioForModel(id)).toEqual(INCLUDED);
    }
    resetPaidSeat();
  });

  it("never treats a non-Sonnet tone as Sonnet 5 or 5.5, whatever the string says", () => {
    expect(isSonnet5Model("claude-opus-5")).toBe(false);
    expect(isSonnet5Model("gpt-sonnet-5")).toBe(false); // resolves to magic, not a Claude tone
    expect(isSonnet55Model("gpt-sonnet-5.5")).toBe(false);
  });

  it("still derives the paid scenario from the tone for Opus and GPT-6", () => {
    expect(getScenarioForModel("claude-opus")).toEqual(PAID);
    expect(getScenarioForModel("claude-opus-5[1m]")).toEqual(PAID);
    expect(getScenarioForModel("gpt-6-think-deeper")).toEqual(PAID);
    // GPT-5.5 is paid only until the account turns out to have no paid seat (§41).
    expect(getScenarioForModel("gpt-5.5-think-deeper")).toEqual(PAID);
    notePaidSeat(false);
    expect(getScenarioForModel("gpt-5.5-think-deeper")).toEqual(INCLUDED);
    resetPaidSeat();
  });

  it("lets the env override win for model routing too", () => {
    process.env.M365_SCENARIO = "SomeOtherScenario";
    expect(getScenarioForModel("claude-sonnet").scenario).toBe("SomeOtherScenario");
    expect(getScenarioForModel("claude-sonnet-5").scenario).toBe("SomeOtherScenario");
    expect(getScenarioForModel("claude-sonnet-5.5").scenario).toBe("SomeOtherScenario");
  });
});

describe("GPT-5.6 model routing", () => {
  it("maps the advertised model ID to the live-validated reasoning tone", () => {
    expect(getToneForModel("gpt-5.6-think-deeper")).toBe("Gpt_5_6_Reasoning");
    expect(getAvailableModels()).toContain("gpt-5.6-think-deeper");
  });

  it("maps both chat IDs to Gpt_5_6_Chat, mirroring the GPT-5.5 naming", () => {
    // The web client calls this one "GPT 5.6 Quick response", so `-quick` is
    // the alias a user reaches for; bare `gpt-5.6` follows `gpt-5.5`.
    expect(getToneForModel("gpt-5.6")).toBe("Gpt_5_6_Chat");
    expect(getToneForModel("gpt-5.6-quick")).toBe("Gpt_5_6_Chat");
    expect(getAvailableModels()).toContain("gpt-5.6");
    expect(getAvailableModels()).toContain("gpt-5.6-quick");
  });

  it("requests NO paid scenario for the chat tone — it serves on the included one", () => {
    // It was entitlement-gated (BotConnection on included, DeepLeo on paid),
    // which §12.15 misread as a dead route. The gate has since lifted, so
    // asking for the paid scenario here would be a bypass attempt, not a fix.
    expect(getScenarioForTone("Gpt_5_6_Chat")).toEqual({
      scenario: "OfficeWebIncludedCopilot",
      licenseType: "Starter",
    });
  });
});

describe("the §44 Gpt_5_N_Auto tones", () => {
  afterEach(() => resetAgentRoutes());

  it("maps gpt-5.N-auto to Gpt_5_N_Auto, routed like its siblings: paid on a premium account, included elsewhere, unmetered", () => {
    for (const n of ["2", "3", "4"]) {
      const id = `gpt-5.${n}-auto`;
      expect(getToneForModel(id)).toBe(`Gpt_5_${n}_Auto`);
      expect(getAvailableModels()).toContain(id);
      notePaidSeat(true);
      expect(getScenarioForModel(id)).toEqual(PAID);
      expect(meteredBudgetOf(id)).toBeNull();
      notePaidSeat(false);
      expect(getScenarioForModel(id)).toEqual(INCLUDED);
      resetPaidSeat();
    }
  });

  it("sends their tool requests with the agent, like the other GPT-5.x tones", () => {
    for (const n of ["2", "3", "4"]) expect(requestUsesAgent(`gpt-5.${n}-auto`, true)).toBe(true);
  });
});

describe("GPT-6 routing", () => {
  it("maps the advertised model ID to the reasoning tone", () => {
    expect(getToneForModel("gpt-6-think-deeper")).toBe("Gpt_6_Reasoning");
    expect(getAvailableModels()).toContain("gpt-6-think-deeper");
  });

  it("advertises no chat variant — Gpt_6_Chat is rejected by the validator", () => {
    // A rejected tone must never be reachable: it errors the whole turn rather
    // than degrading to prose, so shipping an ID that resolves to it would be a
    // model that can only ever fail.
    const advertisedTones = getAvailableModels().map(getToneForModel);
    expect(advertisedTones).not.toContain("Gpt_6_Chat");
  });
});

describe("which tool requests carry the tool agent", () => {
  const usesAgent = (model: string) => toneUsesToolAgent(getToneForModel(model));

  it("keeps GPT-6 off the agent — it is a dead route with it attached (#41)", () => {
    expect(AGENTLESS_TOOL_TONES.has("Gpt_6_Reasoning")).toBe(true);
    expect(usesAgent("gpt-6-think-deeper")).toBe(false);
  });

  it("keeps every Claude model but Opus off the agent, mapped or not", () => {
    for (const id of ["claude", "claude-sonnet", "claude-sonnet-5", "claude-sonnet-5.5", "claude-sonnet-5[1m]", "claude-sonnet-think-deeper", "claude-haiku-9"]) {
      expect(usesAgent(id)).toBe(false);
    }
  });

  it("gives both Opus models the agent — agent-less, Opus 4.8 works in its own sandbox (§24, §36)", () => {
    expect(AGENT_CLAUDE_TONES.has("Claude_Opus")).toBe(true);
    for (const id of ["claude-opus", "claude-opus-5", "claude-opus-5.5", "claude-opus-4.8", "claude-opus-4.5", "claude-opus-5[1m]", "claude-opus-4-8-20261001"]) {
      expect(usesAgent(id)).toBe(true);
    }
  });

  it("still gives the GPT-5.x and default models the agent — they won't act without it", () => {
    for (const id of ["m365-copilot", "quick", "think-deeper", "gpt-5.2", "gpt-5.4", "gpt-5.5", "gpt-5.5-think-deeper", "gpt-5.6", "gpt-5.6-think-deeper"]) {
      expect(usesAgent(id)).toBe(true);
    }
  });
});

describe("GPT-6 Sol routing (#23)", () => {
  afterEach(() => {
    resetAgentRoutes();
    delete process.env.M365_FORCE_AGENT;
  });

  it("maps gpt-6-sol to Gpt_6_Sol_Reasoning — not gated: paid on a premium account, included elsewhere (§41)", () => {
    expect(getToneForModel("gpt-6-sol")).toBe("Gpt_6_Sol_Reasoning");
    expect(getAvailableModels()).toContain("gpt-6-sol");
    expect(PAID_SCENARIO_TONES.has("Gpt_6_Sol_Reasoning")).toBe(false);
    expect(getScenarioForModel("gpt-6-sol")).toEqual(PAID);
    expect(meteredBudgetOf("gpt-6-sol")).toBeNull(); // its paid turns carry GPT61Sol* keys and lower neither (§40 F100)
    notePaidSeat(false);
    expect(getScenarioForModel("gpt-6-sol")).toEqual(INCLUDED);
    resetPaidSeat();
  });

  it("doesn't advertise Gpt_6_Sol_Chat — it serves the GPT-5 chat model", () => {
    expect(getAvailableModels().map(getToneForModel)).not.toContain("Gpt_6_Sol_Chat");
  });

  it("tries the agent first: it serves with it on a premium account", () => {
    expect(PREMIUM_ONLY_AGENT_TONES.has("Gpt_6_Sol_Reasoning")).toBe(true);
    expect(toneUsesToolAgent("Gpt_6_Sol_Reasoning")).toBe(true);
  });

  it("goes agent-less once the account's agent route proved dead, and only for that tone", () => {
    noteAgentRouteDead("Gpt_6_Sol_Reasoning");
    expect(toneUsesToolAgent("Gpt_6_Sol_Reasoning")).toBe(false);
    expect(toneUsesToolAgent("Gpt_5_5_Reasoning")).toBe(true);
    resetAgentRoutes();
    expect(toneUsesToolAgent("Gpt_6_Sol_Reasoning")).toBe(true);
  });

  it("won't mark a route dead once it has answered — a later InternalError is a transient", () => {
    // A premium account produced the dead route's exact wire state mid-conversation (§23).
    noteAgentRouteAlive("Gpt_6_Sol_Reasoning");
    expect(isAgentRouteAlive("Gpt_6_Sol_Reasoning")).toBe(true);
    expect(noteAgentRouteDead("Gpt_6_Sol_Reasoning")).toBe(false);
    expect(toneUsesToolAgent("Gpt_6_Sol_Reasoning")).toBe(true);
    resetAgentRoutes();
    expect(isAgentRouteAlive("Gpt_6_Sol_Reasoning")).toBe(false);
    expect(noteAgentRouteDead("Gpt_6_Sol_Reasoning")).toBe(true);
  });

  it("lets M365_FORCE_AGENT=1 / =0 override the rule either way", () => {
    noteAgentRouteDead("Gpt_6_Sol_Reasoning");
    process.env.M365_FORCE_AGENT = "1";
    expect(toolRequestUsesAgent("Gpt_6_Sol_Reasoning")).toBe(true);
    expect(toolRequestUsesAgent("Claude_Sonnet")).toBe(true);
    process.env.M365_FORCE_AGENT = "0";
    expect(toolRequestUsesAgent("Gpt_5_5_Reasoning")).toBe(false);
    delete process.env.M365_FORCE_AGENT;
    expect(toolRequestUsesAgent("Gpt_5_5_Reasoning")).toBe(true);
    expect(toolRequestUsesAgent("Gpt_6_Sol_Reasoning")).toBe(false);
  });
});

describe("GPT-6.1 Sol routing (§27)", () => {
  afterEach(() => {
    resetAgentRoutes();
  });

  it("maps gpt-6.1-sol to Gpt_61_Sol_Reasoning — paid and metered on a premium account, included elsewhere (§41)", () => {
    expect(getToneForModel("gpt-6.1-sol")).toBe("Gpt_61_Sol_Reasoning");
    expect(getAvailableModels()).toContain("gpt-6.1-sol");
    expect(PAID_SCENARIO_TONES.has("Gpt_61_Sol_Reasoning")).toBe(false);
    expect(getScenarioForModel("gpt-6.1-sol")).toEqual(PAID);
    expect(meteredBudgetOf("gpt-6.1-sol")).toBe("gpt-6.1-sol");
    notePaidSeat(false);
    expect(getScenarioForModel("gpt-6.1-sol")).toEqual(INCLUDED);
    expect(meteredBudgetOf("gpt-6.1-sol")).toBeNull();
    resetPaidSeat();
  });

  it("keeps gpt-6.1-sol-included on the included scenario on every account — its budget fallback", () => {
    expect(getToneForModel("gpt-6.1-sol-included")).toBe("Gpt_61_Sol_Reasoning");
    expect(getScenarioForModel("gpt-6.1-sol-included")).toEqual(INCLUDED);
    expect(meteredBudgetOf("gpt-6.1-sol-included")).toBeNull();
    expect(isPremiumOnlyModel("gpt-6.1-sol-included")).toBe(false);
    process.env.M365_GPT61_SOL_FALLBACK_MODEL = "gpt-6.1-sol-included";
    expect(priorityAccessFallbackModel("gpt-6.1-sol")).toBe("gpt-6.1-sol-included");
    process.env.M365_GPT61_SOL_FALLBACK_MODEL = "gpt-6.1-sol";
    expect(priorityAccessFallbackModel("gpt-6.1-sol")).toBeNull(); // metered itself
    delete process.env.M365_GPT61_SOL_FALLBACK_MODEL;
  });

  it("takes the agent only on a premium account, learned per tone", () => {
    expect(PREMIUM_ONLY_AGENT_TONES.has("Gpt_61_Sol_Reasoning")).toBe(true);
    expect(toneUsesToolAgent("Gpt_61_Sol_Reasoning")).toBe(true);
    noteAgentRouteDead("Gpt_61_Sol_Reasoning");
    expect(toneUsesToolAgent("Gpt_61_Sol_Reasoning")).toBe(false);
    expect(toneUsesToolAgent("Gpt_6_Sol_Reasoning")).toBe(true);
  });
});

describe("the §46 Claude_* tones: Claude_Auto, Claude_Chat, Claude_Container, Claude_Fable", () => {
  afterEach(() => {
    delete process.env.M365_FORCE_AGENT;
    resetAgentRoutes();
  });

  const IDS = { "claude-auto": "Claude_Auto", "claude-chat": "Claude_Chat", "claude-container": "Claude_Container", "claude-fable": "Claude_Fable" };

  it("maps each ID to its tone, and routes it like any other tone: paid on a premium account, included elsewhere (§41)", () => {
    for (const [id, tone] of Object.entries(IDS)) {
      expect(getToneForModel(id)).toBe(tone);
      expect(getAvailableModels()).toContain(id);
      expect(PAID_SCENARIO_TONES.has(tone)).toBe(false);
      expect(prefersPaidScenario(id)).toBe(true);
      notePaidSeat(true);
      expect(getScenarioForModel(id)).toEqual(PAID);
      notePaidSeat(false);
      expect(getScenarioForModel(id)).toEqual(INCLUDED);
      resetPaidSeat();
    }
  });

  it("meters claude-auto on the paid scenario on Opus 5.5's budget, and nothing else (§46 F112)", () => {
    notePaidSeat(true);
    expect(isClaudeAutoPaidModel("claude-auto")).toBe(true);
    expect(meteredBudgetOf("claude-auto")).toBe("opus");
    for (const id of ["claude-chat", "claude-container", "claude-fable"]) expect(meteredBudgetOf(id)).toBeNull();
    notePaidSeat(false);
    expect(isClaudeAutoPaidModel("claude-auto")).toBe(false);
    expect(meteredBudgetOf("claude-auto")).toBeNull();
    process.env.M365_OPUS_FALLBACK_MODEL = "claude-auto";
    notePaidSeat(true);
    expect(priorityAccessFallbackModel("opus")).toBeNull(); // metered itself on a premium account
    delete process.env.M365_OPUS_FALLBACK_MODEL;
  });

  it("sends their tool requests agent-less, like every Claude tone but Opus; M365_FORCE_AGENT still overrides", () => {
    for (const [id, tone] of Object.entries(IDS)) {
      expect(AGENT_CLAUDE_TONES.has(tone)).toBe(false);
      expect(toneUsesToolAgent(tone)).toBe(false);
      expect(requestUsesAgent(id, true)).toBe(false);
      expect(isPremiumOnlyModel(id)).toBe(false);
    }
    process.env.M365_FORCE_AGENT = "1";
    expect(requestUsesAgent("claude-chat", true)).toBe(true);
  });
});

describe("Opus routing", () => {
  afterEach(() => {
    delete process.env.M365_SCENARIO;
    delete process.env.M365_FORCE_AGENT;
  });

  it("maps the advertised Opus IDs to the Claude_Opus tone", () => {
    expect(getToneForModel("claude-opus")).toBe("Claude_Opus");
    expect(getToneForModel("claude-opus-5")).toBe("Claude_Opus");
    expect(getToneForModel("claude-opus-5.5")).toBe("Claude_Opus");
    expect(getToneForModel("claude-opus-4.8")).toBe("Claude_Opus");
    expect(getToneForModel("claude-opus-4.5")).toBe("Claude_Opus");
    for (const id of ["claude-opus", "claude-opus-5.5", "claude-opus-4.8", "claude-opus-4.5"]) {
      expect(getAvailableModels()).toContain(id);
    }
  });

  it("routes an unmapped Opus string to Opus rather than downgrading to Sonnet", () => {
    // What a Claude Code client actually sends: Opus 5 with a context suffix.
    expect(getToneForModel("claude-opus-5[1m]")).toBe("Claude_Opus");
  });

  it("still routes other unmapped claude-* strings to Sonnet", () => {
    expect(getToneForModel("claude-haiku-9")).toBe("Claude_Sonnet");
  });

  it("sends claude-opus-4.8 to the shared tone under the INCLUDED scenario (Opus 4.8)", () => {
    // claude-opus-4.5 is the legacy ID of the same route (Opus 4.5 until October 2026, §36).
    for (const id of ["claude-opus-4.8", "claude-opus-4.5"]) {
      expect(INCLUDED_SCENARIO_MODELS.has(id)).toBe(true);
      expect(isIncludedOpusModel(id)).toBe(true);
      expect(getScenarioForModel(id)).toEqual(INCLUDED);
    }
  });

  it("keeps every other Opus ID on the paid scenario (Opus 5.5)", () => {
    for (const id of ["claude-opus", "claude-opus-5", "claude-opus-5.5", "claude-opus-5[1m]", "claude-opus-4", "claude-opus-4-1", "claude-opus-4.6", "claude-opus-4.50", "claude-opus-4.80"]) {
      expect(isIncludedOpusModel(id)).toBe(false);
      expect(getScenarioForModel(id)).toEqual(PAID);
    }
    // The tone itself stays paid; only the model ID opts out.
    expect(getScenarioForTone("Claude_Opus")).toEqual(PAID);
  });

  it("routes unmapped Opus 4.8 (and legacy 4.5) strings a client may send to the included Opus, not Opus 5.5", () => {
    for (const id of [
      "claude-opus-4-8", "claude-opus-4-8-20261001", "claude-opus-4.8[1m]", "Claude-Opus-4.8", "opus_4_8",
      "claude-opus-4-5", "claude-opus-4-5-20251101", "claude-opus-4.5[1m]", "Claude-Opus-4.5", "opus_4_5",
    ]) {
      expect(getToneForModel(id)).toBe("Claude_Opus");
      expect(isIncludedOpusModel(id)).toBe(true);
      expect(getScenarioForModel(id)).toEqual(INCLUDED);
    }
  });

  it("never treats a non-Opus tone as the included Opus, whatever the string says", () => {
    expect(isIncludedOpusModel("claude-sonnet-4.5")).toBe(false);
    expect(isIncludedOpusModel("claude-sonnet-4.8")).toBe(false);
    expect(isIncludedOpusModel("gpt-4.5")).toBe(false);
  });

  it("lets the env override win for Opus 4.8 too", () => {
    process.env.M365_SCENARIO = "SomeOtherScenario";
    expect(getScenarioForModel("claude-opus-4.8").scenario).toBe("SomeOtherScenario");
  });

  it("attaches the agent to Opus 4.8's tool requests only — it serves agent-less too, unlike Opus 4.5 (§36)", () => {
    for (const id of ["claude-opus-4.8", "claude-opus-4.5", "claude-opus"]) {
      expect(requestUsesAgent(id, false)).toBe(false);
      expect(requestUsesAgent(id, true)).toBe(true);
    }
    // Unchanged for everything else: tool-less requests go agent-less.
    expect(requestUsesAgent("gpt-5.5", false)).toBe(false);
    expect(requestUsesAgent("claude-sonnet", true)).toBe(false);
  });

  it("knows which turns are metered — Opus 5.5's and Sonnet 5.5's, not Opus 4.8's or Sonnet 5's", () => {
    for (const id of ["claude-opus", "claude-opus-5", "claude-opus-5.5", "claude-opus-5[1m]"]) expect(meteredBudgetOf(id)).toBe("opus");
    for (const id of ["claude-sonnet-5.5", "claude-sonnet-5-5-20261001"]) expect(meteredBudgetOf(id)).toBe("sonnet-5.5");
    for (const id of ["claude-opus-4.8", "claude-opus-4.5", "claude-opus-4-5-20251101", "claude-sonnet-5", "claude-sonnet-4.6", "gpt-6-think-deeper"]) {
      expect(meteredBudgetOf(id)).toBeNull();
    }
  });

  it("follows M365_SCENARIO: forced onto the paid scenario, claude-sonnet IS Sonnet 5.5", () => {
    process.env.M365_SCENARIO = "OfficeWebPaidCopilot";
    expect(meteredBudgetOf("claude-sonnet")).toBe("sonnet-5.5");
    delete process.env.M365_SCENARIO;
  });

  it("reads each budget's fallback env var, and refuses a metered model as the fallback", () => {
    expect(priorityAccessFallbackModel("opus")).toBeNull();
    process.env.M365_OPUS_FALLBACK_MODEL = "claude-opus-4.8";
    expect(priorityAccessFallbackModel("opus")).toBe("claude-opus-4.8");
    expect(priorityAccessFallbackModel("sonnet-5.5")).toBeNull();
    process.env.M365_OPUS_FALLBACK_MODEL = "claude-opus";
    expect(priorityAccessFallbackModel("opus")).toBeNull();
    process.env.M365_OPUS_FALLBACK_MODEL = "claude-sonnet-5.5";
    expect(priorityAccessFallbackModel("opus")).toBeNull();
    delete process.env.M365_OPUS_FALLBACK_MODEL;
    process.env.M365_SONNET_FALLBACK_MODEL = "claude-sonnet-5";
    expect(priorityAccessFallbackModel("sonnet-5.5")).toBe("claude-sonnet-5");
    process.env.M365_SONNET_FALLBACK_MODEL = "claude-sonnet-5.5";
    expect(priorityAccessFallbackModel("sonnet-5.5")).toBeNull();
    delete process.env.M365_SONNET_FALLBACK_MODEL;
  });

  it("lets M365_FORCE_AGENT=0 take the agent off Opus 4.8 as well", () => {
    process.env.M365_FORCE_AGENT = "0";
    expect(requestUsesAgent("claude-opus-4.8", true)).toBe(false);
  });

  it("knows the included Opus is premium-only on either path, and nothing else is", () => {
    for (const id of ["claude-opus-4.8", "claude-opus-4.5", "claude-opus-4-8-20261001"]) expect(isPremiumOnlyModel(id)).toBe(true);
    for (const id of ["claude-opus", "claude-opus-5.5", "claude-sonnet", "gpt-6-sol"]) expect(isPremiumOnlyModel(id)).toBe(false);
  });
});

// Tones the live validator REJECTS — every one errored `Failed to invoke 'Chat'`
// in all 3 runs of the 2026-09-28 tone-probe sweep (docs/hypotheses.md §19).
const REJECTED_TONES = [
  "Gpt_Quick", "Gpt_Chat", "Gpt_Reasoning",
  "Gpt_5_2_Quick", "Gpt_5_3_Quick", "Gpt_5_4_Quick",
  "Gpt_6_Chat",
];

describe("retired Quick tones", () => {
  it("never resolves an advertised model to a tone the validator rejects", () => {
    // A rejected tone errors the whole turn, so an ID that resolves to one is a
    // model that can only ever fail.
    const advertisedTones = getAvailableModels().map(getToneForModel);
    for (const tone of REJECTED_TONES) expect(advertisedTones).not.toContain(tone);
  });

  it("maps nothing to a *_Quick tone — the *_Chat tones replaced them", () => {
    for (const id of getAvailableModels()) expect(getToneForModel(id)).not.toMatch(/_Quick$/);
  });

  it("re-points each versioned Quick ID at the same generation's Chat tone", () => {
    expect(getToneForModel("gpt-5.4-quick")).toBe("Gpt_5_4_Chat");
    expect(getToneForModel("gpt-5.3")).toBe("Gpt_5_3_Chat");
    expect(getToneForModel("gpt-5.3-quick")).toBe("Gpt_5_3_Chat");
    expect(getToneForModel("gpt-5.2")).toBe("Gpt_5_2_Chat");
    expect(getToneForModel("gpt-5.2-quick")).toBe("Gpt_5_2_Chat");
    // Bare gpt-5.4 was never a Quick tone and stays on reasoning.
    expect(getToneForModel("gpt-5.4")).toBe("Gpt_5_4_Reasoning");
  });

  it("pins the generic aliases to GPT-5.5, since no unversioned tone survives", () => {
    expect(getToneForModel("quick")).toBe("Gpt_5_5_Chat");
    expect(getToneForModel("think-deeper")).toBe("Gpt_5_5_Reasoning");
  });

  it("keeps advertising every legacy ID, so existing client configs keep working", () => {
    for (const id of ["quick", "think-deeper", "gpt-5.4-quick", "gpt-5.3", "gpt-5.3-quick", "gpt-5.2", "gpt-5.2-quick"]) {
      expect(getAvailableModels()).toContain(id);
    }
  });
});

describe("getScenarioForTone", () => {
  afterEach(() => {
    delete process.env.M365_SCENARIO;
    delete process.env.M365_LICENSE_TYPE;
  });

  it("requests the paid scenario for Opus — the only thing that makes it serve", () => {
    expect(getScenarioForTone("Claude_Opus")).toEqual({
      scenario: "OfficeWebPaidCopilot",
      licenseType: "Premium",
    });
  });

  it("requests the paid scenario for GPT-6 too", () => {
    expect(getScenarioForTone("Gpt_6_Reasoning")).toEqual({
      scenario: "OfficeWebPaidCopilot",
      licenseType: "Premium",
    });
  });

  it("leaves every other tone on the included scenario", () => {
    for (const tone of ["magic", "Claude_Sonnet", "Gpt_5_5_Reasoning", "Gpt_5_6_Reasoning", "Gpt_5_6_Chat"]) {
      expect(getScenarioForTone(tone)).toEqual({
        scenario: "OfficeWebIncludedCopilot",
        licenseType: "Starter",
      });
    }
  });

  it("lets env overrides win, independently", () => {
    process.env.M365_SCENARIO = "SomeOtherScenario";
    expect(getScenarioForTone("Claude_Opus").scenario).toBe("SomeOtherScenario");
    expect(getScenarioForTone("Claude_Opus").licenseType).toBe("Premium");

    delete process.env.M365_SCENARIO;
    process.env.M365_LICENSE_TYPE = "Enterprise";
    expect(getScenarioForTone("magic")).toEqual({
      scenario: "OfficeWebIncludedCopilot",
      licenseType: "Enterprise",
    });
  });
});

describe("premium accounts on the paid scenario (docs §41)", () => {
  afterEach(() => {
    resetPaidSeat();
    delete process.env.M365_PAID_SEAT;
    delete process.env.M365_SCENARIO;
  });

  const PREFERS_PAID = ["m365-copilot", "quick", "think-deeper", "gpt-5.2", "gpt-5.3-think-deeper", "gpt-5.4", "gpt-5.5",
    "gpt-5.6", "gpt-5.6-think-deeper", "gpt-6-sol", "gpt-6.1-sol", "claude-sonnet-think-deeper", "claude-sonnet", "claude"];
  const PAID_ONLY = ["claude-opus", "claude-sonnet-5", "claude-sonnet-5.5", "claude-sonnet-5-5-20261001", "gpt-6-think-deeper"];
  const QUOTA_FREE = ["claude-sonnet-4.6", "claude-sonnet-4.5", "claude-opus-4.8", "gpt-6.1-sol-included"];

  it("tries the paid scenario while the seat is unknown, and keeps it on a premium account", () => {
    expect(hasPaidSeat()).toBeUndefined();
    for (const id of PREFERS_PAID) {
      expect(prefersPaidScenario(id)).toBe(true);
      expect(getScenarioForModel(id)).toEqual(PAID);
    }
    notePaidSeat(true);
    for (const id of PREFERS_PAID) expect(getScenarioForModel(id)).toEqual(PAID);
  });

  it("routes them on the included scenario once the account has no paid seat; paid-only models stay paid", () => {
    notePaidSeat(false);
    for (const id of PREFERS_PAID) expect(getScenarioForModel(id)).toEqual(INCLUDED);
    for (const id of PAID_ONLY) expect(getScenarioForModel(id)).toEqual(PAID);
  });

  it("keeps the quota-free IDs on the included scenario even on a premium account", () => {
    notePaidSeat(true);
    for (const id of QUOTA_FREE) {
      expect(prefersPaidScenario(id)).toBe(false);
      expect(getScenarioForModel(id)).toEqual(INCLUDED);
    }
  });

  it("takes M365_PAID_SEAT over what the process learned", () => {
    notePaidSeat(true);
    process.env.M365_PAID_SEAT = "0";
    expect(getScenarioForModel("gpt-5.5")).toEqual(INCLUDED);
    notePaidSeat(false);
    process.env.M365_PAID_SEAT = "1";
    expect(getScenarioForModel("gpt-5.5")).toEqual(PAID);
  });

  it("still lets M365_SCENARIO win", () => {
    notePaidSeat(false);
    process.env.M365_SCENARIO = "OfficeWebPaidCopilot";
    expect(getScenarioForModel("gpt-5.5").scenario).toBe("OfficeWebPaidCopilot");
  });

  it("recognises the licence refusal (F46), and its text while it streams", () => {
    expect(isLicenceRefusal({ value: "ForbiddenRequest", errorCode: "InvalidCopilotLicense" })).toBe(true);
    expect(isLicenceRefusal({ value: "ForbiddenRequest" })).toBe(true);
    expect(isLicenceRefusal({ value: "ForbiddenRequest", errorCode: "SomethingElse" })).toBe(false);
    expect(isLicenceRefusal({ value: "InternalError" })).toBe(false);
    expect(isLicenceRefusal(null)).toBe(false);
    expect(couldBeLicenceRefusalPrefix("")).toBe(true);
    expect(couldBeLicenceRefusalPrefix("It looks like you don\u2019t have a val")).toBe(true);
    expect(couldBeLicenceRefusalPrefix("It looks like you don't have a valid licence. To get access")).toBe(true);
    expect(couldBeLicenceRefusalPrefix("It looks like the file is empty")).toBe(false);
    expect(couldBeLicenceRefusalPrefix("4")).toBe(false);
  });
});

describe("tool-less requests on a tone whose agent-less route is dead (docs §40 F101)", () => {
  afterEach(() => {
    resetAgentRoutes();
    delete process.env.M365_FORCE_AGENT;
  });

  it("go agent-less until that's learned, then carry the agent", () => {
    expect(requestUsesAgent("gpt-6-sol", false)).toBe(false);
    expect(noteAgentlessRouteDead("Gpt_6_Sol_Reasoning")).toBe(true);
    expect(requestUsesAgent("gpt-6-sol", false)).toBe(true);
    expect(requestUsesAgent("gpt-6-sol", true)).toBe(true);
    expect(requestUsesAgent("gpt-6.1-sol", false)).toBe(false); // per tone
  });

  it("don't, when the agent route is dead too, or under M365_FORCE_AGENT=0", () => {
    noteAgentlessRouteDead("Gpt_6_Sol_Reasoning");
    process.env.M365_FORCE_AGENT = "0";
    expect(requestUsesAgent("gpt-6-sol", false)).toBe(false);
    delete process.env.M365_FORCE_AGENT;
    noteAgentRouteDead("Gpt_6_Sol_Reasoning");
    expect(requestUsesAgent("gpt-6-sol", false)).toBe(false);
  });

  it("isn't learned once agent-less has answered for that tone", () => {
    noteAgentlessRouteAlive("Gpt_6_Sol_Reasoning");
    expect(isAgentlessRouteAlive("Gpt_6_Sol_Reasoning")).toBe(true);
    expect(noteAgentlessRouteDead("Gpt_6_Sol_Reasoning")).toBe(false);
    expect(requestUsesAgent("gpt-6-sol", false)).toBe(false);
  });
});
