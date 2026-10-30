import { afterEach, describe, expect, it, vi } from "vitest";
import { deflateRawSync } from "node:zlib";
import {
  AugLoopTurnStream,
  ChatDeleter,
  chatDeletionsSettled,
  applyOutput,
  augLoopLabelSetting,
  buildChatSignal,
  deleteExcelChats,
  turnMessages,
  outcomeOfStatus,
  pickSensitivityLabel,
  lastAgentTurnOf,
  readZipEntry,
  redactFrame,
} from "./augloop.js";
import { AUGLOOP_MODEL_IDS, augLoopModelOf, isAugLoopModel, isClaudeAugLoopModel } from "./augloop-models.js";
import { getAvailableModels, getToneForModel, meteredBudgetOf, prefersPaidScenario, requestUsesAgent, resetPaidSeat } from "./copilot.js";
import { defaultFramingForModel, FRAMING_VARIANT_NAMES } from "./fenced.js";
import { priorityAccessExhaustionOf } from "./priority-access.js";
import { parseEvents } from "./turn-budget.js";

const ENV = { ...process.env };
afterEach(() => {
  process.env = { ...ENV };
  resetPaidSeat();
});

describe("AugLoop model IDs", () => {
  it("map each advertised ID to its slot and the flights that name the slot's model", () => {
    const opus = augLoopModelOf("augloop-claude-opus-5.5")!;
    expect(opus.modelType).toBe("ClaudeSlot6");
    expect(opus.flights).toContain("Microsoft.Office.Excel.FA000000124.AugLoop.ExcelAgent.ClaudeSlot6ModelId:custom-substrate-anthropic-claude-opus-5-5");
    expect(opus.flights).toContain("Microsoft.Office.Excel.FA000000124.AugLoop.OfficeAgent.OfficeAgent.ClaudeSlot6ModelId:custom-substrate-anthropic-claude-opus-5-5");
    expect(opus.flights).toContain("Microsoft.Office.Excel.AugLoop.DailyMeteringUnit.ClaudeSlot6:ClaudeOpus55QueryDailyExcel");
    expect(augLoopModelOf("augloop-claude-sonnet-5.5")!.modelType).toBe("ClaudeSlot7");
    expect(augLoopModelOf("augloop-gpt-6.1-sol")!.flights).toContain("Microsoft.Office.Excel.AugLoop.ExcelAgent.Gpt5Slot6ModelId:custom-substrate-gpt-61-sol");
    for (const id of AUGLOOP_MODEL_IDS) expect(isAugLoopModel(id)).toBe(true);
  });

  it("send any slot as augloop:<agentModelType>, with the flights only of a slot it knows", () => {
    expect(augLoopModelOf("augloop:ClaudeSlot7")).toEqual(augLoopModelOf("augloop-claude-sonnet-5.5"));
    expect(augLoopModelOf("augloop:ClaudeSonnetSlot5")).toEqual({ modelType: "ClaudeSonnetSlot5", flights: [] });
    expect(augLoopModelOf("augloop:not a slot")).toBeNull();
    expect(augLoopModelOf("claude-opus-5.5")).toBeNull();
    expect(isClaudeAugLoopModel("augloop-claude-opus-5")).toBe(true);
    expect(isClaudeAugLoopModel("augloop-gpt-6-sol")).toBe(false);
  });

  it("route as the Excel agent, not as Chathub's tones", () => {
    expect(getToneForModel("augloop-claude-opus-5.5")).toBe("ClaudeSlot6"); // not Claude_Opus
    expect(getToneForModel("augloop:Gpt5Slot4")).toBe("Gpt5Slot4");
    process.env.M365_FORCE_AGENT = "1";
    expect(requestUsesAgent("augloop-claude-opus-5.5", true)).toBe(false);
    expect(meteredBudgetOf("augloop-claude-opus-5.5")).toBeNull();
    expect(prefersPaidScenario("augloop-claude-sonnet-5.5")).toBe(false);
    expect(getAvailableModels()).toEqual(expect.arrayContaining([...AUGLOOP_MODEL_IDS]));
    expect(getAvailableModels().some((m) => m.startsWith("augloop:"))).toBe(false);
  });

  it("default to a user-voice framing", () => {
    const framing = defaultFramingForModel("augloop-claude-opus-5.5");
    expect(FRAMING_VARIANT_NAMES).toContain(framing);
    expect(framing).not.toBe("baseline");
  });
});

describe("outcomeOfStatus", () => {
  it("maps the agent's statuses onto the results the handler acts on", () => {
    expect(outcomeOfStatus("partial")).toBeNull();
    expect(outcomeOfStatus("complete")).toEqual({ kind: "done", result: { value: "Success" } });
    expect(outcomeOfStatus("error_user_throttled")).toMatchObject({ kind: "done", result: { value: "Throttled", errorCode: "error_user_throttled" } });
    expect(outcomeOfStatus("error_internal_error")).toEqual({ kind: "done", result: { value: "InternalError", errorCode: "error_internal_error" } });
    expect(outcomeOfStatus("error_rai_policy")).toMatchObject({ kind: "disengaged" });
    expect(outcomeOfStatus("error_dlp_blocked_prompt")).toMatchObject({ kind: "disengaged" });
    expect(outcomeOfStatus("error_query_canceled")).toEqual({ kind: "canceled" });
    expect(outcomeOfStatus("error_user_not_licensed")).toMatchObject({ kind: "fail" });
  });

  it("reports a turn limit as OutOfCredits with its window, which the priority-access 429 reads", () => {
    const weekly = outcomeOfStatus("error_claude_opus_weekly_limit");
    expect(weekly).toMatchObject({ kind: "done", result: { value: "OutOfCredits" } });
    const now = new Date("2026-10-30T12:00:00Z");
    if (weekly?.kind !== "done") throw new Error("unexpected");
    expect(priorityAccessExhaustionOf(weekly.result, "", now)?.window).toBe("week");
    const daily = outcomeOfStatus("error_model_daily_limit");
    if (daily?.kind !== "done") throw new Error("unexpected");
    expect(priorityAccessExhaustionOf(daily.result, "", now)).toMatchObject({ window: "day", resetsAt: new Date("2026-10-31T00:00:00Z") });
  });
});

describe("buildChatSignal", () => {
  const base = {
    modelType: "ClaudeSlot6", query: "hi", queryId: "q", signalId: "s", conversationId: "",
    messages: [{ role: "user", content: "hi" }], isClaudeAvailable: true, hostSessionId: "abc", cv: "cv.1", contextId: "C1", messageId: "c9",
  };

  it("asks for Chat only, with the slot and the transcript", () => {
    const body = buildChatSignal({ ...base, sensitivityLabel: { id: "label-id", priority: 2, name: "General Business (G)" } }).ops[0].items[0].body;
    expect(body.documentConfig).toMatchObject({ isReadOnly: true, readOnlyReason: 21, sensitivityLabel: { id: "label-id", priority: 2 } });
    expect(body.documentConfig.sensitivityLabel).not.toHaveProperty("name");
    expect(body.agentModelType).toBe("ClaudeSlot6");
    expect(body.conversation).toEqual({ conversationId: "", messages: [{ role: "user", content: "hi" }], grantedEntityPermissions: [] });
    expect(body.H_.T_).toBe("AugLoop_ExcelAgentExperimental_ExcelAgentExperimentalSignal"); // required: the server refuses a body without it
  });

  it("has no container speedbump, the account's isClaudeAvailable, and no document state", () => {
    const body = buildChatSignal({ ...base, sensitivityLabel: null }).ops[0].items[0].body;
    expect(body.featureOverrides.SpeedbumpInContainerEnabled).toBe(false);
    expect(body.isClaudeAvailable).toBe(true);
    expect(buildChatSignal({ ...base, isClaudeAvailable: false, sensitivityLabel: null }).ops[0].items[0].body.isClaudeAvailable).toBe(false);
    expect(body).not.toHaveProperty("documentState");
  });

  it("sends no label when none is configured", () => {
    const body = buildChatSignal({ ...base, sensitivityLabel: null }).ops[0].items[0].body;
    expect(body.documentConfig).not.toHaveProperty("sensitivityLabel");
  });
});

describe("sensitivity label", () => {
  const labels = [
    { id: "id-nb", name: "Non-Business Information (NB)", displayName: "Non-Business (NB)", priority: 0 },
    { id: "id-g", name: "General Business Information (G)", displayName: "General Business (G)", priority: 2 },
    { id: "id-snpi", name: "Sensitive Group", displayName: "Sensitive", priority: 4, sublabels: [{ id: "id-sub", name: "Sensitive (SNPI)", displayName: "Sensitive (SNPI)", priority: 6 }] },
  ];

  it("is picked by ID, by name, or by the start of a name, case-insensitive", () => {
    expect(pickSensitivityLabel(labels, "General Business")).toEqual({ id: "id-g", priority: 2, name: "General Business (G)" });
    expect(pickSensitivityLabel(labels, "general business (g)")?.id).toBe("id-g");
    expect(pickSensitivityLabel(labels, "ID-NB")?.id).toBe("id-nb");
    expect(pickSensitivityLabel(labels, "Sensitive (SNPI)")?.id).toBe("id-sub");
    expect(pickSensitivityLabel(labels, "Top Secret")).toBeNull();
  });

  it("defaults to General Business, and `none` turns it off", () => {
    delete process.env.M365_AUGLOOP_LABEL;
    expect(augLoopLabelSetting()).toBe("General Business");
    process.env.M365_AUGLOOP_LABEL = "none";
    expect(augLoopLabelSetting()).toBeNull();
    process.env.M365_AUGLOOP_LABEL = "Confidential (C)";
    expect(augLoopLabelSetting()).toBe("Confidential (C)");
  });
});

describe("AugLoopTurnStream", () => {
  async function drain(out: AugLoopTurnStream): Promise<string[]> {
    const parts: string[] = [];
    for await (const d of out.stream()) parts.push(d);
    return parts;
  }

  it("streams the answer, keeps thoughts apart, and ends on the final status", async () => {
    const out = new AugLoopTurnStream();
    const reasoning: string[] = [];
    out.stream().onReasoning?.((d) => reasoning.push(d));
    expect(applyOutput(out, { streamedChunk: { type: "thoughts", text: "Thinking." }, responseStatus: "partial" })).toBeNull();
    expect(applyOutput(out, { streamedChunk: { type: "response", text: "Hello" }, responseStatus: "partial" })).toBeNull();
    expect(applyOutput(out, { chunkType: "response", chunkContent: " world", responseStatus: "partial" })).toBeNull();
    const outcome = applyOutput(out, { finalResponse: { assistantResponse: "Hello world!", thoughts: "Thinking." }, responseStatus: "complete" });
    expect(outcome).toEqual({ kind: "done", result: { value: "Success" } });
    out.finish();
    expect(await drain(out)).toEqual(["Hello", " world", "!"]);
    expect(out.answer).toBe("Hello world!");
    expect(reasoning.join("")).toBe("Thinking.");
  });

  it("starts a new paragraph for a new message, and keeps it when the final text is only the last one", async () => {
    const out = new AugLoopTurnStream();
    applyOutput(out, { streamedChunk: { type: "response", text: "Looking at it." } });
    applyOutput(out, { streamedChunk: { type: "response", text: "```bash\nls\n```", newMessage: true } });
    applyOutput(out, { finalResponse: { assistantResponse: "```bash\nls\n```" }, responseStatus: "complete" });
    out.finish();
    await drain(out);
    expect(out.answer).toBe("Looking at it.\n\n```bash\nls\n```");
  });
});

/** A ZIP with stored and deflated entries (CRCs zero: the reader doesn't check them). */
function makeZip(entries: Array<{ name: string; data: Buffer; deflate?: boolean }>): Buffer {
  const locals: Buffer[] = [];
  const centrals: Buffer[] = [];
  let offset = 0;
  for (const e of entries) {
    const body = e.deflate ? deflateRawSync(e.data) : e.data;
    const name = Buffer.from(e.name);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(e.deflate ? 8 : 0, 8);
    local.writeUInt32LE(body.length, 18);
    local.writeUInt32LE(e.data.length, 22);
    local.writeUInt16LE(name.length, 26);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(e.deflate ? 8 : 0, 10);
    central.writeUInt32LE(body.length, 20);
    central.writeUInt32LE(e.data.length, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt32LE(offset, 42);
    locals.push(local, name, body);
    centrals.push(central, name);
    offset += 30 + name.length + body.length;
  }
  const cd = Buffer.concat(centrals);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(cd.length, 12);
  eocd.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, cd, eocd]);
}

describe("the agent's saved state", () => {
  it("names the models and the agent's own tools of its last turn", () => {
    const events = [
      { type: "session.start", data: { selectedModel: "claude-opus-5.5" } },
      { type: "user.message", data: {} },
      { type: "assistant.message", data: { model: "earlier-model", content: "hi" } },
      { type: "tool.execution_start", data: { toolName: "mcp__ExcelAgentTools__run_script" } },
      { type: "user.message", data: {} },
      { type: "assistant.message", data: { model: "custom-substrate-anthropic-claude-opus-5-5", content: "computing" } },
      { type: "tool.execution_start", data: { toolName: "mcp__ExcelAgentTools__python_execution" } },
      { type: "assistant.message", data: { model: "custom-substrate-anthropic-claude-opus-5-5", content: "done" } },
    ].map((e) => JSON.stringify(e)).join("\n");
    const zip = makeZip([
      { name: "workspace.yaml", data: Buffer.from("cwd: /app\n") },
      { name: "events.jsonl", data: Buffer.from(events), deflate: true },
    ]);
    expect(readZipEntry(zip, "workspace.yaml")?.toString()).toBe("cwd: /app\n");
    expect(lastAgentTurnOf(zip.toString("base64"))).toEqual({ models: ["custom-substrate-anthropic-claude-opus-5-5"], tools: ["python_execution"] });
    expect(lastAgentTurnOf("not a zip")).toEqual({ models: [], tools: [] });
  });

  it("is left out of frame dumps, and so are tokens", () => {
    const token = `eyJhbGciOiJkaXIiLCJlbmMiOiJBMjU2Q0JDLUhTNTEyIn0..${"a".repeat(20)}.${"b".repeat(40)}.${"c".repeat(20)}`;
    const out = redactFrame(JSON.stringify({ authToken: token, content: { stateData: "UEsDBAoAAAAA".repeat(10) } }));
    expect(out).not.toContain("eyJhbGci");
    expect(out).not.toContain("UEsDB");
    expect(out).toContain("<token:");
  });
});

describe("deleteExcelChats", () => {
  afterEach(() => vi.unstubAllGlobals());
  const OID = "01234567-89ab-cdef-0123-456789abcdef";
  const TID = "fedcba98-7654-3210-fedc-ba9876543210";
  const token = `h.${Buffer.from(JSON.stringify({ aud: "a", iss: "i", oid: OID, tid: TID, exp: 0 })).toString("base64url")}.s`;

  it("deletes chats from Copilot history as Excel's Delete does", async () => {
    const fetch = vi.fn(async () => new Response(JSON.stringify({ conversationIdsDeleted: ["conv-1", "conv-2"], result: { value: "Success" } }), { status: 200 }));
    vi.stubGlobal("fetch", fetch);
    await deleteExcelChats(token, ["conv-1", "conv-2"]);
    const [url, init] = fetch.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toBe("https://substrate.office.com/m365Copilot/DeleteConversation");
    expect(init.method).toBe("POST");
    expect(init.headers).toMatchObject({
      Authorization: `Bearer ${token}`,
      "X-Scenario": "ExcelSidepaneWorkCopilot",
      "X-AnchorMailbox": `Oid:${OID}@${TID}`,
      "X-RoutingParameter-SessionKey": OID,
    });
    expect(JSON.parse(String(init.body))).toMatchObject({ conversationIdsToDelete: ["conv-1", "conv-2"], source: "excel" });
  });

  it("fails when the server refuses, or doesn't name every chat as deleted", async () => {
    vi.stubGlobal("fetch", vi.fn(async () => new Response("nope", { status: 401 })));
    await expect(deleteExcelChats(token, ["conv-1"])).rejects.toThrow("HTTP 401");
    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify({ conversationIdsDeleted: ["conv-1"], result: { message: "1 conversation(s) deleted" } }), { status: 200 })));
    await expect(deleteExcelChats(token, ["conv-1", "conv-2"])).rejects.toThrow("1 conversation(s) deleted");
  });
});

describe("a temporary chat (ChatDeleter)", () => {
  afterEach(() => vi.useRealTimers());

  it("is deleted once no turn has started for the delay, and not when the next turn comes first", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
    const deleted: string[][] = [];
    const d = new ChatDeleter(async (ids) => { deleted.push(ids); }, 5_000);
    d.turnEnded("conv-1");
    await vi.advanceTimersByTimeAsync(4_999);
    await d.turnStarting();
    await vi.advanceTimersByTimeAsync(10_000);
    expect(deleted).toEqual([]);
    d.turnEnded("conv-1");
    await vi.advanceTimersByTimeAsync(5_000);
    expect(deleted).toEqual([["conv-1"]]);
    d.turnEnded(""); // nothing due: the server never assigned an ID, and conv-1 is deleted
    await vi.advanceTimersByTimeAsync(5_000);
    expect(deleted).toEqual([["conv-1"]]);
  });

  it("deletes every ID its turns ran in, and tries a failed one again after the next turn", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
    const calls: string[][] = [];
    let failNext = true;
    const d = new ChatDeleter(async (ids) => {
      calls.push(ids);
      if (failNext) { failNext = false; throw new Error("HTTP 500"); }
    }, 5_000);
    d.turnEnded("conv-1");
    await d.turnStarting(); // the next turn came first, and the server moved it to conv-2
    d.turnEnded("conv-2");
    await vi.advanceTimersByTimeAsync(5_000);
    expect(d.wasDeleted("conv-1")).toBe(false); // that deletion failed
    d.turnEnded("conv-2");
    await vi.advanceTimersByTimeAsync(5_000);
    expect(calls).toEqual([["conv-1", "conv-2"], ["conv-1", "conv-2"]]);
    expect(d.wasDeleted("conv-1") && d.wasDeleted("conv-2")).toBe(true);
    expect(d.wasDeleted("conv-3")).toBe(false);
  });

  it("holds the next turn until a deletion in progress is done, and a failed one doesn't stop it", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
    let fail!: () => void;
    const d = new ChatDeleter(() => new Promise<void>((_, reject) => { fail = () => reject(new Error("HTTP 500")); }), 5_000);
    d.turnEnded("conv-1");
    await vi.advanceTimersByTimeAsync(5_000);
    let started = false;
    const next = d.turnStarting().then(() => { started = true; });
    await vi.advanceTimersByTimeAsync(0);
    expect(started).toBe(false);
    fail();
    await next;
    expect(started).toBe(true);
  });

  it("lets a script wait for the deletion before it exits", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
    const deleted: string[][] = [];
    const d = new ChatDeleter(async (ids) => { deleted.push(ids); }, 5_000);
    d.turnEnded("conv-1");
    const idle = d.idle();
    await vi.advanceTimersByTimeAsync(5_000);
    await idle;
    expect(deleted).toEqual([["conv-1"]]);
  });

  it("lets a process wait for every conversation's deletion before it exits", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
    const deleted: string[][] = [];
    const a = new ChatDeleter(async (ids) => { deleted.push(ids); }, 5_000);
    const b = new ChatDeleter(async (ids) => { deleted.push(ids); }, 5_000);
    a.turnEnded("conv-1");
    b.turnEnded("conv-2");
    let settled = false;
    const all = chatDeletionsSettled().then(() => { settled = true; });
    await vi.advanceTimersByTimeAsync(4_999);
    expect(settled).toBe(false);
    await vi.advanceTimersByTimeAsync(1);
    await all;
    expect(deleted).toEqual([["conv-1"], ["conv-2"]]);
  });
});

describe("turnMessages", () => {
  it("sends a turn's own message, and the whole transcript when the session can't be resumed", () => {
    const transcript = [{ role: "user", content: "a" }, { role: "assistant", content: "b" }, { role: "user", content: "c" }];
    expect(turnMessages(transcript, false)).toEqual([{ role: "user", content: "c" }]);
    expect(turnMessages(transcript, true)).toEqual(transcript);
    expect(turnMessages(transcript, true)[0]).not.toBe(transcript[0]);
  });
});

describe("turn budget", () => {
  it("counts the Excel agent's turns and throttles with Chathub's", () => {
    const log = [
      "[2026-10-30T07:00:30.101Z] [INFO] [augloop] Chat turn 0: model=augloop-claude-opus-5.5, agentModelType=ClaudeSlot6",
      "[2026-10-30T07:01:30.101Z] [INFO] [augloop] Turn result: Throttled (error_user_throttled)",
      "[2026-10-30T07:02:30.101Z] [INFO] [model] run: model=augloop-claude-opus-5.5, agent=none, turn=1, via=AugLoop",
    ].join("\n");
    expect(parseEvents(log).map((e) => e.kind)).toEqual(["turn", "throttled"]);
  });
});
