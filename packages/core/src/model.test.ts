import { afterEach, describe, expect, it, vi } from "vitest";
import { TurnBudgetHold } from "./turn-budget.js";

// ModelSession.run's reconnect: when the first chat() fails it sends the turn
// again on a fresh CopilotSession, and that turn waits on the budget like any
// other (#77). The CopilotSession is faked: each chat() is one upstream turn.
const wire = vi.hoisted(() => ({ events: [] as string[], failures: 0 }));

// The Excel agent's conversation (augloop.ts), faked the same way.
vi.mock("./augloop.js", () => ({
  AugLoopConversation: class {
    turnCount = 0;
    constructor(private opts: { getToken: () => Promise<string> }) {
      wire.events.push("augloop:new");
    }
    async chat(text: string, model: string) {
      wire.events.push(`augloop:${model}:${await this.opts.getToken()}:${text}`);
      this.turnCount++;
      return { fullText: "ok" };
    }
    close() {
      wire.events.push("augloop:close");
    }
  },
}));

vi.mock("./session.js", () => ({
  CopilotSession: class {
    turnCount = 0;
    chat(token: string) {
      wire.events.push(`chat:${token}`);
      if (wire.failures > 0) {
        wire.failures--;
        return Promise.reject(new Error("WebSocket error: connection failed"));
      }
      return Promise.resolve({ fullText: "ok" });
    }
  },
}));

const { ModelSession } = await import("./model.js");

function session() {
  let n = 0;
  return new ModelSession({ useAgent: false, getToken: async () => `tok${++n}` });
}

afterEach(() => {
  wire.events = [];
  wire.failures = 0;
});

describe("ModelSession.run's reconnect", () => {
  it("waits on the turn budget before it sends the turn again, with a fresh token after a wait", async () => {
    wire.failures = 1;
    const waitForBudget = async () => { wire.events.push("wait"); return { waitedMs: 30_000 }; };
    await session().run("hi", "claude-sonnet", undefined, false, [], waitForBudget);
    expect(wire.events).toEqual(["chat:tok1", "wait", "chat:tok2"]);
  });

  it("keeps the token when the budget let it through at once", async () => {
    wire.failures = 1;
    const waitForBudget = async () => { wire.events.push("wait"); return { waitedMs: 0 }; };
    await session().run("hi", "claude-sonnet", undefined, false, [], waitForBudget);
    expect(wire.events).toEqual(["chat:tok1", "wait", "chat:tok1"]);
  });

  it("doesn't send it when the budget refuses the turn", async () => {
    wire.failures = 1;
    const hold = new TurnBudgetHold("refill", 5 * 60_000, null);
    const waitForBudget = async () => { throw hold; };
    await expect(session().run("hi", "claude-sonnet", undefined, false, [], waitForBudget)).rejects.toBe(hold);
    expect(wire.events).toEqual(["chat:tok1"]);
  });

  it("doesn't wait again when the first turn goes through", async () => {
    const waitForBudget = vi.fn(async () => ({ waitedMs: 0 }));
    await session().run("hi", "claude-sonnet", undefined, false, [], waitForBudget);
    expect(waitForBudget).not.toHaveBeenCalled();
    expect(wire.events).toEqual(["chat:tok1"]);
  });

  it("sends an AugLoop model to the Excel agent with the AugLoop token, and says images don't go", async () => {
    const s = new ModelSession({ useAgent: true, getToken: async () => "sydney", getAugLoopToken: async () => "augloop" });
    await s.run("hi", "augloop-excel-claude-opus-5.5", undefined, true, ["data:image/png;base64,AAAA"]);
    expect(wire.events[0]).toBe("augloop:new");
    expect(wire.events[1]).toMatch(/^augloop:augloop-excel-claude-opus-5\.5:augloop:hi\n\n\(Note: an attached image could not be passed on/);
    expect(wire.events.some((e) => e.startsWith("chat:"))).toBe(false);
    expect(s.turnCount).toBe(1);
    expect(await s.refreshAgent()).toBe(false); // no Copilot Studio agent to re-resolve
  });

  it("keeps the Excel agent's conversation across turns, and starts a new one with the session's", async () => {
    const s = new ModelSession({ useAgent: false, getAugLoopToken: async () => "augloop" });
    await s.run("one", "augloop-excel-gpt-6-sol", undefined, false);
    await s.run("two", "augloop-excel-gpt-6-sol", undefined, false);
    expect(wire.events.filter((e) => e === "augloop:new")).toHaveLength(1);
    expect(s.turnCount).toBe(2);
    s.newConversation();
    expect(s.turnCount).toBe(0);
    await s.run("three", "augloop-excel-gpt-6-sol", undefined, false);
    expect(wire.events.filter((e) => e === "augloop:new")).toHaveLength(2);
    expect(wire.events).toContain("augloop:close");
  });

  it("waits with awaitTurnBudget by default (a no-op unless M365_AVOID_THROTTLING is set)", async () => {
    const saved = process.env.M365_AVOID_THROTTLING;
    delete process.env.M365_AVOID_THROTTLING;
    try {
      wire.failures = 1;
      await session().run("hi", "claude-sonnet", undefined, false);
      expect(wire.events).toEqual(["chat:tok1", "chat:tok1"]);
    } finally {
      if (saved !== undefined) process.env.M365_AVOID_THROTTLING = saved;
    }
  });
});
