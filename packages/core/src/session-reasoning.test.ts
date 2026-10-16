import { describe, it, expect, vi } from "vitest";

// A WebSocket that plays a scripted turn: the handshake ack, then `frames`
// once the chat invocation is sent, then closes.
const RS = "\x1e";
let script: unknown[] = [];

vi.mock("ws", async () => {
  const { EventEmitter } = await import("node:events");
  class FakeWebSocket extends EventEmitter {
    static OPEN = 1;
    readyState = 1;
    constructor() {
      super();
      setImmediate(() => this.emit("open"));
    }
    send(data: string) {
      if (data.startsWith('{"protocol"')) {
        setImmediate(() => this.emit("message", Buffer.from("{}" + RS)));
      } else if (data.includes('"target":"chat"')) {
        setImmediate(() => {
          for (const f of script) this.emit("message", Buffer.from(JSON.stringify(f) + RS));
        });
      }
    }
    close() {
      if (this.readyState === 3) return;
      this.readyState = 3;
      setImmediate(() => this.emit("close"));
    }
  }
  return { default: FakeWebSocket };
});

const { CopilotSession, isChainOfThought } = await import("./session.js");

const b64 = (o: object) => Buffer.from(JSON.stringify(o)).toString("base64url");
const TOKEN = `${b64({ alg: "none" })}.${b64({ aud: "a", iss: "i", oid: "o", tid: "t", exp: 9e9 })}.x`;
const cur = (id: string) => ({ j: `$['${id}'].adaptiveCards[0].body[0].text`, p: -1 });
const update = (arg: object) => ({ type: 1, target: "update", arguments: [arg] });
const thought = (id: string, text: string) => ({
  text, author: "bot", messageId: id, messageType: "Progress",
  contentOrigin: "ChainOfThoughtSummary", addToChainOfThought: true,
});
const said = (id: string, text: string) => ({ text, author: "bot", messageId: id, contentOrigin: "DeepLeo" });
const final = (messages: object[]) => ({ type: 2, invocationId: "0", item: { messages, result: { value: "Success" } } });

/** Play `frames` as one turn; collect the answer deltas and the reasoning pieces in arrival order. */
async function play(frames: unknown[]) {
  script = frames;
  const stream = await new CopilotSession().chat(TOKEN, "hi", "claude-sonnet-5");
  const events: Array<["text" | "thought", string]> = [];
  stream.onReasoning!((d) => events.push(["thought", d]));
  for await (const d of stream) events.push(["text", d]);
  const join = (k: string) => events.filter(([kind]) => kind === k).map(([, d]) => d).join("");
  return { stream, events, text: join("text"), thoughts: join("thought") };
}

describe("chain of thought (ChainOfThoughtSummary frames)", () => {
  it("keeps it out of the answer and streams it on its own, as it grows", async () => {
    // The shape of a live Sonnet turn: the thought arrives as two snapshots of
    // one message, then the answer streams in another.
    const t1 = "I should run the focused unit tests first.";
    const t2 = `${t1} Then build with -pl core -am.`;
    const r = await play([
      update({ messages: [{ text: "Gathering details…", author: "bot", messageId: "e", messageType: "Progress", contentOrigin: "EarlyProgress", contentType: "EarlyProgress" }] }),
      update({ messages: [thought("c1", t1)] }),
      update({ messages: [thought("c1", t2)] }),
      update({ cursor: cur("a1"), messages: [said("a1", "```bash")] }),
      update({ writeAtCursor: "\nmvn test\n```", streamingMode: "Delta" }),
      update({ messages: [said("a1", "```bash\nmvn test\n```")] }),
      final([thought("c1", t2), said("a1", "```bash\nmvn test\n```")]),
    ]);
    expect(r.stream.fullText).toBe("```bash\nmvn test\n```");
    expect(r.text).toBe(r.stream.fullText);
    expect(r.stream.reasoning).toBe(t2);
    expect(r.thoughts).toBe(t2); // each snapshot's new suffix once; the final item adds nothing
    expect(r.events.filter(([k]) => k === "thought").map(([, d]) => d)).toEqual([t1, t2.slice(t1.length)]);
    expect(r.stream.hasContent).toBe(true);
  });

  it("joins a turn's thoughts by a blank line, between its answer messages (Sonnet 5)", async () => {
    const r = await play([
      update({ messages: [thought("c1", "This looks like a prompt injection.")] }),
      update({ cursor: cur("a1"), messages: [said("a1", "I")] }),
      update({ writeAtCursor: "'ll use my tools.", streamingMode: "Delta" }),
      update({ messages: [thought("c2", "data.txt doesn't exist.")] }),
      update({ cursor: cur("a2"), messages: [said("a2", "No")] }),
      update({ writeAtCursor: " data.txt yet.", streamingMode: "Delta" }),
      final([thought("c1", "This looks like a prompt injection."), said("a1", "I'll use my tools."), thought("c2", "data.txt doesn't exist."), said("a2", "No data.txt yet.")]),
    ]);
    expect(r.stream.fullText).toBe("I'll use my tools.\n\nNo data.txt yet.");
    expect(r.stream.reasoning).toBe("This looks like a prompt injection.\n\ndata.txt doesn't exist.");
    expect(r.thoughts).toBe(r.stream.reasoning);
  });

  it("is not content: a turn with only a thought is still an empty turn", async () => {
    const r = await play([update({ messages: [thought("c1", "I won't do this.")] }), final([thought("c1", "I won't do this.")])]);
    expect(r.stream.hasContent).toBe(false);
    expect(r.stream.fullText).toBe("");
    expect(r.stream.reasoning).toBe("I won't do this.");
  });

  it("takes a thought that only the final item carries", async () => {
    const r = await play([update({ cursor: cur("a1"), messages: [said("a1", "ok")] }), final([thought("c1", "Easy one."), said("a1", "ok")])]);
    expect(r.stream.reasoning).toBe("Easy one.");
    expect(r.thoughts).toBe("Easy one.");
  });

  it("is empty for a turn without one, and ignores the other Progress messages", async () => {
    const r = await play([
      update({ messages: [{ text: "Coding and executing", author: "bot", messageId: "p", messageType: "Progress", contentType: "Code", addToChainOfThought: true }] }),
      update({ cursor: cur("a1"), messages: [said("a1", "pong")] }),
      final([said("a1", "pong")]),
    ]);
    expect(r.stream.reasoning).toBe("");
    expect(r.thoughts).toBe("");
    expect(r.stream.fullText).toBe("pong");
  });

  it("recognises the message by its contentOrigin", () => {
    expect(isChainOfThought(thought("c", "x"))).toBe(true);
    expect(isChainOfThought(thought("c", ""))).toBe(false);
    expect(isChainOfThought({ contentOrigin: "EarlyProgress", text: "Hang on a sec…" })).toBe(false);
    expect(isChainOfThought(said("a", "x"))).toBe(false);
  });
});
