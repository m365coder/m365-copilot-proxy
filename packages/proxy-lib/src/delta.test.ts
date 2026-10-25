import { describe, it, expect } from "vitest";
import { transcriptStyleForVariant } from "@m365-copilot/core";
import { formatDeltaMessages } from "./handler.js";

const history = [
  { role: "user" as const, content: "fix the bug" },
  { role: "assistant" as const, content: "", tool_calls: [{ id: "c1", type: "function" as const, function: { name: "bash", arguments: '{"command":"cat calc.py"}' } }] },
  { role: "tool" as const, tool_call_id: "c1", content: "def add(a, b): return a - b" },
];

describe("formatDeltaMessages", () => {
  it("sends only the tool result by default", () => {
    const out = formatDeltaMessages(history.slice(1), history, transcriptStyleForVariant("relay_batch"));
    expect(out).toBe('<tool_response tool="bash" command="cat calc.py">\ndef add(a, b): return a - b\n</tool_response>');
  });

  it("adds relay_turn's note after a follow-up's tool results (docs §41)", () => {
    const out = formatDeltaMessages(history.slice(1), history, transcriptStyleForVariant("relay_turn"));
    expect(out.startsWith("<tool_response")).toBe(true);
    expect(out.endsWith(transcriptStyleForVariant("relay_turn").turnNote!)).toBe(true);
  });

  it("leaves the note off a follow-up that brings no tool result", () => {
    const user = [{ role: "user" as const, content: "and the tests?" }];
    expect(formatDeltaMessages(user, user, transcriptStyleForVariant("relay_turn"))).toBe("<user>\nand the tests?\n</user>");
  });

  it("pastes the result back in the user's voice under relay_echo", () => {
    const out = formatDeltaMessages(history.slice(1), history, transcriptStyleForVariant("relay_echo"));
    expect(out).toBe("Here's what your bash block printed on my machine:\n```text\ndef add(a, b): return a - b\n```");
  });
});
