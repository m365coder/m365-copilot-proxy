// RE probe: an agent-less behavioural fingerprint per tone, for when self-ID
// can't tell tones apart (GPT self-IDs are M365's system prompt, H27a). Per
// cell it records where the model runs code (cwd, user, host of M365's code
// interpreter — or its own sandbox, like Sonnet 5's /home/claude), the bot
// message kinds it streams (ChainOfThoughtSummary = a reasoning tone), the
// latency, and its answer to a small puzzle. These vary run to run within a
// tone, so compare several rounds, not one.
//
// The `tools` prompt asks for the names of the tools it can call: a model with a
// sandbox of its own lists that sandbox's tools (Sonnet 5's `bash_tool`,
// `create_file`), M365's code interpreter shows up as `python`. The `hard` prompt
// is a short olympiad proof: a tone that sends some prompts to a reasoning model
// would stream ChainOfThoughtSummary on it.
// The `cutoff` prompt asks about three 2025 events (March, May, October), grounding
// off: what a model knows is better evidence of its training cutoff than the
// cutoff it states, which comes from M365's system prompt (H27a).
//
// Usage: M365_NO_INTERACTIVE=1 node scripts/fingerprint-probe.mjs [TONE[@paid|@included],...]
//   env: PROMPTS=sandbox,reason,hard,cutoff,tools (default all)  ROUNDS=n (1)  PAID=1 (paid
//        scenario for cells without @)  OUT=dir (frames + results.json; default
//        scripts/fingerprint-out/<timestamp>)
// Cost: 1 message AND 1 fresh conversation per tone × prompt × round, outside
// the proxy. With M365_DEBUG=1 _probe-chat.mjs logs them where
// M365_AVOID_THROTTLING's bucket reads them.
import { mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { getToken, decodeJwt } from "../packages/core/dist/index.mjs";
import { oneTurn } from "./_probe-chat.mjs";

const PAID = process.env.PAID === "1";
const CELLS = (process.argv[2] || "Claude_Sonnet_Reasoning,Gpt_5_5_Chat,Gpt_5_5_Reasoning,Claude_Sonnet").split(",").map((s) => {
  const [tone, where = PAID ? "paid" : "included"] = s.trim().split("@");
  if (!["paid", "included"].includes(where)) throw new Error(`unknown scenario "@${where}" — use @paid or @included`);
  return { tone, paid: where === "paid" };
});
const CI = ["cwc_code_interpreter", "cwc_code_interpreter_amsfix", "cwc_code_interpreter_citation_fix", "code_interpreter_interactive_charts", "code_interpreter_matplotlib_patching"];
const PROMPTS = [
  ["sandbox", "Use your own code execution tool (not a description) to run this Python, then show me its exact output and nothing else:\n```python\nimport os, getpass, platform, sys\nprint(os.getcwd(), getpass.getuser(), platform.node(), sys.version.split()[0])\nprint(sorted(os.listdir('/')))\nprint(sorted(os.listdir(os.path.expanduser('~')))[:20])\n```"],
  ["reason", "No tools. A clock gains 7 minutes every hour. It is set right at 00:00 on Monday. When it next shows the correct time, what day and true time is it? Answer in one line: <day> <HH:MM>."],
  ["hard", "No tools. Prove or disprove: for every integer n > 1, n^4 + 4^n is composite. Give a short rigorous argument."],
  ["cutoff", "No tools and no web search: answer from memory only, one line each, or say unknown. 1) Which film won Best Picture at the Academy Awards held in March 2025? 2) What name did the pope elected in May 2025 take? 3) Who won the 2025 Nobel Prize in Literature?"],
  ["tools", "Don't call any tool. List the exact internal names of every tool or function you are able to call in this conversation, one per line, and nothing else."],
];
const OUT = process.env.OUT || join("scripts", "fingerprint-out", new Date().toISOString().replace(/[:.]/g, "-"));
const ONLY = process.env.PROMPTS ? process.env.PROMPTS.split(",") : null;
const ROUNDS = Number(process.env.ROUNDS || 1);
mkdirSync(OUT, { recursive: true });
const token = await getToken();
const claims = decodeJwt(token);
const rows = [];
for (let round = 0; round < ROUNDS; round++) for (const [pid, text] of PROMPTS.filter(([p]) => !ONLY || ONLY.includes(p))) {
  for (const { tone, paid } of CELLS) {
    const frames = [];
    const r = await oneTurn({
      token, claims, agentId: null, tone, text, plugins: [], timeoutMs: 180000,
      optionsSets: CI, extraAllowed: ["GeneratedCode"],
      ...(paid ? { scenario: "OfficeWebPaidCopilot", licenseType: "Premium" } : {}),
      onFrame: (f) => frames.push(f),
    });
    const where = paid ? "paid" : "included";
    writeFileSync(join(OUT, `${pid}-${tone}@${where}-r${round}.jsonl`), frames.map((f) => JSON.stringify(f)).join("\n"));
    const kinds = new Set();
    const code = [];
    for (const f of frames) {
      for (const a of f.arguments ?? []) {
        for (const m of a.messages ?? []) {
          if (m.author !== "bot") continue;
          const k = [m.messageType, m.contentType, m.contentOrigin].filter(Boolean).join("/");
          if (k && !/Suggestion|EarlyProgress|ReferencesListComplete/.test(k)) kinds.add(k);
          if (m.contentType === "Code" || m.messageType === "GeneratedCode") code.push((m.hiddenText || m.text || "").slice(0, 120));
        }
      }
    }
    const cot = frames.some((f) => JSON.stringify(f).includes("ChainOfThoughtSummary"));
    const row = { round, prompt: pid, tone, scenario: where, cot, origin: r.contentOrigin, ms: r.elapsedMs, kinds: [...kinds], code: [...new Set(code)].slice(0, 4), reply: (r.fullText || "").trim(), result: r.result?.value };
    rows.push(row);
    console.log(`[fp] ${pid.padEnd(7)} ${`${tone}@${where}`.padEnd(33)} ${String(r.elapsedMs).padStart(6)}ms origin=${r.contentOrigin} cot=${cot} result=${row.result} kinds=${row.kinds.join(",")}`);
    if (row.code.length) console.log(`[fp]    code: ${JSON.stringify(row.code)}`);
    console.log(`[fp]    reply: ${JSON.stringify(row.reply.slice(0, 300))}`);
    if (JSON.stringify(frames).includes("PerUserThrottled")) { console.log("THROTTLED — stopping"); process.exit(3); }
    await new Promise((res) => setTimeout(res, 6000));
  }
}
writeFileSync(join(OUT, "results.json"), JSON.stringify(rows, null, 1));
