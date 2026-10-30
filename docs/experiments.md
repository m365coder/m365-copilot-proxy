# Experiments — a runnable catalog

Reusable experiments for this project. Each is a **hypothesis + an exact way to
run it + how to read the result**. This is the "press go" layer; the messy
thinking lives in [`hypotheses.md`](hypotheses.md), confirmed facts in
[`m365-copilot-api.md`](m365-copilot-api.md).

> **Science discipline (see `AGENTS.md`):** change ONE variable, give it a
> `--label`, record the number. Don't trust small differences at `n=1` — use
> `--repeat`. **Run comparative experiments on a RESTED account** — degradation
> (api doc §7) makes `Disengaged` look like a format/prompt failure and poisons
> A/Bs. Pace requests; if `Disengaged`/empty spikes across *fresh* conversations,
> stop and wait ~15 min.

## Setup

```sh
pnpm build && pnpm run proxy 4141          # one shell; add env flags per experiment
# probes: M365_NO_INTERACTIVE=1 CHROMIUM_PATH=$(which chromium) node scripts/<probe>.mjs
# bench:  node scripts/bench/run.mjs --base-url http://localhost:4141/v1 --model <id> --label <name>
```

Bench scorecards land in `scripts/bench/out/<label>-<ts>.json` — diff them.

---

## A. Tool-call compliance — the headline problem (baseline: 0/5, §8.12)

The model prose-hallucinates instead of emitting tool JSON. Goal: any config that
gets a non-zero `SOLVED` / tool-call rate. **Always diff against the magic baseline.**

### E-C0 — Re-baseline (run first, on a rested account)
- **Why:** the `0/5` was measured while degraded; confirm it holds when fresh.
- **Run:** `node scripts/bench/run.mjs --model m365-copilot --label baseline --repeat 2`
- **Read:** `SOLVED %` + outcome mix. If still ~0 with few disengages → the prose
  failure is real (not throttle). If disengages vanish but prose stays → confirms
  the two failure modes are independent.
- **Cost:** ~25–50 msgs.

### E-C1 — Fenced format vs JSON  ✅ RESOLVED (JSON deleted)
- **Hypothesis (H4):** the model emits ` ```bash `/` ```edit ` blocks far more readily
  than `{"tool":...}` JSON. **Confirmed and then some** — fenced is now the *only* format;
  the JSON format was removed (it scored 0/5 on real agentic tasks). See E-C1b for the win.
- Format per tool: fence info-string = tool name, scalar args as `key: value` header lines,
  one free-form body arg as the fence body, an `old`/`new` pair as an aider-style
  `SEARCH/REPLACE` diff. Code: `packages/core/src/fenced.ts`, wired via `tools.ts`/`agent.ts`.
- **Known weakness:** a `write_file` body that itself contains a ` ``` ` fence can't be
  carried unambiguously — but in practice the model routes file writes through ```` ```bash ````
  heredocs (shell-routing), which sidesteps it.

### E-C1b — Shell-routing ⭐ THE WIN (June 14, hypotheses F12)
- **Result:** fenced format + the proxy's shell-first framing turns 0/5 into real
  multi-turn agent loops (verified 9-tool-call `fix-bug` solve on a NEUTRAL harness
  prompt). The model won't "be an agent" but will write ```bash; the proxy routes
  that block to the harness's shell tool (any name) and executes it. **Shipped as the
  default** (no env flag — fenced + shell-routing is the only path).
- **Run (winning config):**
  ```sh
  pnpm build && pnpm run proxy 4141   # fenced + shell framing are the default
  node scripts/bench/run.mjs --model m365-copilot --label tier1-neutral --tasks fix-bug,count-lines
  ```
- **New bench knobs (this session):** `--system <file>` / `BENCH_SEED=ls|cat`; prompt
  hypotheses live in `scripts/bench/prompts/p*.txt` (p0 neutral … p8/p9 bash-elicitation).
- **Read:** any multi-turn ```bash loop ending in a verifier pass. **Cost:** ~10 msgs
  but **2 threads** — mind F13 thread-throttle; keep runs small and spaced.
- **Caveat:** fakeable create-tasks (count-lines, fizzbuzz) still hallucinate; unfakeable
  ones (fix-bug, find-needle, edit-config) solve. See hypotheses §9 "Remaining gap".

### E-C2 — Task-type sensitivity
- **Hypothesis:** fakeable tasks (fizzbuzz, count-lines) hallucinate success;
  unfakeable tasks (find-needle, fix-bug, edit-config) disengage. Compliance
  depends on whether the model *can* fake an answer.
- **Run:** `--tasks fizzbuzz,count-lines --label fakeable` vs
  `--tasks find-needle,fix-bug,edit-config --label unfakeable`.
- **Read:** outcome mix per group. Confirms the §8.12 pattern; tells us whether
  "force a tool" framing should target fakeable tasks specifically.
- **Cost:** ~15 msgs.

### E-C3 — Anti-hallucination framing
- **Hypothesis:** a system/agent prompt that asserts the model has *no* prior
  knowledge of the sandbox ("the filesystem is unknown to you; you MUST inspect it;
  any claim about a file's contents without a tool_response is a hard error")
  lowers prose-hallucination.
- **Build:** variant of `getAgentInstructions` (or pass via the harness system
  prompt in `scripts/bench/run.mjs` `SYSTEM`).
- **Run:** bench A/B.  **Read:** SOLVED + prose-count delta.  **Cost:** ~25 msgs.

### E-C4 — `tool_choice` enforcement
- **Hypothesis:** sending `tool_choice:"required"` raises tool-call rate on
  actionable tasks (vs the false-bash foot-gun the docs found on prose questions —
  F3 — which the bench's tasks avoid, since all are actionable).
- **Run:** the bench always sends tools; add a flag to set `tool_choice` in the
  request, A/B.  **Read:** SOLVED delta.  **Cost:** ~25 msgs.

### E-C5 — Model comparison (agent path)
- **Hypothesis:** `gpt-5.5` (or another `*_Chat` tone) complies better than `magic` on the
  agent path. (Claude is **not** testable here — the agent forces GPT / disengages,
  §5; it's plain-chat only.)
- **Run:** `--model m365-copilot --label magic` vs `--model gpt-5.5 --label gpt55`
  vs `--model gpt-5.4-quick --label gpt54chat`. (Not `quick`: it now aliases `gpt-5.5`,
  so it would repeat the second arm — hypotheses §19.)
- **Read:** SOLVED per model.  **Cost:** ~25 msgs/model.

---

## B. Throttle / degradation (api doc §7)

### E-T1 — Characterise the throttle  (`scripts/throttle-probe.mjs`)
- **Hypothesis (H8.20):** account degradation is request-rate (RPM) driven with a
  recovery window — staying under some RPM avoids it.
- **Run:** `M365_NO_INTERACTIVE=1 CHROMIUM_PATH=$(which chromium) node scripts/throttle-probe.mjs --rpm 30 --max 25 --recover`
  Sweep `--rpm 10 / 30 / 60 / 120`.
- **Read:** request index where `Disengaged`/empty onset begins per RPM, and the
  recovery delay. Output → a safe client-side pacing config (requests/min the proxy
  should self-limit to).
- **Cost:** up to `--max` msgs/run (bursty by design).

### E-T2 — Is degradation grounding-path dependent?
- **Hypothesis (H8.8):** tenant-graph-grounded turns degrade sooner than ungrounded
  ones. (Pairs with E-O1's search toggle.)
- **Run:** throttle-probe with `plugins:[]` vs default Bing; compare onset.
- **Cost:** bursty.

### E-T3 — Does a fresh login clear throttle, or is it just idle time?  (`scripts/throttle-recovery-ab.mjs`)
- **Hypothesis (H-R1, §11):** re-auth does NOT clear thread-rate throttle (it's
  `oid`-keyed); F13's "fresh login recovered it" was confounded by ~4 min of rest.
  If true, auto-reauth is pure downside — it carries all the F25 login-fingerprint
  flag-risk for zero throttle benefit.
- **Run (needs a DEGRADED account):**
  `CHROMIUM_PATH=$(which chromium) node scripts/throttle-recovery-ab.mjs --rounds=12 --gap=45`
  Add `--induce=20` to force degradation first (burns ~20 threads). Within-episode
  two-token control: OLD (cache) vs NEW (fresh login), same `oid`, alternated probes.
- **Read:** the printed `verdict` — `H-R1_CONFIRMED_TOKEN_IRRELEVANT` (both recover
  together) vs `H-R1_REJECTED_TOKEN_IS_LEVER` (NEW recovers ≥2 rounds earlier).
- **Cost:** ~2 threads/round + one full login. Refuses to run on a rested account.

---

## C. License-free capability unlocks (optionsSets — §8.1)

All run with `scripts/_probe-chat.mjs` overrides; no license needed.

### E-O1 — Web-search toggle (H8.9)
- **Run:** a probe sending `plugins:[]` + `optionsSets:["nosearchall"]` vs default;
  same fresh-fact question.  **Read:** `InternalSearchQuery` frames + latency only
  in the search-on arm.  **Cost:** ~4 msgs.

### E-O2 — Memory / custom-instructions (H8.14)
- **Run:** `optionsSets:["add_custom_instructions","update_memory_plugin",
  "enable_inferred_memory_read"]`; turn 1 plant a code word, NEW conversation ask
  for it.  **Read:** recalled across conversations vs not.  **Cost:** ~3 msgs.

### E-O3 — Image input (H8.10)
- **Run:** POST a PNG to the substrate `UploadFile` endpoint (see PyRIT), attach
  `messageAnnotations`, ask "what's in this image?".  **Read:** pixel-level vision
  vs filename echo.  **Cost:** ~3 msgs.  **Needs:** the upload-flow probe.

### E-O4 — Code interpreter regression (already wired)
- **Run:** `node scripts/code-interpreter-probe.mjs` — SHA-256 oracle.
- **Read:** correct digest = real Python.  **Cost:** 1 msg. Use as a smoke test
  that the agent-less optionsSets path still works after changes.

---

## D. Limits / regression (confirmed once — re-run to catch M365 changes)

### E-L1 — Input ceiling / retrieval depth (F9/F10)
- **Run:** `node scripts/input-size-bisect.mjs` ; `code-interp-egress`-style needle
  tests.  **Read:** still ≥500k-token accept, dispersed-fact recall, benign size
  never disengages.  **Cost:** ~6 msgs.

### E-L2 — Output ceiling (F9)
- **Run:** essay word-target probe (output-ceiling-probe is integer-only; use an
  incompressible task).  **Read:** still wraps ~3k tokens.  **Cost:** ~2 msgs.

### E-L3 — Disengaged tool-count threshold (F6)
- **Hypothesis:** find the tool-count at which a *clean* (non-jailbreak) tool block
  disengages. Note: §8.12 showed tool *count* doesn't change the 0-compliance, but
  the disengage threshold is a separate axis.
- **Run:** `frame-dump-probe.mjs --many-tools` style, escalating tool count, watch
  `dea_violation` + `messageType`.  **Cost:** ~1 msg/step.

---

## E. Claude Sonnet 5 (paid scenario — §21)

Sonnet 5 is the `Claude_Sonnet_5` tone since 2026-10-06; `Claude_Sonnet` on the paid scenario is now
Sonnet 5.5, which has the same sandbox and a metered budget (80/day) — §26.

### E-S1 — What gates Sonnet 5's own sandbox tools? (F36)
- **Hypothesis:** some client-side field turns `bash_tool`/`create_file` off. Falsified so far for
  optionsSets, plugins, variants and `allowedMessageTypes`; next candidates: `gptDefinitions`,
  `clientOverrides.capabilities`, a different `clientInfo.clientPlatform`.
- **Run:** `node scripts/sonnet5-native-tools-probe.mjs pwd-proxy,pwd-none,pwd-bare,pwd-noprogress`
  (add a cell per candidate; `TONE=Claude_Sonnet` probes Sonnet 5.5 and spends its budget). **Read:** `native tool calls: N` and whether the reply says
  `/home/claude`. **Cost:** 1 fresh thread per cell — needs a paid seat.

### E-S2 — Framing sweep for Sonnet 5 (F37, F43)
- **Done:** `relay_inline` (the relay note inside the first `<user>` block) is falsified: 1/20 vs
  relay's 9/10 in the same session (F43). The variant was removed; F43 gives its prompt layout.
- **Hypothesis under test next (H-opening):** Sonnet 5 settles provenance from how the real message
  opens. Needs two single-change variants registered in `fenced.ts`: (a) relay with the harness
  block moved first, note still untagged; (b) the `<user>`-tagged note + task first, harness block
  after. Predicted: (a) fails like relay_inline, (b) works like relay.
- **Run:** `ARMS="<a> default <b> <a>" TAG=s5c bash scripts/bench/sonnet5-sweep.sh` (it starts
  a proxy per arm), alternating so each variant has a concurrent relay control.
  (`scripts/bench/phase-sweep.sh` now does the proxy and archiving too, and
  `scripts/bench/analyze-arms.mjs` reads the result back; see scripts/bench/README.md.)
  **Read:** SOLVED per arm, then `ChainOfThoughtSummary` frames in the archived
  `~/.config/opencode-m365/s5-sweep/<tag>-<n>-<arm>-frames` for *why* each miss refused.
  **Cost:** ~12 fresh threads per arm; ~190 threads in one day tripped `PerUserThrottled` (F40),
  so run at most ~8 arms per account per day.

---

## F. Account-level gates (Opus 4.8 on non-premium accounts — §47)

M365 serves Opus 4.8 agent-less to one non-premium account and not to another, and nothing we send separates
them (F114, F116). These recipes check an account, and compare two.

### E-A1 — Does M365 serve the tone to this account yet? (H47a)
- **Run** (as each account, once a day): `M365_DEBUG=1 node scripts/tone-frames-probe.mjs` (default:
  `Claude_Opus` on the included scenario). **Read:** `contentOrigin` `DeepLeo` and the self-ID, or
  `BotConnection` with `finalItem.result.value: "InternalError"`. **Cost:** 1 turn.

### E-A2 — What separates two accounts, without a model turn?
- **Run** (as each account): `node scripts/entitlements-probe.mjs > ent.json` (licences, plans, groups) and
  `node scripts/region-probe.mjs` (backend region). **Read:** diff the two `ent.json`; the region lines.
  **Cost:** no turn. The outputs name the account's groups: keep them out of the repo.

### E-A3 — Does the web client's own request change the answer? (F116)
- **Run** (as each account, from its checkout): `OUT=<dir> PORT=<port> TOKEN=<random> node
  scripts/webclient/driver.mjs`, then through `scripts/webclient/drive.sh <port> <token>`: sign in
  (`h.bg("login", …)`, see the driver's header), `h.html("chat-page")`, `h.routeTone("Claude_Opus")`,
  `h.bg("send", () => h.send("<self-ID prompt>"))`, `quit`. **Read:**
  `node scripts/webclient/analyze.mjs <dir>` (the client's frame, the tone sent, the final result) and
  `node scripts/webclient/diff.mjs <dir A> <dir B>` (request, ECS configs, page flags, picker).
  **Cost:** 1 turn per account; the first sign-in per account needs its TOTP secret (F117).

---

## G. The Excel agent's models over AugLoop (§48, API doc §13)

The `augloop-*` models: a slot (`agentModelType`) and the flights that name its model (F120). Their turns
count in the same turn budget; Opus 5.5 there has a limit of its own, 40 turns a day (F124).

### E-X1 — Which model does each slot serve? (F120, F121)
- **Run:** `M365_DEBUG=1 node scripts/augloop-probe.mjs augloop-excel-claude-opus-5.5,augloop-excel-claude-sonnet-5.5,…`
  (or `augloop-excel:<agentModelType>` for a slot without a model ID). **Read:** `origin=AugLoop/<model>`, the model
  the agent's saved state names; never the reply's self-ID. **Cost:** 1 turn per model.

### E-X2 — A slot's flights, from Excel for the web
- **Run:** the web-client driver (E-A3's setup), signed in through `www.microsoft365.com/launch/excel?auth=2`;
  attach a `page.on("websocket")` logger for `augloop` URLs through `eval`; "Create blank workbook", open
  Copilot ("Chat with Copilot"), set the mode menu (`data-testid="EditModeToggle"`) to "Chat only", pick the
  model, send one message. **Read:** the session init's `clientMetadata.flights` (`…<slot>ModelId`,
  `…MeteringUnit.<slot>`, `…AgentMode.<slot>Enabled`) and the chat signal's `agentModelType`.
  **Cost:** 1 turn; the workbook is a new file in the account's OneDrive, and a tenant that requires labels
  asks for one before the pane edits it.

### E-X3 — Is the label needed once SetClp is answered? (F126: no, n=1)
- **Run:** `M365_AUGLOOP_LABEL=none node scripts/augloop-probe.mjs augloop-excel-gpt-6-sol`. **Read:**
  `result=Success`, or `error_clp_set_failed` / another status in the debug log. **Cost:** 1 turn.

### E-X4 — Framing sweep on the Excel agent (F125)
- **Run:** `M365_AVOID_THROTTLING=1 TAG=<tag> PHASES='G@MODEL=augloop-excel-gpt-6-sol:default,relay_batch,baseline,pi=fix-bug'
  nix develop --command bash scripts/bench/phase-sweep.sh`, then `analyze-arms.mjs` on the archive. **Read:**
  solved, turns per task, and the sandbox column (turns the agent ran its own tools). **Cost:** ~2 turns per
  task; `augloop-excel-gpt-6-sol` has no limit seen, `augloop-excel-claude-opus-5.5` stops the sweep after 40 turns a day (F124).

### E-X5 — The conversation across a deleted chat and a new connection (F129)
- **Run:** `M365_DEBUG=1 RECONNECT_BEFORE=2 TURNS='Remember the codeword <word>. Reply with just OK.|||What is
  17 + 25?|||What was the codeword?|||Say the codeword once more.' node scripts/augloop-probe.mjs augloop-excel-gpt-6-sol`,
  then `node scripts/augloop-history.mjs`. **Read:** turns 3 and 4 recall the codeword; the debug log's `sends=`
  (the transcript only on turn 3), `The server moved the conversation`, and `Deleted chat` lines; no chat left
  in history. With `M365_SAVE_HISTORY=1`, turn 3 sends only its message, the conversation keeps its ID, and the
  chat stays (delete it with `augloop-history.mjs --app excel --delete <id>`). **Cost:** 4 turns, no limit seen.

### E-X6 — Word's conversation across a deleted chat (§49)
- **Run:** `M365_DEBUG=1 SETTLE_BEFORE=2 TURNS='Remember the codeword <word>. Reply with just OK.|||What is
  17 + 25?|||What was the codeword?' node scripts/augloop-probe.mjs augloop-word-gpt-5.6-sol-quick`, then the
  same with `RECONNECT_BEFORE=2` in place of `SETTLE_BEFORE=2`, then `node scripts/augloop-history.mjs --app word`.
  **Read:** whether turn 3 recalls the codeword after the chat's deletion (same connection, then a new one),
  the debug log's `Deleted chat` and `The server moved the conversation` lines, and what is left in history.
  **Cost:** 6 turns.

### E-X7 — Which IDs Word's agent takes, and which model each Excel slot serves (§49 F137)
- **Run:** `node scripts/augloop-word-tone-probe.mjs [MODELID[:LEVEL] …]` and
  `node scripts/augloop-excel-tone-probe.mjs [SLOT …]`; `PROMPT='…'` asks something else (a knowledge cutoff,
  say). **Read:** each row against the made-up control. Word takes any ID, so read its self-ID's family and
  apostrophe (F137); Excel names the served model in the saved state. **Cost:** one turn per ID or slot,
  and one for the control.

---

## Adding an experiment

1. State the hypothesis + falsification criterion in `hypotheses.md`.
2. Add a runnable recipe here (commands + readout + cost).
3. Reuse `scripts/_probe-chat.mjs` (qualitative) or `scripts/bench/` (quantitative).
4. Record the result back in `hypotheses.md` with sample size + evidence pointer.
