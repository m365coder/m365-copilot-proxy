// Word's Copilot agent over AugLoop (API doc §13, hypotheses §49).
//
// Word for the web's Copilot pane runs on the shared OfficeAgent workflow. A turn
// is one `OfficeAgentSignal` carrying the query and an AppContext annotation that
// names the model (`selectedModel`, a model ID from the model table in Word's
// client) and the mode: Chat only is `interactionMode:
// "app_read_only"`, so the agent doesn't edit the document. The server keeps the
// conversation: later turns send its ID and only their own message. The agent
// reads the document through a function it asks the client to run,
// `word.get_active_document_html`; there is no document, so the proxy answers with
// an empty one. The answer comes as `OfficeAgentAnnotation`s: `Progress`, then
// `Chat` with the whole message so far, `status: "completed"` at the end.
import { randomUUID } from "node:crypto";
import type { CopilotStream } from "./copilot.js";
import { WORD_AGENT_FLIGHTS, wordAgentModelOf, type WordAgentModel } from "./augloop-models.js";
import {
  AugLoopConnection,
  AugLoopTurnStream,
  ChatDeleter,
  SESSION_PROTOCOL,
  TURN_SILENCE_MS,
  WORD_CLIENT,
  outcomeOfStatus,
  resolveSensitivityLabel,
  typeHeader as H,
  type AugLoopConversationOptions,
  type AugLoopOutcome,
  type Json,
  type SensitivityLabel,
} from "./augloop.js";
import { createLogger, trunc } from "./log.js";
import { recordThrottle, recordTurn } from "./turn-budget.js";

const log = createLogger("augloop");

const OA = "AugLoop_OfficeAgent_";
const SIGNAL = `${OA}OfficeAgentSignal`;

/** Chat only in Word's mode menu ("Copilot will respond in chat only"). */
const CHAT_ONLY = "app_read_only";

/** What `word.get_active_document_html` gets back: Word's empty document, without its properties (author, dates). */
export const EMPTY_DOCUMENT_HTML =
  '<html><head><word-document-view page-count="1"><word-document-settings change-tracking-mode="Off"></word-document-settings></word-document-view></head>' +
  '<body><div class="WordSection1"><p><span data-user-selection=""></span></p></div></body></html>';

/**
 * The chat signal for one turn, as Word for the web sends it in Chat only mode.
 * Where it differs from Word's: no document (`FileUrl`), and no enterprise
 * search, as in the Excel agent's signal, so that the agent doesn't ground a
 * coding conversation in the account's mail and files. Exported for tests.
 */
export function buildWordSignal(opts: {
  model: WordAgentModel;
  query: string;
  requestId: string;
  signalId: string;
  conversationId: string;
  sensitivityLabel: SensitivityLabel | null;
  cv: string;
  contextId: string;
  messageId: string;
}): Json {
  const label = opts.sensitivityLabel
    ? { xMetadata: { conversationSensitivityLabel: JSON.stringify({ id: opts.sensitivityLabel.id, ...(opts.sensitivityLabel.priority !== undefined ? { priority: opts.sensitivityLabel.priority } : {}) }) } }
    : {};
  return {
    cv: opts.cv,
    ops: [{
      parentPath: ["Signal", SIGNAL],
      items: [{
        id: opts.signalId,
        body: {
          conversationId: opts.conversationId,
          requestId: opts.requestId,
          content: {
            messageType: "Chat",
            query: opts.query,
            messageAnnotations: [{
              id: "app-context",
              messageAnnotationType: "AppContext",
              text: "",
              messageAnnotationSource: "ClientAnnotated",
              messageAnnotationMetadata: {
                app: "Word",
                selectedModel: opts.model.modelId,
                ...(opts.model.reasoningEffort ? { reasoningEffort: opts.model.reasoningEffort } : {}),
                webSearch: true,
                enterpriseSearch: false,
                interactionMode: CHAT_ONLY,
                isConversationHandoff: false,
              },
            }],
            H_: H(`${OA}OfficeAgentMessage`, [`${OA}OfficeAgentRequest`]),
          },
          capabilities: { skills: [], connectors: [] },
          isNewConversation: !opts.conversationId,
          H_: H(SIGNAL, ["AugLoop_MicrosoftAiPlugins_PluginFunctionInputSignal", "AugLoop_Signals_Signal"]),
          ...label,
        },
        contextId: opts.contextId,
      }],
      H_: H("AugLoop_Signals_SignalOperation", ["AugLoop_Core_Operation"]),
    }],
    H_: H(`${SESSION_PROTOCOL}SyncMessage`),
    messageId: opts.messageId,
  };
}

/** The message a turn's `Chat` annotations are on, and the answer before it. */
export interface WordAnswerState {
  messageId: string | null;
  base: string;
}

/**
 * Fold one `OfficeAgentAnnotation`'s content into the turn's stream. Each `Chat`
 * annotation carries its whole message so far; a new `messageId` is a new
 * paragraph. Returns the outcome once the turn has ended, else null. Exported
 * for tests.
 */
export function applyWordAnswer(out: AugLoopTurnStream, state: WordAnswerState, content: Json): AugLoopOutcome | null {
  const status: string = content.status ?? "";
  // A stop ends as Disengaged, `canceled` ("Operation cancelled"), which isn't part of the answer.
  if (content.messageType === "Disengaged" && (status === "canceled" || status === "cancelled")) return { kind: "canceled" };
  if (typeof content.text === "string" && content.text.trim() && (content.messageType === "Chat" || content.messageType === "Disengaged")) {
    const id: string | null = content.messageId ?? state.messageId;
    if (id !== state.messageId) {
      state.messageId = id;
      state.base = out.answer ? `${out.answer}\n\n` : "";
    }
    out.finalAnswer(state.base + content.text.replace(/\s+$/, ""));
  }
  if (content.messageType === "Disengaged") {
    return { kind: "disengaged", result: { value: "Disengaged", errorCode: String(content.errorCode ?? (status || "disengaged")) } };
  }
  if (content.messageType === "RequireHumanInTheLoop") {
    return { kind: "fail", message: "Word's agent asked for an approval, which this client can't give" };
  }
  if (status === "completed") return { kind: "done", result: { value: "Success" } };
  if (status === "failed" || status === "error") {
    const code = typeof content.errorCode === "string" ? content.errorCode : status;
    const mapped = code.startsWith("error_") ? outcomeOfStatus(code) : null;
    return mapped ?? { kind: "done", result: { value: "InternalError", errorCode: code, ...(content.text ? { message: String(content.text) } : {}) } };
  }
  return null;
}

/**
 * One conversation with Word's agent. Turns go one at a time over a connection
 * opened on the first and reused while it lives; the server keeps the
 * conversation, found by its ID. A temporary chat is deleted from Copilot
 * history after its turns (ChatDeleter).
 */
export class WordAgentConversation {
  private conversationId = "";
  private connection: AugLoopConnection | null = null;
  private readonly deleter: ChatDeleter | null;
  private chain: Promise<unknown> = Promise.resolve();
  private turns = 0;
  private adjustment: Json | null = null;

  constructor(private readonly opts: AugLoopConversationOptions) {
    this.deleter = opts.deleteChats ? new ChatDeleter(opts.deleteChats) : null;
  }

  get turnCount(): number {
    return this.turns;
  }

  /** The last turn's `OfficeAgentModelSelectionAdjustedAnnotation`, if the server sent one: it changed the model. */
  get lastModelAdjustment(): Json | null {
    return this.adjustment;
  }

  /** The server's ID for this conversation, once it has assigned one. */
  get serverConversationId(): string {
    return this.conversationId;
  }

  close(): void {
    this.connection?.close();
    this.connection = null;
  }

  /** Resolves once the chat's deletion from history, if one is due, has run. A script awaits it before exiting. */
  historySettled(): Promise<void> {
    return this.deleter?.idle() ?? Promise.resolve();
  }

  /** Send one user message and stream the answer. Resolves once the message is sent. */
  chat(text: string, model: string, signal?: AbortSignal): Promise<CopilotStream> {
    const m = wordAgentModelOf(model);
    if (!m) return Promise.reject(new Error(`${model} is not a model of Word's agent`));
    let release!: () => void;
    const turnDone = new Promise<void>((r) => { release = r; });
    const previous = this.chain;
    this.chain = previous.then(() => turnDone);
    return previous
      .then(() => this.startTurn(text, model, m, signal, release))
      .catch((err) => { release(); throw err; });
  }

  /** Take the conversation ID the server names in its answers. */
  private follow(conversationId: string | undefined): void {
    if (!conversationId || conversationId === this.conversationId) return;
    if (this.conversationId) log.info(`The server moved the conversation from ${this.conversationId} to ${conversationId}`);
    this.conversationId = conversationId;
  }

  private async open(signal?: AbortSignal): Promise<AugLoopConnection> {
    const c = await AugLoopConnection.open(WORD_AGENT_FLIGHTS.map((f) => `${f};`).join(""), await this.opts.getToken(), signal, WORD_CLIENT, this.opts.onFrame);
    this.connection = c;
    return c;
  }

  private async ensureConnection(signal?: AbortSignal): Promise<AugLoopConnection> {
    const c = this.connection;
    if (!c || !c.alive) {
      c?.close();
      return this.open(signal);
    }
    if (c.tokenExpiring) {
      try {
        await c.provisionToken(await this.opts.getToken());
      } catch (err: any) {
        log.info(`Token re-provision failed (${err.message}) — reconnecting`);
        c.close();
        return this.open(signal);
      }
    }
    return c;
  }

  private async startTurn(text: string, model: string, m: WordAgentModel, signal: AbortSignal | undefined, release: () => void): Promise<CopilotStream> {
    if (signal?.aborted) throw new Error("aborted");
    await this.deleter?.turnStarting();
    const label = await resolveSensitivityLabel(this.opts.getGraphToken);
    const c = await this.ensureConnection(signal);

    const turn = this.turns++;
    this.adjustment = null;
    recordTurn(); // the throttle-avoidance bucket counts what the next line counts (turn-budget.ts)
    log.info(`Chat turn ${turn}: model=${model}, selectedModel=${m.modelId}${m.reasoningEffort ? `, reasoningEffort=${m.reasoningEffort}` : ""}, cid=${this.conversationId || "new"}, text=${JSON.stringify(trunc(text, 200))}`);

    const out = new AugLoopTurnStream();
    out.turnCount = this.turns;
    out.contentOrigin = `AugLoop/word:${m.modelId}`;
    const requestId = randomUUID();
    out.messageId = requestId;
    const signalId = randomUUID();
    const signalMessageId = c.nextMessageId();
    const answer: WordAnswerState = { messageId: null, base: "" };
    let lastFrame = Date.now();
    let stopped = false;
    let unlisten = () => {};

    const finish = (err?: Error) => {
      if (out.finished) return;
      clearInterval(silence);
      unlisten();
      signal?.removeEventListener("abort", onAbort);
      const status = out.result?.value ?? (err ? `error: ${err.message}` : stopped ? "stopped" : "no status");
      log.info(`Turn ${turn} ended: ${status}${out.result?.errorCode ? ` (${out.result.errorCode})` : ""}, ${out.answer.length} chars`);
      this.deleter?.turnEnded(this.conversationId);
      out.finish(err);
      release();
    };

    const silence = setInterval(() => {
      if (Date.now() - lastFrame > TURN_SILENCE_MS) finish(new Error(`AugLoop: no reply from Word's agent in ${TURN_SILENCE_MS / 60_000} min`));
    }, 10_000);
    silence.unref?.();

    const onAbort = () => {
      if (stopped || out.finished) return;
      stopped = true;
      log.info("Aborted — cancelling Word's agent's turn");
      try { c.cancelSignal(SIGNAL, signalId); } catch {}
      setTimeout(() => finish(), 3_000).unref?.();
    };

    const onAnswer = (b: Json) => {
      this.follow(b.conversationId);
      const outcome = applyWordAnswer(out, answer, b.content ?? {});
      if (!outcome) return;
      if (outcome.kind === "canceled") { finish(); return; }
      if (outcome.kind === "fail") { finish(new Error(outcome.message)); return; }
      out.result = outcome.result;
      if (outcome.kind === "disengaged") out.messageType = "Disengaged";
      if (outcome.result.value !== "Success") log.info(`Turn result: ${outcome.result.value}${outcome.result.errorCode ? ` (${outcome.result.errorCode})` : ""}`);
      if (outcome.result.value === "Throttled") recordThrottle(Date.now(), outcome.result.errorCode);
      finish();
    };

    unlisten = c.listen((msg) => {
      if (msg.H_?.T_ === "closed") { finish(out.answer ? undefined : new Error("AugLoop connection closed during the turn")); return; }
      if (msg.H_?.T_ === `${SESSION_PROTOCOL}ErrorResponse` && msg.messageId === signalMessageId) {
        finish(new Error(`AugLoop refused the chat signal: ${msg.code ?? ""} ${msg.error ?? ""}`.trim()));
        return;
      }
      const at: string = msg.annotationType ?? "";
      if (!at.startsWith(OA)) return;
      for (const item of (msg.ops ?? []).flatMap((o: Json) => o.items ?? [])) {
        const b: Json = item.body ?? {};
        const wec = b.workflowExecutionCorrelation;
        if (at === `${OA}OfficeAgentAnnotation`) {
          if (b.requestId && b.requestId !== requestId) continue;
          lastFrame = Date.now();
          onAnswer(b);
        } else if (at === `${OA}OfficeAgentInvokeFunctionAnnotation` && wec) {
          if (b.invokeMetadata?.requestId && b.invokeMetadata.requestId !== requestId) continue;
          lastFrame = Date.now();
          const fn = `${b.namespace ?? ""}.${b.functionName ?? ""}`;
          const readsDocument = fn === "word.get_active_document_html";
          if (!readsDocument) log.info(`Word's agent called ${fn} (${trunc(String(b.argumentsJson ?? ""), 200)}) — answering that there is no document`);
          const result = readsDocument ? EMPTY_DOCUMENT_HTML : `This client has no document open, so ${fn} has nothing to return.`;
          c.respond(wec, { items: [{ contentType: "text/plain", content: result }], H_: H(`${OA}OfficeAgentInvokeFunctionResultResponse`, []) });
        } else if (at === `${OA}OfficeAgentSetClpAnnotation` && wec) {
          lastFrame = Date.now();
          c.respond(wec, { success: true, H_: H(`${OA}OfficeAgentSetClpResponseObject`, []) });
        } else if (at === `${OA}OfficeAgentModelSelectionAdjustedAnnotation`) {
          this.adjustment = b;
          log.info(`Word's agent adjusted the model selection: ${trunc(JSON.stringify(b), 400)}`);
        } else {
          log.info(`Unhandled ${at.slice(OA.length)}: ${trunc(JSON.stringify(b), 300)}`);
        }
      }
    });

    if (signal) signal.addEventListener("abort", onAbort, { once: true });

    c.send(buildWordSignal({
      model: m,
      query: text,
      requestId,
      signalId,
      conversationId: this.conversationId,
      sensitivityLabel: label,
      cv: c.nextCv(),
      contextId: c.nextContextId(),
      messageId: signalMessageId,
    }));
    return out.stream();
  }
}
