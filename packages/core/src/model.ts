import { getAugLoopToken, getGraphToken, getToken } from "./auth.js";
import { getOrCreateAgent } from "./agent.js";
import { CopilotSession } from "./session.js";
import { AugLoopConversation } from "./augloop.js";
import { isAugLoopModel } from "./augloop-models.js";
import { uploadImage, type ImageAnnotation } from "./image-input.js";
import { createLogger, trunc } from "./log.js";
import { awaitTurnBudget } from "./turn-budget.js";
import type { CopilotStream } from "./copilot.js";

const log = createLogger("model");

export interface ModelSessionOptions {
  /** Pre-resolved auth token. If not provided, getToken() is called. */
  getToken?: () => Promise<string>;
  /** Token for the Excel agent's AugLoop models. If not provided, getAugLoopToken() is called. */
  getAugLoopToken?: () => Promise<string>;
  /** Whether to attempt agent resolution. Default: true. */
  useAgent?: boolean;
  /**
   * Prevent proxy conversations appearing in the M365 Copilot history sidebar.
   * Defaults to true. Set false only when persistent Copilot history is wanted.
   */
  temporaryChat?: boolean;
}

/**
 * A stateful session for running M365 Copilot.
 * Manages auth, agent resolution, and conversation continuity.
 * String in, stream out.
 *
 * The same sessionId and conversationId are reused across CopilotSession
 * reconnections so M365 finds the existing server-side conversation
 * instead of creating a new one.
 */
export class ModelSession {
  private resolveToken: () => Promise<string>;
  private resolveAugLoopToken: () => Promise<string>;
  private useAgent: boolean;
  private temporaryChat: boolean;
  private copilotSession: CopilotSession | null = null;
  /** The Excel agent's side of this conversation, when its model is an AugLoop one (augloop.ts). */
  private augloop: AugLoopConversation | null = null;
  /** Whether the last run() went to AugLoop, which takes no Copilot Studio agent. */
  private lastRunAugLoop = false;
  private cachedAgentId: string | null | undefined = undefined;

  /** Stable IDs reused across CopilotSession reconnections. */
  readonly sessionId: string = crypto.randomUUID();
  private _conversationId: string = crypto.randomUUID();
  /** Current M365 ConversationId (the throttle/Disengage state keys on this). */
  get conversationId(): string { return this._conversationId; }
  /**
   * Rotate to a FRESH conversation. A conversation that has Disengaged appears to
   * STAY Disengaged (a clean retry in the same conversation kept refusing), so the
   * Disengage-recovery retry needs a new ConversationId, not just a different prompt.
   * See docs §10 F22.
   */
  newConversation(): void {
    this._conversationId = crypto.randomUUID();
    this.augloop?.close();
    this.augloop = null;
    this.reset();
  }

  constructor(options: ModelSessionOptions = {}) {
    this.resolveToken = options.getToken ?? getToken;
    this.resolveAugLoopToken = options.getAugLoopToken ?? getAugLoopToken;
    this.useAgent = options.useAgent !== false;
    this.temporaryChat =
      options.temporaryChat ?? process.env.M365_SAVE_HISTORY !== "1";
  }

  /** Number of turns completed in this session (on whichever service it runs). */
  get turnCount(): number {
    return (this.copilotSession?.turnCount ?? 0) + (this.augloop?.turnCount ?? 0);
  }

  /** Agent id baked into the current copilotSession (undefined = no agent). */
  private currentAgentId: string | undefined = undefined;

  private createCopilotSession(agentId: string | undefined): CopilotSession {
    return new CopilotSession({
      agentId,
      sessionId: this.sessionId,
      conversationId: this.conversationId,
      temporaryChat: this.temporaryChat,
    });
  }

  /**
   * Send text to M365 Copilot and stream back the response.
   *
   * `useAgent` decides whether THIS turn attaches the tool-calling Copilot Studio
   * agent (`threadLevelGptId`). The handler passes `false` for tool-less requests
   * — the agent is only needed to coerce tool-call output — and for tool requests
   * on tones that don't serve with it (`toneUsesToolAgent`: Claude, GPT-6; with
   * the agent attached those return a dead route or, in June 2026, silently
   * routed to GPT — docs/hypotheses.md H8.6, §22). Plain chat therefore always
   * gets the real model the tone selects.
   *
   * If `signal` aborts (the HTTP client disconnects) the in-flight turn is
   * cancelled by sending M365's Stop frame, mirroring the real UI's Stop button.
   *
   * `waitForBudget` waits until the account's turn budget can afford one more
   * turn (M365_AVOID_THROTTLING; awaitTurnBudget). The caller waits before it
   * calls run(); run() waits again before the reconnect's turn. It throws
   * TurnBudgetHold rather than wait too long.
   */
  async run(
    text: string,
    model: string = "m365-copilot",
    signal?: AbortSignal,
    useAgent: boolean = true,
    images: string[] = [],
    waitForBudget: () => Promise<{ waitedMs: number }> = () => awaitTurnBudget({ signal }),
  ): Promise<CopilotStream> {
    this.lastRunAugLoop = isAugLoopModel(model);
    if (this.lastRunAugLoop) return this.runAugLoop(text, model, signal, images, waitForBudget);
    let token = await this.resolveToken();
    // Image input (#39): upload each image for this conversation; the message points at them.
    // A failed upload is said in the text, so the model doesn't answer about an image it never got.
    const annotations: ImageAnnotation[] = [];
    for (const url of images) {
      try {
        annotations.push(await uploadImage(token, this.conversationId, url));
      } catch (err: any) {
        log.info(`Image upload failed: ${err.message}`);
        text += `

(Note: an attached image could not be passed on — ${err.message.slice(0, 120)}. Don't describe it; say it didn't arrive.)`;
      }
    }
    const wantAgent = this.useAgent && useAgent;

    // Resolve agent ID lazily (persists across resets), only when wanted.
    if (wantAgent && this.cachedAgentId === undefined) {
      try {
        this.cachedAgentId = await getOrCreateAgent();
        if (this.cachedAgentId) log.info(`Using agent: ${this.cachedAgentId}`);
        else log.info("No agent available");
      } catch {
        this.cachedAgentId = null;
      }
    }
    const agentForTurn = wantAgent ? (this.cachedAgentId ?? undefined) : undefined;

    // Create (or recreate) the session when missing or when this turn's
    // agent-ness differs from the current session's — switching the agent on/off
    // changes routing, so the WS session must carry the right threadLevelGptId.
    if (!this.copilotSession || this.currentAgentId !== agentForTurn) {
      this.copilotSession = this.createCopilotSession(agentForTurn);
      this.currentAgentId = agentForTurn;
    }

    log.info(`run: model=${model}, agent=${agentForTurn ?? "none"}, turn=${this.copilotSession.turnCount}, sid=${this.sessionId}, cid=${this.conversationId}, text=${JSON.stringify(trunc(text, 200))}`);

    // Match the GUI: on the agent-less path, image gen is always available, so a
    // plain "draw me an image" produces one with no special mode (§14). Kept OFF on
    // the tool-calling agent path — image optionsSets are agent-less, and we don't
    // want the image tool competing with fenced tool-call emission. Opt out with
    // M365_NO_IMAGE_GEN for a pure-text agent-less turn.
    const turnOpts = { generateImages: !agentForTurn && !process.env.M365_NO_IMAGE_GEN, images: annotations };

    try {
      return await this.copilotSession.chat(token, text, model, signal, turnOpts);
    } catch (err: any) {
      // Session might be stale — reconnect with same IDs
      log.info("Session error, reconnecting:", err.message);
      // The reconnect is one more upstream turn, so it waits on the budget too:
      // otherwise it skips the wait and its recordTurn takes the grant of a turn
      // another request was let through for (#77). A token held through a long
      // wait may be near its end, so fetch it again.
      if ((await waitForBudget()).waitedMs > 0) token = await this.resolveToken();
      this.copilotSession = this.createCopilotSession(agentForTurn);
      this.currentAgentId = agentForTurn;
      return await this.copilotSession.chat(token, text, model, signal, turnOpts);
    }
  }

  /**
   * A turn with one of the Excel agent's models (augloop.ts). It takes no Copilot
   * Studio agent and no images; the conversation is the Excel agent's own, kept
   * for this session's conversation ID until newConversation().
   */
  private async runAugLoop(
    text: string,
    model: string,
    signal: AbortSignal | undefined,
    images: string[],
    waitForBudget: () => Promise<{ waitedMs: number }>,
  ): Promise<CopilotStream> {
    if (images.length > 0) {
      text += `

(Note: ${images.length === 1 ? "an attached image" : `${images.length} attached images`} could not be passed on — this model takes text only here. Don't describe ${images.length === 1 ? "it" : "them"}; say ${images.length === 1 ? "it" : "they"} didn't arrive.)`;
    }
    this.augloop ??= new AugLoopConversation({ getToken: this.resolveAugLoopToken, getGraphToken });
    log.info(`run: model=${model}, agent=none, turn=${this.augloop.turnCount}, sid=${this.sessionId}, cid=${this.conversationId}, via=AugLoop, text=${JSON.stringify(trunc(text, 200))}`);
    try {
      return await this.augloop.chat(text, model, signal);
    } catch (err: any) {
      if (signal?.aborted) throw err;
      // Nothing was sent (chat() resolves once the message is): reconnect and try once more.
      log.info("AugLoop error, reconnecting:", err.message);
      await waitForBudget();
      this.augloop.close();
      return await this.augloop.chat(text, model, signal);
    }
  }

  /**
   * Force a re-resolution of the tool-calling agent against the tenant and
   * reconnect if it changed. Recovers from the "deleted-agent trap": a
   * long-lived host caches its agent id for the whole process lifetime, so once
   * that agent is deleted from the tenant (e.g. by a newer-build's cleanup) it
   * would otherwise send `threadLevelGptId` at a dead bot forever and get empty
   * replies. Call this on an empty upstream reply before retrying. No-op when
   * the session isn't using an agent. Returns true if the agent id changed (and
   * the session was dropped to reconnect), so the caller can resend the original
   * prompt to the fresh agent.
   */
  async refreshAgent(): Promise<boolean> {
    if (!this.useAgent || this.lastRunAugLoop) return false;
    try {
      const fresh = await getOrCreateAgent({ forceRefresh: true });
      if (fresh !== this.cachedAgentId) {
        log.info(`Agent refreshed: ${this.cachedAgentId ?? "none"} -> ${fresh ?? "none"}`);
        this.cachedAgentId = fresh;
        this.copilotSession = null; // reconnect so the new agent id takes effect
        this.currentAgentId = undefined;
        return true;
      }
    } catch (err: any) {
      log.info(`Agent refresh failed: ${err.message}`);
    }
    return false;
  }

  /**
   * Reset the CopilotSession. Next run() reconnects with the same
   * sessionId/conversationId. Also drops the cached agent id so the next run()
   * re-resolves it (cheap cache fast-path in the normal case).
   */
  reset() {
    this.copilotSession = null;
    this.cachedAgentId = undefined;
    this.currentAgentId = undefined;
    // The next AugLoop turn reconnects; its conversation stays (newConversation drops it).
    this.augloop?.close();
  }
}
