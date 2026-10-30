// AugLoop: the Excel agent's models over the AugLoop service (API doc §13).
//
// Excel for the web's Copilot (Agent Mode) talks to `wss://augloop.office.com`,
// not to Chathub, and picks its model by a slot (`agentModelType`, augloop-models.ts).
// The proxy drives it the way that client does, with an Entra ID token for
// `https://augloop.office.com/v2/.default` that our own first-party client gets
// silently (hypotheses §48 F118), and two settings the client's UI offers:
//  - Chat only (Excel's mode menu): `documentConfig.isReadOnly` with reason 21, so
//    the agent doesn't edit the workbook. There is no workbook: the proxy answers
//    the agent's workbook script with an empty one-sheet book.
//  - A sensitivity label on the document, for tenants that require one before
//    editing: M365_AUGLOOP_LABEL (default "General Business"), matched against the
//    tenant's labels from Graph.
//
// The agent runs in a container and keeps its conversation there. The server asks
// the client to save that state after each turn and to load it when the container
// is gone, and takes the conversation so far as `messages` on each turn. The proxy
// does both, in memory, per conversation.
import WebSocket from "ws";
import { randomBytes, randomUUID } from "node:crypto";
import { appendFileSync, mkdirSync } from "node:fs";
import { join } from "node:path";
import { inflateRawSync } from "node:zlib";
import type { CopilotStream } from "./copilot.js";
import { augLoopModelOf, isClaudeAugLoopModel } from "./augloop-models.js";
import { foldStreamText } from "./session.js";
import { createLogger, trunc, FRAME_DIR } from "./log.js";
import { recordThrottle, recordTurn } from "./turn-budget.js";

const log = createLogger("augloop");

const P = "AugLoop_ExcelAgentExperimental_ExcelAgentExperimental";
const SESSION = "AugLoop_Session_Protocol_";
const MESSAGE = [`${SESSION}Message`];

// Flights every connection sends, as Excel's client does: the server assigns the
// conversation ID (without it the agent fails with `error_internal_error` before
// it starts, F120), and Excel's model metering.
const BASE_FLIGHTS = [
  "Microsoft.Office.Excel.FA000000124.AugLoop.EnableServerConversationId:True",
  "Microsoft.Office.Excel.AugLoop.EnableModelMetering:True",
];

// The annotations this client reads or answers. Not ExecutePython: it hands the
// agent's Python to the client to run (Excel for the web runs it in the browser),
// and this client runs none.
const ACTIVATED = [
  "ForbiddenAnnotation",
  "UserAllowedAnnotation",
  "OutputAnnotation",
  "RunScriptAnnotation",
  "SetClpAnnotation",
  "SaveAgentSessionStateAnnotation",
  "LoadAgentSessionStateAnnotation",
].map((a) => P + a);

/** What the agent's workbook script gets back: Excel's own reply for an empty one-sheet book. */
const EMPTY_WORKBOOK = JSON.stringify({
  activeCell: "Sheet1!A1",
  selectedRanges: "Sheet1!A1",
  activeSheet: "Sheet1",
  sheets: [{ name: "Sheet1", tables: [], charts: [] }],
  workbookNames: [],
  rulesCandidateCount: 0,
});

/** `readOnlyReason` of Excel's "Chat only" mode ("Copilot will respond in chat only"). */
const CHAT_ONLY_REASON = 21;

const USER_AGENT =
  process.env.M365_LOGIN_UA ??
  "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36";

const CONNECT_TIMEOUT_MS = 30_000;
const KEEPALIVE_MS = 30_000;
const IDLE_CLOSE_MS = 10 * 60_000;
/** A turn with no frame for this long is given up. The agent can work for minutes. */
const TURN_SILENCE_MS = 5 * 60_000;
const TOKEN_MARGIN_MS = 5 * 60_000;

function host(): string {
  return process.env.M365_AUGLOOP_HOST || "augloop.office.com";
}

const H = (T: string, B: string[] = MESSAGE) => ({ T_: T, B_: B });

// --- Frame dumps (M365_DUMP_FRAMES) -----------------------------------------------

const DUMP_FRAMES = !!process.env.M365_DUMP_FRAMES;
const TOKEN = /eyJ[A-Za-z0-9_-]{10,}(?:\.[A-Za-z0-9_-]*){2,4}/g;

/** A frame without its tokens and without the agent's saved state, which holds the user's profile. */
export function redactFrame(raw: string): string {
  return raw
    .replace(TOKEN, (m) => `<token:${m.length}>`)
    .replace(/"stateData":"([^"]*)"/g, (_m, s: string) => `"stateData":"<${s.length} chars>"`);
}

function dumpFrame(file: string, dir: "send" | "recv", raw: string) {
  if (!DUMP_FRAMES) return;
  try {
    mkdirSync(FRAME_DIR, { recursive: true });
    appendFileSync(join(FRAME_DIR, `${file}.ndjson`), JSON.stringify({ t: Date.now(), dir, frame: redactFrame(raw) }) + "\n");
  } catch {
    // best effort
  }
}

// --- Sensitivity label ----------------------------------------------------------------

export interface SensitivityLabel {
  id: string;
  priority?: number;
  name?: string;
}

/** A label as Graph's `/me/informationProtection/sensitivityLabels` lists it. */
export interface GraphSensitivityLabel {
  id: string;
  name?: string;
  displayName?: string;
  priority?: number;
  sublabels?: GraphSensitivityLabel[];
}

/** M365_AUGLOOP_LABEL: the label the agent's document carries (name, display name or ID); `none` sends none. */
export function augLoopLabelSetting(): string | null {
  const v = (process.env.M365_AUGLOOP_LABEL ?? "General Business").trim();
  return v === "" || /^none$/i.test(v) ? null : v;
}

/**
 * The wanted label among the tenant's: by ID, then by name or display name, then
 * by a name that starts with it ("General Business" → "General Business (G)"),
 * case-insensitive, sublabels included. Null when none matches.
 */
export function pickSensitivityLabel(labels: GraphSensitivityLabel[], wanted: string): SensitivityLabel | null {
  const all = labels.flatMap((l) => [l, ...(l.sublabels ?? [])]);
  const w = wanted.trim().toLowerCase();
  const names = (l: GraphSensitivityLabel) => [l.displayName, l.name].filter((n): n is string => !!n).map((n) => n.toLowerCase());
  const hit = all.find((l) => l.id.toLowerCase() === w)
    ?? all.find((l) => names(l).includes(w))
    ?? all.find((l) => names(l).some((n) => n.startsWith(w)));
  if (!hit) return null;
  return { id: hit.id, ...(typeof hit.priority === "number" ? { priority: hit.priority } : {}), name: hit.displayName ?? hit.name };
}

let labelCache: SensitivityLabel | null | undefined;

/**
 * The configured label, resolved once per process against the tenant's labels.
 * A tenant that has labels but none of that name is a configuration error, and
 * the request fails. If the labels can't be read, the turn goes without one.
 */
export async function resolveSensitivityLabel(getGraphToken: () => Promise<string | null>): Promise<SensitivityLabel | null> {
  const wanted = augLoopLabelSetting();
  if (!wanted) return null;
  if (labelCache !== undefined) return labelCache;
  let labels: GraphSensitivityLabel[];
  try {
    const token = await getGraphToken();
    if (!token) throw new Error("no Graph token");
    const res = await fetch("https://graph.microsoft.com/v1.0/me/informationProtection/sensitivityLabels", { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) throw new Error(`Graph sensitivityLabels: HTTP ${res.status}`);
    labels = ((await res.json()) as { value?: GraphSensitivityLabel[] }).value ?? [];
  } catch (err: any) {
    log.info(`Sensitivity labels unavailable (${err.message}) — the agent's document goes without one`);
    return null;
  }
  if (labels.length === 0) {
    labelCache = null;
    return null;
  }
  const label = pickSensitivityLabel(labels, wanted);
  if (!label) {
    const known = labels.flatMap((l) => [l, ...(l.sublabels ?? [])]).map((l) => l.displayName ?? l.name).join(", ");
    throw new Error(`M365_AUGLOOP_LABEL "${wanted}" is none of this tenant's sensitivity labels (${known})`);
  }
  log.info(`Sensitivity label for the agent's document: ${label.name} (priority ${label.priority ?? "?"})`);
  labelCache = label;
  return label;
}

/** Forget the resolved label. For tests. */
export function resetSensitivityLabel(): void {
  labelCache = undefined;
}

// --- What the agent's saved state says ---------------------------------------------

/** One entry of a ZIP archive, or null. Stored and deflated entries only. */
export function readZipEntry(zip: Buffer, name: string): Buffer | null {
  let eocd = -1;
  for (let i = zip.length - 22; i >= Math.max(0, zip.length - 65_557); i--) {
    if (zip.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) return null;
  const entries = zip.readUInt16LE(eocd + 10);
  let p = zip.readUInt32LE(eocd + 16);
  for (let n = 0; n < entries && p + 46 <= zip.length; n++) {
    if (zip.readUInt32LE(p) !== 0x02014b50) return null;
    const method = zip.readUInt16LE(p + 10);
    const size = zip.readUInt32LE(p + 20);
    const nameLen = zip.readUInt16LE(p + 28);
    const extraLen = zip.readUInt16LE(p + 30);
    const commentLen = zip.readUInt16LE(p + 32);
    const local = zip.readUInt32LE(p + 42);
    if (zip.toString("utf8", p + 46, p + 46 + nameLen) === name) {
      if (zip.readUInt32LE(local) !== 0x04034b50) return null;
      const start = local + 30 + zip.readUInt16LE(local + 26) + zip.readUInt16LE(local + 28);
      const data = zip.subarray(start, start + size);
      if (method === 0) return data;
      if (method === 8) return inflateRawSync(data);
      return null;
    }
    p += 46 + nameLen + extraLen + commentLen;
  }
  return null;
}

/**
 * What the agent's saved state says about its last turn (the events after the
 * last `user.message` in events.jsonl): the models its assistant messages came
 * from, which is the server's own record of the model and, unlike a self-ID,
 * names it (F121); and the tools it ran in its own container (`python_execution`
 * and the like), the AugLoop counterpart of a Chathub sandbox turn.
 */
export function lastAgentTurnOf(stateData: string): { models: string[]; tools: string[] } {
  const models = new Set<string>();
  const tools: string[] = [];
  try {
    const events = readZipEntry(Buffer.from(stateData, "base64"), "events.jsonl");
    if (!events) return { models: [], tools: [] };
    for (const line of events.toString("utf8").split("\n")) {
      if (!line.trim()) continue;
      const e = JSON.parse(line) as { type?: string; data?: { model?: unknown; toolName?: unknown } };
      if (e.type === "user.message") { models.clear(); tools.length = 0; }
      else if (e.type === "assistant.message" && typeof e.data?.model === "string") models.add(e.data.model);
      else if (e.type === "tool.execution_start" && typeof e.data?.toolName === "string") tools.push(e.data.toolName.replace(/^mcp__\w+?Tools__/, ""));
    }
  } catch {
    // not a state archive this reader knows
  }
  return { models: [...models], tools };
}

// --- Response status → the turn's result -------------------------------------------

export type AugLoopOutcome =
  | { kind: "done"; result: { value: string; errorCode?: string; message?: string } }
  | { kind: "disengaged"; result: { value: string; errorCode: string } }
  | { kind: "canceled" }
  | { kind: "fail"; message: string };

/**
 * What an `OutputAnnotation.responseStatus` means for the proxy, mapped onto the
 * `result` values the handler already acts on: `Throttled` → 429, `OutOfCredits`
 * → the priority-access 429 until the reset, `InternalError` → one re-send in a
 * fresh conversation; the policy refusals act as `Disengaged`. The statuses are
 * the ones Excel's client knows (hypotheses §48). Null for a turn still running.
 */
export function outcomeOfStatus(status: string): AugLoopOutcome | null {
  if (status === "partial" || status === "in_progress") return null;
  if (status === "complete") return { kind: "done", result: { value: "Success" } };
  if (status === "cancelled" || status === "canceled" || status.startsWith("error_query_canceled")) return { kind: "canceled" };
  if (/throttled$/.test(status) || status === "error_metering_enterprise_at_capacity") {
    return { kind: "done", result: { value: "Throttled", errorCode: status, message: `the Excel agent answered ${status}` } };
  }
  const limit = /^error_(?:claude_opus|model)_(daily|weekly)_limit$/.exec(status)?.[1] ?? (status === "error_model_limit" ? "daily" : null);
  if (limit) {
    const window = limit === "weekly" ? "week" : "day";
    return { kind: "done", result: { value: "OutOfCredits", errorCode: status, message: `The Excel agent's turn limit for this model is used up for the ${window} (${status}).` } };
  }
  if (status === "error_metering_consumer_credits_exhausted") {
    return { kind: "done", result: { value: "OutOfCredits", errorCode: status, message: `The Excel agent's credits are used up for the day (${status}).` } };
  }
  if (/^error_(rai_policy|dlp_)/.test(status)) return { kind: "disengaged", result: { value: "Disengaged", errorCode: status } };
  if (status === "error_user_not_licensed") return { kind: "fail", message: `M365 didn't serve the Excel agent to this account: it isn't licensed for it (${status}).` };
  if (status === "error_context_exhausted") return { kind: "fail", message: `The conversation no longer fits the Excel agent's context (${status}); start a new one.` };
  return { kind: "done", result: { value: "InternalError", errorCode: status } };
}

// --- Connection ------------------------------------------------------------------------

type Json = Record<string, any>;

interface Socket {
  ws: WebSocket;
  send(obj: Json | string): void;
}

/** One AugLoop session: the main endpoint names a slice, and the session runs on the slice, authenticated with TokenProvision. */
export class AugLoopConnection {
  readonly sessionId = randomUUID();
  private readonly baseCv = randomBytes(16).toString("base64").replace(/=+$/, "").slice(0, 22);
  private msgN = 0;
  private cvN = 0;
  private ctxN = 0;
  private respN = 0;
  private socket: Socket | null = null;
  private listeners = new Set<(m: Json) => void>();
  private keepalive: ReturnType<typeof setInterval> | null = null;
  private idleTimer: ReturnType<typeof setTimeout> | null = null;
  private closed = false;
  /** The account's permissions in the agent (`UserAllowedAnnotation`): `isAnthropicAvailable`, `isOutOfCredits`, … */
  allowed: Json | null = null;
  tokenExpiresAt = 0;

  private constructor(readonly flights: string, private readonly dumpId: string) {}

  get alive(): boolean {
    return !this.closed && this.socket?.ws.readyState === WebSocket.OPEN;
  }

  nextMessageId(): string { return `c${++this.msgN}`; }
  nextCv(): string { return `${this.baseCv}.${++this.cvN}`; }
  nextContextId(): string { return `C${++this.ctxN}`; }

  static async open(flights: string, token: string, signal?: AbortSignal): Promise<AugLoopConnection> {
    const c = new AugLoopConnection(flights, `augloop-${randomUUID()}`);
    try {
      await c.connect(token, signal);
      return c;
    } catch (err) {
      c.close();
      throw err;
    }
  }

  private clientMetadata() {
    return {
      appName: "Excel",
      appPlatform: "Web",
      appVersion: "16.0.20626.42304",
      uiLanguage: "en-US",
      releaseAudienceGroup: "Production",
      releaseChannel: null,
      releaseFork: null,
      sessionId: this.sessionId,
      flights: this.flights,
      userSystemTimezone: Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC",
      isClientTelemetrySampled: false,
      runtimeVersion: "3.0.667",
      docSessionId: this.sessionId,
    };
  }

  private openSocket(url: string, onMessage: (m: Json) => void, signal?: AbortSignal): Promise<Socket> {
    return new Promise((resolve, reject) => {
      const ws = new WebSocket(url, {
        headers: { Origin: "https://excel.officeapps.live.com", "User-Agent": USER_AGENT, "Cache-Control": "no-cache" },
        perMessageDeflate: false,
      });
      const timer = setTimeout(() => { ws.terminate(); reject(new Error(`AugLoop: no connection to ${url.slice(0, 60)} in ${CONNECT_TIMEOUT_MS / 1000}s`)); }, CONNECT_TIMEOUT_MS);
      const onAbort = () => { clearTimeout(timer); ws.terminate(); reject(new Error("aborted")); };
      signal?.addEventListener("abort", onAbort, { once: true });
      const socket: Socket = {
        ws,
        send: (obj) => {
          const raw = typeof obj === "string" ? obj : JSON.stringify(obj);
          dumpFrame(this.dumpId, "send", raw);
          ws.send(raw);
        },
      };
      ws.on("open", () => { clearTimeout(timer); signal?.removeEventListener("abort", onAbort); resolve(socket); });
      ws.on("unexpected-response", (_req, res) => { clearTimeout(timer); reject(new Error(`AugLoop: HTTP ${res.statusCode} on connect`)); });
      ws.on("error", (err) => { clearTimeout(timer); log.info(`AugLoop WebSocket error: ${err.message}`); reject(new Error(`AugLoop WebSocket error: ${err.message}`)); });
      ws.on("message", (data: WebSocket.RawData, isBinary: boolean) => {
        let raw = Buffer.isBuffer(data) ? data : Array.isArray(data) ? Buffer.concat(data) : Buffer.from(data);
        // Binary frames carry the same JSON after a 1-byte type (3) and a 4-byte length.
        if (isBinary) raw = raw[0] === 3 ? raw.subarray(5) : Buffer.alloc(0);
        const text = raw.toString("utf8");
        if (!text || text === "~") return;
        dumpFrame(this.dumpId, "recv", text);
        let m: Json;
        try { m = JSON.parse(text); } catch { return; }
        onMessage(m);
      });
    });
  }

  private waitFor(pred: (m: Json) => boolean, ms: number, what: string): Promise<Json> {
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => { this.listeners.delete(l); reject(new Error(`AugLoop: no ${what} in ${ms / 1000}s`)); }, ms);
      const l = (m: Json) => {
        if (m.H_?.T_ === "closed") { clearTimeout(timer); this.listeners.delete(l); reject(new Error(`AugLoop: the connection closed while waiting for ${what}`)); return; }
        if (pred(m)) { clearTimeout(timer); this.listeners.delete(l); resolve(m); }
      };
      this.listeners.add(l);
    });
  }

  private dispatch(m: Json): void {
    const T: string = m.H_?.T_ ?? "";
    // The server's own messages (`s…` ids) get a Response, as Excel's client sends.
    if (typeof m.messageId === "string" && m.messageId.startsWith("s") && !/Response$/.test(T)) {
      this.socket?.send({ H_: H(`${SESSION}Response`, []), messageId: m.messageId });
    }
    if (T === `${SESSION}ErrorResponse`) log.info(`AugLoop error for ${m.messageId}: ${m.code ?? ""} ${m.error ?? ""}`);
    for (const l of [...this.listeners]) l(m);
  }

  private async connect(token: string, signal?: AbortSignal): Promise<void> {
    // The main endpoint names the session's slice.
    let settle!: { resolve: (m: Json) => void; reject: (e: Error) => void };
    const initP = new Promise<Json>((resolve, reject) => { settle = { resolve, reject }; });
    const main = await this.openSocket(`wss://${host()}/`, (m) => {
      if (m.sliceUrl) settle.resolve(m);
      else if (m.error) settle.reject(new Error(`AugLoop: session init refused (${m.code ?? ""} ${m.error})`));
    }, signal);
    main.send("~");
    main.send({ protocolVersion: 2, clientMetadata: this.clientMetadata(), returnWorkflowInputTypes: false, enableRemoteExecutionNotification: false, createBlobStorageContainer: false, H_: H(`${SESSION}SessionInitMessage`), cv: this.baseCv, messageId: this.nextMessageId() });
    let timer: ReturnType<typeof setTimeout> | undefined;
    const init = await Promise.race([initP, new Promise<never>((_, rej) => { timer = setTimeout(() => rej(new Error("AugLoop: no SessionInitResponse")), CONNECT_TIMEOUT_MS); })])
      .finally(() => { clearTimeout(timer); main.ws.close(); });

    // The slice: the session proper.
    const s = await this.openSocket(init.sliceUrl, (m) => this.dispatch(m), signal);
    this.socket = s;
    s.ws.on("close", (code) => { if (!this.closed) log.info(`AugLoop connection closed (${code})`); this.close(); });
    s.send("~");
    s.send({ protocolVersion: 2, clientMetadata: this.clientMetadata(), sessionKey: init.sessionKey, origin: init.origin, returnWorkflowInputTypes: false, enableRemoteExecutionNotification: false, createBlobStorageContainer: false, H_: H(`${SESSION}SessionInitMessage`), cv: this.nextCv(), messageId: this.nextMessageId() });
    const provisioned = this.provisionToken(token);
    let n = 0;
    for (const type of ACTIVATED) {
      s.send({ annotationType: type, token: `${type}-${++n}`, ignoreExistingAnnotations: false, sendStateUpdates: false, returnAnnotationDoesNotExist: true, sendApologies: false, H_: H(`${SESSION}AnnotationActivationMessage`), cv: this.nextCv(), messageId: this.nextMessageId() });
    }
    const permission = this.waitFor((m) => /(UserAllowed|Forbidden)Annotation$/.test(m.annotationType ?? ""), CONNECT_TIMEOUT_MS, "permission check");
    s.send({ cv: this.nextCv(), ops: [{ parentPath: ["Signal", `${P}CheckPermissionSignal`], items: [{ id: randomUUID(), body: { H_: H(`${P}CheckPermissionSignal`, ["AugLoop_Signals_Signal"]) }, contextId: this.nextContextId() }], H_: H("AugLoop_Signals_SignalOperation", ["AugLoop_Core_Operation"]) }], H_: H(`${SESSION}SyncMessage`), messageId: this.nextMessageId() });
    await provisioned;
    const verdict = await permission;
    const body = verdict.ops?.[0]?.items?.[0]?.body ?? {};
    if (/Forbidden/.test(verdict.annotationType)) throw new Error(`AugLoop: the Excel agent is forbidden for this account (${trunc(JSON.stringify(body), 300)})`);
    this.allowed = body;
    log.info(`AugLoop session ready: isAnthropicAvailable=${body.isAnthropicAvailable}, isOutOfCredits=${body.isOutOfCredits}, token for ${Math.round((this.tokenExpiresAt - Date.now()) / 60_000)} min`);
    this.keepalive = setInterval(() => { try { this.socket?.send("~"); } catch {} }, KEEPALIVE_MS);
    this.keepalive.unref?.();
    this.touch();
  }

  /** Hand the session a (new) Entra ID token; resolves when the server took it. */
  async provisionToken(token: string): Promise<void> {
    const messageId = this.nextMessageId();
    const reply = this.waitFor((m) => m.messageId === messageId, CONNECT_TIMEOUT_MS, "TokenProvisionResponse");
    this.socket?.send({ authToken: token, version: 1, H_: H(`${SESSION}TokenProvisionMessage`), cv: this.nextCv(), messageId });
    const r = await reply;
    if (r.H_?.T_ !== `${SESSION}TokenProvisionResponse`) throw new Error(`AugLoop: token refused (${r.code ?? ""} ${r.error ?? r.H_?.T_})`);
    this.tokenExpiresAt = Date.now() + (Number(r.tokenExpirationSeconds) || 3600) * 1000;
  }

  /** Whether the provisioned token runs out within the margin. */
  get tokenExpiring(): boolean {
    return Date.now() > this.tokenExpiresAt - TOKEN_MARGIN_MS;
  }

  send(obj: Json): void {
    if (!this.socket) throw new Error("AugLoop: not connected");
    this.touch();
    this.socket.send(obj);
  }

  /** Answer an execution-correlated request (RunScript, SetClp, Load/Save state). */
  respond(wec: { callerMessageId: string; workflowExecutionId: string }, content: Json): void {
    this.send({ workflowExecutionCorrelation: { callerMessageId: wec.callerMessageId, workflowExecutionId: wec.workflowExecutionId }, finalResponse: true, content, H_: H(`${SESSION}ExecutionCorrelatedClientResponse`), messageId: `cst-1-${this.respN++}`, cv: this.nextCv() });
  }

  listen(l: (m: Json) => void): () => void {
    this.listeners.add(l);
    return () => { this.listeners.delete(l); };
  }

  /** Stop the session's running query, as Excel's Stop button does. */
  cancelQuery(): void {
    const cv = this.nextCv();
    this.send({
      annotationTypes: [`${P}CancelQueryResponseAnnotation`],
      transientItems: [{ parentPath: ["Signal", `${P}CancelQuerySignal`], id: randomUUID(), body: { H_: H(`${P}CancelQuerySignal`, ["AugLoop_Signals_Signal"]) } }],
      maxDelayMs: 10_000, sourceInfo: {}, caller: "Client", correlationInfo: { cvString: cv }, cv,
      H_: H(`${SESSION}Internal_GetAnnotationsRequestMessage`, [`${SESSION}StreamingRequest`, ...MESSAGE]),
      messageId: this.nextMessageId(),
    });
  }

  private touch(): void {
    if (this.idleTimer) clearTimeout(this.idleTimer);
    this.idleTimer = setTimeout(() => { log.info("AugLoop connection idle — closing"); this.close(); }, IDLE_CLOSE_MS);
    this.idleTimer.unref?.();
  }

  close(): void {
    if (this.closed) return;
    this.closed = true;
    if (this.keepalive) clearInterval(this.keepalive);
    if (this.idleTimer) clearTimeout(this.idleTimer);
    try { this.socket?.ws.close(); } catch {}
    for (const l of [...this.listeners]) l({ H_: { T_: "closed" } });
    this.listeners.clear();
  }
}

// --- One turn's stream -----------------------------------------------------------------

/** A CopilotStream that one AugLoop turn fills in. Exported for tests. */
export class AugLoopTurnStream {
  answer = "";
  reasoning = "";
  private queue: string[] = [];
  private done = false;
  private error: Error | null = null;
  private waiting: { resolve: (r: IteratorResult<string>) => void; reject: (e: Error) => void } | null = null;
  private reasoningListeners: Array<(d: string) => void> = [];
  result: { value: string; errorCode?: string; message?: string } | null = null;
  messageType: string | null = null;
  contentOrigin: string | null = "AugLoop";
  messageId: string | null = null;
  turnCount: number | null = null;

  /** Append streamed answer text; a new message starts a new paragraph, as in Excel's pane. */
  pushAnswer(text: string, newMessage = false): void {
    if (!text) return;
    const piece = newMessage && this.answer ? `\n\n${text}` : text;
    this.answer += piece;
    this.emit(piece);
  }

  /** The final text: adopted when longer, streamed only where it extends what streamed (prefix-safe, foldStreamText). */
  finalAnswer(text: string): void {
    const r = foldStreamText(this.answer, text);
    this.answer = r.answer;
    if (r.emit) this.emit(r.emit);
  }

  pushReasoning(text: string): void {
    if (!text) return;
    this.reasoning += text;
    for (const l of this.reasoningListeners) { try { l(text); } catch {} }
  }

  finalReasoning(text: string): void {
    const r = foldStreamText(this.reasoning, text);
    this.reasoning = r.answer;
    if (r.emit) for (const l of this.reasoningListeners) { try { l(r.emit); } catch {} }
  }

  private emit(text: string): void {
    if (this.waiting) { const w = this.waiting; this.waiting = null; w.resolve({ value: text, done: false }); }
    else this.queue.push(text);
  }

  finish(err?: Error): void {
    if (this.done) return;
    this.done = true;
    if (err) this.error = err;
    if (this.waiting) {
      const w = this.waiting;
      this.waiting = null;
      if (err) w.reject(err); else w.resolve({ value: undefined as any, done: true });
    }
  }

  get finished(): boolean { return this.done; }

  stream(): CopilotStream {
    const self = this;
    return {
      get fullText() { return self.answer; },
      get reasoning() { return self.reasoning; },
      onReasoning(l: (d: string) => void) { if (self.reasoning) l(self.reasoning); self.reasoningListeners.push(l); },
      images: [],
      get hasContent() { return self.answer.length > 0; },
      throttle: null,
      get contentOrigin() { return self.contentOrigin; },
      get messageType() { return self.messageType; },
      get messageId() { return self.messageId; },
      scores: null,
      get turnCount() { return self.turnCount; },
      turnState: null,
      get result() { return self.result; },
      metering: null,
      [Symbol.asyncIterator]() {
        return {
          next(): Promise<IteratorResult<string>> {
            if (self.queue.length) return Promise.resolve({ value: self.queue.shift()!, done: false });
            if (self.error) return Promise.reject(self.error);
            if (self.done) return Promise.resolve({ value: undefined as any, done: true });
            return new Promise((resolve, reject) => { self.waiting = { resolve, reject }; });
          },
        };
      },
    };
  }
}

/**
 * Fold one `OutputAnnotation` body into the turn's stream. Returns the outcome
 * once the turn has ended (a final `responseStatus`), else null. Exported for tests.
 */
export function applyOutput(out: AugLoopTurnStream, b: Json): AugLoopOutcome | null {
  const chunk = b.streamedChunk ?? (b.chunkType ? { type: b.chunkType, text: b.chunkContent } : null);
  if (chunk?.text) {
    if (chunk.type === "thoughts") out.pushReasoning(chunk.text);
    else out.pushAnswer(chunk.text, chunk.newMessage === true);
  }
  if (b.finalResponse) {
    if (typeof b.finalResponse.thoughts === "string" && b.finalResponse.thoughts) out.finalReasoning(b.finalResponse.thoughts);
    if (typeof b.finalResponse.assistantResponse === "string") out.finalAnswer(b.finalResponse.assistantResponse);
  }
  return typeof b.responseStatus === "string" ? outcomeOfStatus(b.responseStatus) : null;
}

// --- Conversation ------------------------------------------------------------------------

export interface AugLoopConversationOptions {
  /** An Entra ID token for `https://augloop.office.com/v2/.default` (getAugLoopToken). */
  getToken: () => Promise<string>;
  /** A Graph token, for the tenant's sensitivity labels. */
  getGraphToken: () => Promise<string | null>;
}

/**
 * The chat signal for one turn, in Chat only mode. Exported for tests. Where it
 * differs from Excel for the web's (hypotheses §48 F127): no container speedbump,
 * which would pause the agent for a confirmation the proxy can't give;
 * `isClaudeAvailable` is the account's (Excel sends it only in Auto mode); and
 * no `documentState`, which the server doesn't require.
 * The body's `H_` is required: without it the server refuses the message.
 */
export function buildChatSignal(opts: {
  modelType: string;
  query: string;
  queryId: string;
  signalId: string;
  conversationId: string;
  messages: { role: string; content: string }[];
  sensitivityLabel: SensitivityLabel | null;
  /** The account's `isAnthropicAvailable` (UserAllowedAnnotation). */
  isClaudeAvailable: boolean;
  hostSessionId: string;
  cv: string;
  contextId: string;
  messageId: string;
}): Json {
  const label = opts.sensitivityLabel
    ? { sensitivityLabel: { id: opts.sensitivityLabel.id, ...(opts.sensitivityLabel.priority !== undefined ? { priority: opts.sensitivityLabel.priority } : {}) } }
    : {};
  return {
    cv: opts.cv,
    ops: [{
      parentPath: ["Signal", `${P}Signal`],
      items: [{
        id: opts.signalId,
        contextId: opts.contextId,
        body: {
          clientConfig: {
            clientName: "Aurora",
            clientUILocale: "en-US",
            localizedStringMap: { RetryMessage: "I ran into a problem, let me try again.", TooManyIterations: "Too many iterations." },
            hostAriaSessionId: opts.hostSessionId.toUpperCase(),
          },
          queryId: opts.queryId,
          query: opts.query,
          allowResume: true,
          conversation: { conversationId: opts.conversationId, messages: opts.messages, grantedEntityPermissions: [] },
          documentConfig: { documentId: "", documentName: "Book.xlsx", documentType: "xlsx", documentUrl: "", isReadOnly: true, readOnlyReason: CHAT_ONLY_REASON, ...label },
          featureOverrides: {
            WebSearchEnabled: true, EnterpriseSearchEnabled: false, PowerBiMcpEnabled: false, PythonToolToggleEnabled: false,
            EnableExtractDataFromPageThumbnail: false, AgentInContainerEnabled: true, AgentStateStorageInContainerEnabled: true,
            IsWebSearchInAgentContainerEnabled: true, SpeedbumpInContainerEnabled: false, ExcelCopilotForReadOnlyFiles: true,
            CotLocaleInContainerEnabled: true, QuickAnswerEnabled: false,
          },
          agentModelType: opts.modelType,
          isAutoModeSelected: false,
          isAutoSelected: false,
          isClaudeAvailable: opts.isClaudeAvailable,
          enabledSkills: [],
          includeCustomSkills: false,
          connections: [],
          installedAddIns: [],
          registeredFormulas: [],
          changeHistoryEnabled: false,
          H_: H(`${P}Signal`, ["AugLoop_Core_TaskSignal", "AugLoop_Signals_Signal"]),
        },
      }],
      H_: H("AugLoop_Signals_SignalOperation", ["AugLoop_Core_Operation"]),
    }],
    H_: H(`${SESSION}SyncMessage`),
    messageId: opts.messageId,
  };
}

/**
 * One conversation with the Excel agent. Turns go one at a time over a
 * connection opened on the first and reused while it lives; the server's
 * conversation ID, the transcript (`messages`) and the agent's saved state
 * carry the conversation across connections.
 */
export class AugLoopConversation {
  private conversationId = "";
  private readonly messages: { role: "user" | "assistant"; content: string }[] = [];
  private readonly states = new Map<string, string>();
  private connection: AugLoopConnection | null = null;
  private chain: Promise<unknown> = Promise.resolve();
  private turns = 0;
  /** The models the last saved agent state names (lastAgentTurnOf). */
  private lastServed: string[] = [];

  constructor(private readonly opts: AugLoopConversationOptions) {}

  get turnCount(): number {
    return this.turns;
  }

  /** The server's ID for this conversation, once it has assigned one. */
  get serverConversationId(): string {
    return this.conversationId;
  }

  close(): void {
    this.connection?.close();
    this.connection = null;
  }

  /** Send one user message and stream the answer. Resolves once the message is sent. */
  chat(text: string, model: string, signal?: AbortSignal): Promise<CopilotStream> {
    const aug = augLoopModelOf(model);
    if (!aug) return Promise.reject(new Error(`${model} is not an AugLoop model`));
    // One turn at a time: the next waits until this one's stream has ended.
    let release!: () => void;
    const turnDone = new Promise<void>((r) => { release = r; });
    const previous = this.chain;
    this.chain = previous.then(() => turnDone);
    return previous
      .then(() => this.startTurn(text, model, aug.modelType, aug.flights, signal, release))
      .catch((err) => { release(); throw err; });
  }

  private async ensureConnection(flights: string, signal?: AbortSignal): Promise<AugLoopConnection> {
    let c = this.connection;
    if (c && (!c.alive || c.flights !== flights)) { c.close(); c = null; }
    if (!c) {
      c = await this.connect(flights, signal);
    } else if (c.tokenExpiring) {
      try {
        await c.provisionToken(await this.opts.getToken());
      } catch (err: any) {
        log.info(`Token re-provision failed (${err.message}) — reconnecting`);
        c.close();
        c = await this.connect(flights, signal);
      }
    }
    return c;
  }

  /** A new connection, with the conversation's agent state handled on it for as long as it lives. */
  private async connect(flights: string, signal?: AbortSignal): Promise<AugLoopConnection> {
    const c = await AugLoopConnection.open(flights, await this.opts.getToken(), signal);
    this.connection = c;
    // Load and Save carry the conversation's ID, not a turn's query ID, and a Save
    // can come after the turn's last output.
    c.listen((m) => {
      const kind = (m.annotationType ?? "").slice(P.length);
      if (kind !== "LoadAgentSessionStateAnnotation" && kind !== "SaveAgentSessionStateAnnotation") return;
      for (const item of (m.ops ?? []).flatMap((o: Json) => o.items ?? [])) {
        const b: Json = item.body ?? {};
        const wec = b.workflowExecutionCorrelation;
        if (!wec) continue;
        if (b.conversationId && !this.conversationId) this.conversationId = b.conversationId;
        if (kind === "LoadAgentSessionStateAnnotation") {
          const stateData = this.states.get(b.conversationId ?? "");
          log.info(`Agent state requested for ${b.conversationId || "a new conversation"}: ${stateData ? "sending the saved one" : "none saved"}`);
          c.respond(wec, { success: true, ...(stateData ? { stateData } : {}), H_: H(`${P}LoadAgentSessionStateResponseObject`, []) });
        } else {
          if (b.conversationId && typeof b.stateData === "string") {
            this.states.set(b.conversationId, b.stateData);
            const t = lastAgentTurnOf(b.stateData);
            this.lastServed = t.models;
            log.info(`Agent state saved (turn ${this.turns - 1}): served by ${t.models.join(", ") || "?"}${t.tools.length ? `; the agent ran its own tools: ${t.tools.join(", ")}` : ""}`);
          }
          c.respond(wec, { success: true, H_: H(`${P}SaveAgentSessionStateResponseObject`, []) });
        }
      }
    });
    return c;
  }

  private async startTurn(text: string, model: string, modelType: string, slotFlights: readonly string[], signal: AbortSignal | undefined, release: () => void): Promise<CopilotStream> {
    if (signal?.aborted) throw new Error("aborted");
    const flights = [...BASE_FLIGHTS, ...slotFlights].map((f) => `${f};`).join("");
    const label = await resolveSensitivityLabel(this.opts.getGraphToken);
    const c = await this.ensureConnection(flights, signal);
    if (isClaudeAugLoopModel(model) && c.allowed?.isAnthropicAvailable === false) {
      throw new Error(`M365 doesn't offer Anthropic models to this account in Excel's Copilot (isAnthropicAvailable=false), so ${model} isn't available`);
    }

    const turn = this.turns++;
    recordTurn(); // the throttle-avoidance bucket counts what the next line counts (turn-budget.ts)
    log.info(`Chat turn ${turn}: model=${model}, agentModelType=${modelType}, cid=${this.conversationId || "new"}, text=${JSON.stringify(trunc(text, 200))}`);
    this.messages.push({ role: "user", content: text });

    const out = new AugLoopTurnStream();
    out.turnCount = this.turns;
    const queryId = randomUUID();
    out.messageId = queryId;
    let lastFrame = Date.now();
    let stopped = false;
    let unlisten = () => {};
    this.lastServed = [];

    const finish = (err?: Error) => {
      if (out.finished) return;
      clearInterval(silence);
      unlisten();
      signal?.removeEventListener("abort", onAbort);
      if (out.answer) this.messages.push({ role: "assistant", content: out.answer });
      if (this.lastServed.length) out.contentOrigin = `AugLoop/${this.lastServed.join(",")}`;
      const status = out.result?.value ?? (err ? `error: ${err.message}` : stopped ? "stopped" : "no status");
      log.info(`Turn ${turn} ended: ${status}${out.result?.errorCode ? ` (${out.result.errorCode})` : ""}, ${out.answer.length} chars`);
      out.finish(err);
      release();
    };

    const silence = setInterval(() => {
      if (Date.now() - lastFrame > TURN_SILENCE_MS) finish(new Error(`AugLoop: no reply from the Excel agent in ${TURN_SILENCE_MS / 60_000} min`));
    }, 10_000);
    silence.unref?.();

    const onAbort = () => {
      if (stopped || out.finished) return;
      stopped = true;
      log.info("Aborted — cancelling the Excel agent's query");
      try { c.cancelQuery(); } catch {}
      setTimeout(() => finish(), 3_000).unref?.();
    };

    const onOutput = (b: Json) => {
      if (b.sydneyConversationId && !this.conversationId) this.conversationId = b.sydneyConversationId;
      const outcome = applyOutput(out, b);
      if (!outcome) return;
      if (outcome.kind === "canceled") { finish(); return; }
      if (outcome.kind === "fail") { finish(new Error(outcome.message)); return; }
      out.result = outcome.result;
      if (outcome.kind === "disengaged") out.messageType = "Disengaged";
      if (outcome.result.value !== "Success") log.info(`Turn result: ${outcome.result.value}${outcome.result.errorCode ? ` (${outcome.result.errorCode})` : ""}`);
      if (outcome.result.value === "Throttled") recordThrottle(Date.now(), outcome.result.errorCode);
      finish();
    };

    const signalMessageId = c.nextMessageId();
    unlisten = c.listen((m) => {
      if (m.H_?.T_ === "closed") { finish(out.answer ? undefined : new Error("AugLoop connection closed during the turn")); return; }
      // The server refused the chat signal itself (a malformed one: `1003 … type validation error`).
      if (m.H_?.T_ === `${SESSION}ErrorResponse` && m.messageId === signalMessageId) {
        finish(new Error(`AugLoop refused the chat signal: ${m.code ?? ""} ${m.error ?? ""}`.trim()));
        return;
      }
      const at: string = m.annotationType ?? "";
      if (!at.startsWith(P)) return;
      const kind = at.slice(P.length);
      // The conversation's listener (connect) answers these.
      if (kind === "LoadAgentSessionStateAnnotation" || kind === "SaveAgentSessionStateAnnotation") { lastFrame = Date.now(); return; }
      for (const item of (m.ops ?? []).flatMap((o: Json) => o.items ?? [])) {
        const b: Json = item.body ?? {};
        if (b.queryId && b.queryId !== queryId) continue;
        lastFrame = Date.now();
        const wec = b.workflowExecutionCorrelation;
        if (kind === "OutputAnnotation") {
          onOutput(b);
        } else if (kind === "RunScriptAnnotation" && wec) {
          c.respond(wec, { scriptResponse: { result: EMPTY_WORKBOOK, resultType: "returnResult", consoleOutput: "" }, H_: H(`${P}RunScriptResponseObject`, []), timeout: false });
        } else if (kind === "SetClpAnnotation" && wec) {
          c.respond(wec, { success: true, H_: H(`${P}SetClpResponseObject`, []) });
        } else if (wec) {
          // A request this client didn't activate: refuse it rather than leave the agent waiting.
          log.info(`Unhandled ${kind} — answering that it isn't supported`);
          c.respond(wec, { success: false, error: "Not supported by this client", H_: H(`${P}${kind.replace(/Annotation$/, "ResponseObject")}`, []) });
        }
      }
    });

    if (signal) signal.addEventListener("abort", onAbort, { once: true });

    c.send(buildChatSignal({
      modelType,
      query: text,
      queryId,
      signalId: randomUUID(),
      conversationId: this.conversationId,
      messages: this.messages.map((m) => ({ ...m })),
      sensitivityLabel: label,
      isClaudeAvailable: c.allowed?.isAnthropicAvailable === true,
      hostSessionId: c.sessionId,
      cv: c.nextCv(),
      contextId: c.nextContextId(),
      messageId: signalMessageId,
    }));
    return out.stream();
  }
}
