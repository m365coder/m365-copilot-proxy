// A long-lived headless browser signed in to Microsoft's own M365 Copilot web client as this account,
// controlled through a localhost HTTP endpoint, so that one sign-in serves a whole investigation and each
// step (screenshot, click, request) is one short command. Every response and every Chathub WebSocket frame
// is logged to OUT for scripts/webclient/analyze.mjs and diff.mjs.
//
// Run as the account, from its checkout (CHROMIUM_PATH defaults to the system Chromium):
//   OUT=<dir> PORT=<port> TOKEN=<random> node scripts/webclient/driver.mjs
// OUT defaults to ~/.cache/webclient-capture/<time>; without PORT/TOKEN the driver picks both and prints them.
// Then, from any user that can reach 127.0.0.1 (JS for eval on stdin; `page`, `context` and `h` in scope):
//   echo 'return h.bg("login", async () => { await page.goto("https://m365.cloud.microsoft/chat/"); return h.login(); })' |
//     scripts/webclient/drive.sh <port> <token>
//   echo 'return h.results.login' | scripts/webclient/drive.sh <port> <token>
//   echo 'return h.controls()' | scripts/webclient/drive.sh <port> <token>
//   echo 'await h.routeTone("Claude_Opus"); return h.bg("send", () => h.send("Which model are you?"))' | scripts/webclient/drive.sh …
//   scripts/webclient/drive.sh <port> <token> quit < /dev/null
// Helpers (h): shot(name), html(name), text(), controls(), login(), bg(name, fn) + results, routeTone(tone),
// send(text), log(line), redact(text).
//
// Sign-in uses the account's ~/.config/opencode-m365/secrets.json and a copy of its browser profile
// (~/.cache/webclient-profile, kept between runs, so later runs need no MFA). On "Verify your identity" it
// clicks only "Use a verification code" (never Text or Call, which send an SMS to or ring the account's
// phone) and enters the TOTP code (docs/hypotheses.md §47 F117). Like the proxy's login, it presents the
// device certificate that secrets.json names (clientCertPath, clientCertPassword) to
// device.login.microsoftonline.com, and takes its User-Agent from M365_LOGIN_UA.
//
// The endpoint runs any JS it is sent, as this account: it listens on 127.0.0.1 only, needs the token, and
// exits after IDLE_MS (default 45 min) without a command. Saved bodies skip the sign-in hosts and redact
// token-shaped strings, but they and the HTML still name the account: keep OUT private and delete it after.
import http from "node:http";
import { randomBytes } from "node:crypto";
import { appendFileSync, cpSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { join, resolve } from "node:path";
import { pathToFileURL } from "node:url";

const OUT = process.env.OUT ?? join(homedir(), ".cache", "webclient-capture", new Date().toISOString().replace(/[:.]/g, "-"));
const PORT = Number(process.env.PORT ?? 40000 + Math.floor(Math.random() * 10000));
const TOKEN = process.env.TOKEN ?? randomBytes(16).toString("hex");
const IDLE_MS = Number(process.env.IDLE_MS ?? 45 * 60 * 1000);

const repo = pathToFileURL(process.cwd()).href;
const pw = await import(`${repo}/node_modules/.pnpm/playwright@1.58.2/node_modules/playwright/index.js`);
const chromium = pw.chromium ?? pw.default?.chromium;
const { TOTP } = await import(`${repo}/node_modules/.pnpm/otpauth@9.5.0/node_modules/otpauth/dist/otpauth.esm.js`);

const cfg = join(homedir(), ".config", "opencode-m365");
const creds = JSON.parse(readFileSync(join(cfg, "secrets.json"), "utf-8"));
mkdirSync(join(OUT, "bodies"), { recursive: true });

const log = (s) => { appendFileSync(join(OUT, "driver.log"), `[${new Date().toISOString()}] ${s}\n`); return s; };
const JWT = /eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}/g;
const redact = (u) => String(u).replace(/(access_token|id_token|refresh_token|token|code|client_info|session_state)=[^&#"]+/gi, "$1=<redacted>").replace(JWT, "<jwt>");

const PROFILE = join(homedir(), ".cache", "webclient-profile");
if (!existsSync(PROFILE)) cpSync(join(cfg, "browser-profile"), PROFILE, { recursive: true });
for (const lock of ["SingletonLock", "SingletonCookie", "SingletonSocket"]) rmSync(join(PROFILE, lock), { force: true });

const context = await chromium.launchPersistentContext(PROFILE, {
  headless: true,
  executablePath: process.env.CHROMIUM_PATH ?? "/usr/sbin/chromium-browser",
  args: ["--no-sandbox", "--disable-dev-shm-usage", "--disable-blink-features=AutomationControlled"],
  userAgent: process.env.M365_LOGIN_UA ?? "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36",
  locale: "en-GB", timezoneId: "Europe/Copenhagen", viewport: { width: 1400, height: 900 },
  clientCertificates: creds.clientCertPath
    ? [{ origin: "https://device.login.microsoftonline.com", pfxPath: resolve(cfg, creds.clientCertPath), passphrase: creds.clientCertPassword }]
    : undefined,
});
await context.addInitScript(() => Object.defineProperty(navigator, "webdriver", { get: () => undefined }));
let page = context.pages()[0] ?? (await context.newPage());

// --- traffic log ----------------------------------------------------------------------------------------------
const STATIC = /\.(js|mjs|css|png|jpe?g|gif|svg|webp|woff2?|ttf|ico|map)(\?|$)/i;
const SIGN_IN = /login\.(microsoftonline|live)\.com|aadcdn|msauth/i;
let bodyN = 0;
context.on("response", async (res) => {
  const url = res.url();
  const req = res.request();
  const ct = res.headers()["content-type"] ?? "";
  const rec = { t: new Date().toISOString(), m: req.method(), url: redact(url).slice(0, 600), status: res.status(), ct: ct.slice(0, 50) };
  for (const k of ["server-flights", "x-calculatedbetarget", "x-bepartition"]) if (res.headers()[k]) rec[k] = res.headers()[k];
  if (!STATIC.test(url) && !SIGN_IN.test(url)) {
    if (req.method() !== "GET") { const pd = req.postData(); if (pd) rec.post = redact(pd).slice(0, 20000); }
    if (/json|text\/plain|xml/i.test(ct)) {
      try {
        const n = String(++bodyN).padStart(4, "0");
        writeFileSync(join(OUT, "bodies", `${n}.txt`), `${rec.m} ${rec.url}\n\n${redact(await res.text()).slice(0, 3_000_000)}`);
        rec.body = n;
      } catch {}
    }
  }
  appendFileSync(join(OUT, "responses.ndjson"), JSON.stringify(rec) + "\n");
});
function watchPage(p) {
  p.on("websocket", (ws) => {
    const id = Math.random().toString(36).slice(2, 8);
    appendFileSync(join(OUT, "websockets.txt"), `${new Date().toISOString()} ${id} ${redact(ws.url())}\n`);
    if (!/Chathub/i.test(ws.url())) return;
    const frame = (dir) => (f) => {
      const payload = typeof f === "string" ? f : f?.payload ?? "";
      appendFileSync(join(OUT, "ws-frames.ndjson"), JSON.stringify({ t: new Date().toISOString(), id, dir, p: redact(payload).slice(0, 200000) }) + "\n");
    };
    ws.on("framesent", frame("sent"));
    ws.on("framereceived", frame("recv"));
  });
}
watchPage(page);
context.on("page", (p) => { watchPage(p); log(`new page ${redact(p.url()).slice(0, 120)}`); });

// --- helpers --------------------------------------------------------------------------------------------------
const RS = "\x1e";
const h = {
  OUT, log, redact,
  setPage(p) { page = p; return "ok"; },
  async shot(name, fullPage = false) { const f = join(OUT, `${name}.png`); await page.screenshot({ path: f, fullPage }); return f; },
  async html(name) { const f = join(OUT, `${name}.html`); writeFileSync(f, redact(await page.content())); return f; },
  async text(max = 4000) { return (await page.evaluate(() => document.body.innerText).catch(() => "")).slice(0, max); },
  async controls() {
    return page.evaluate(() => [...document.querySelectorAll('button, [role="button"], [role="menuitem"], [role="menuitemradio"], [role="option"], [role="tab"], [role="combobox"], a[href]')]
      .filter((e) => { const r = e.getBoundingClientRect(); return r.width > 0 && r.height > 0; })
      .map((e, i) => ({ i, tag: e.tagName, role: e.getAttribute("role"), text: (e.innerText || "").trim().replace(/\s+/g, " ").slice(0, 60), aria: e.getAttribute("aria-label") || "", tid: e.getAttribute("data-testid") || "", popup: e.getAttribute("aria-haspopup") || "" })));
  },
  async login() { return login(); },
  // Long actions run here so that no control call waits on them; read h.results[name] later.
  results: {},
  bg(name, fn) {
    h.results[name] = { running: true, since: new Date().toISOString() };
    Promise.resolve().then(fn).then(
      (r) => { h.results[name] = { done: true, at: new Date().toISOString(), result: r }; log(`bg ${name} done`); },
      (e) => { h.results[name] = { done: true, at: new Date().toISOString(), error: String(e?.message ?? e).slice(0, 2000) }; log(`bg ${name} failed: ${e?.message}`); },
    );
    return `started ${name}`;
  },
  // Route the web client's Chathub sockets and rewrite the tone of each chat frame it sends (null: leave it).
  // Everything else in the request stays the client's own. Playwright installs the route when a page loads,
  // so this reloads the page the first time. The original and the sent frame go to ws-routed.ndjson.
  async routeTone(tone) {
    h.toneOverride = tone;
    if (h.routed) return `tone override ${tone}`;
    h.routed = true;
    const file = join(OUT, "ws-routed.ndjson");
    const rec = (dir, id, p) => appendFileSync(file, JSON.stringify({ t: new Date().toISOString(), id, dir, p: redact(p).slice(0, 200000) }) + "\n");
    await page.routeWebSocket(/Chathub/i, (ws) => {
      const id = Math.random().toString(36).slice(2, 8);
      rec("url", id, ws.url());
      const server = ws.connectToServer();
      ws.onMessage((msg) => {
        let out = msg;
        if (typeof msg === "string" && msg.includes('"target":"chat"')) {
          rec("sent-original", id, msg);
          if (h.toneOverride) {
            out = msg.split(RS).map((f) => {
              try {
                const j = JSON.parse(f);
                if (j.type === 4 && j.target === "chat" && j.arguments?.[0]) { log(`tone ${j.arguments[0].tone} -> ${h.toneOverride}`); j.arguments[0].tone = h.toneOverride; return JSON.stringify(j); }
              } catch {}
              return f;
            }).join(RS);
          }
        }
        rec("sent", id, typeof out === "string" ? out : "(binary)");
        server.send(out);
      });
      server.onMessage((msg) => { rec("recv", id, typeof msg === "string" ? msg : "(binary)"); ws.send(msg); });
    });
    await page.reload({ waitUntil: "domcontentloaded", timeout: 60000 });
    await page.waitForTimeout(9000);
    return `routed, tone override ${tone}`;
  },
  // Start a new chat, send `text` from the composer, wait, and return the page's last text.
  async send(text, waitMs = 25000) {
    await page.getByRole("link", { name: "New chat" }).first().click({ timeout: 5000 }).catch(() => {});
    await page.waitForTimeout(2500);
    await page.locator('[contenteditable="true"]:visible, textarea:visible, [role="textbox"]:visible').first().click({ timeout: 10000 });
    await page.keyboard.type(text, { delay: 5 });
    await page.waitForTimeout(500);
    await page.keyboard.press("Enter");
    await page.waitForTimeout(waitMs);
    return (await h.text(6000)).slice(-1500);
  },
};

async function chooseVerificationCode(tries) {
  await h.html("mfa-choice"); await h.shot("mfa-choice");
  const codeField = async () => (await page.locator('input[name="otc"]:visible, #idTxtBx_SAOTCC_OTC:visible').count()) > 0;
  const attempt = async (name, fn) => {
    let info = "";
    try { info = (await fn()) ?? ""; } catch (e) { tries.push(`${name}: ${String(e.message).split("\n")[0].slice(0, 200)}`); return false; }
    for (let i = 0; i < 12; i++) { if (await codeField()) { tries.push(`${name}: ok ${String(info).slice(0, 200)}`); return true; } await page.waitForTimeout(500); }
    tries.push(`${name}: clicked, no code field ${String(info).slice(0, 200)}`);
    return false;
  };
  // The accessible button is the one that works (F117); the others are kept as fallbacks, each checked
  // for the code field before the next is tried. All target the "Use a verification code" option only.
  const label = page.getByText("Use a verification code", { exact: true }).first();
  return (await attempt("role-button", () => page.getByRole("button", { name: /^use a verification code$/i }).first().click({ timeout: 4000 })))
    || (await attempt("label-click", () => label.click({ timeout: 4000 })))
    || (await attempt("dom-click", () => page.evaluate(() => {
      const leaf = [...document.querySelectorAll("div, span, a, button, p, td")].find((e) => e.children.length === 0 && e.textContent.trim() === "Use a verification code");
      if (!leaf) throw new Error("label not found");
      const target = leaf.closest('[role="button"], [data-value], a, button, [tabindex]') || leaf;
      target.click();
      return target.outerHTML.slice(0, 300);
    })))
    || (await attempt("keyboard", async () => { await label.focus(); await page.keyboard.press("Enter"); }));
}

// Whichever sign-in page is showing, act on it; stop when the page has left the sign-in host.
async function login() {
  const steps = [];
  const tries = [];
  let otpSent = false;
  const seen = async (sel) => (await page.locator(sel).count().catch(() => 0)) > 0;
  const submit = () => page.locator('input[type="submit"]:visible, button[type="submit"]:visible').first().click({ timeout: 5000 });
  const deadline = Date.now() + 120000;
  while (Date.now() < deadline) {
    if (!/login\.microsoftonline\.com|login\.live\.com|\/oauth2\//i.test(page.url())) {
      steps.push(`signed in: ${redact(page.url()).slice(0, 100)}`);
      log(`login: ${steps.join(" | ")} || ${tries.join(" | ")}`);
      return { ok: true, steps, tries };
    }
    let did = "nothing";
    try {
      if (await seen('input[name="passwd"]:visible')) {
        await page.locator('input[name="passwd"]:visible').first().fill(creds.password); await submit(); did = "password";
      } else if (!otpSent && (await seen('input[name="otc"]:visible, #idTxtBx_SAOTCC_OTC:visible'))) {
        const sel = (await seen("#idTxtBx_SAOTCC_OTC:visible")) ? "#idTxtBx_SAOTCC_OTC" : 'input[name="otc"]';
        await page.locator(`${sel}:visible`).first().fill(new TOTP({ secret: creds.mfaSecret }).generate());
        if (sel === "#idTxtBx_SAOTCC_OTC" && (await seen("#idSubmit_SAOTCC_Continue:visible"))) await page.locator("#idSubmit_SAOTCC_Continue:visible").first().click({ timeout: 5000 });
        else await submit();
        otpSent = true; did = `code (${sel})`;
      } else if (await seen('input[name="loginfmt"]:visible')) {
        await page.locator('input[name="loginfmt"]:visible').first().fill(creds.email); await submit(); did = "email";
      } else if (await page.getByText("Use a verification code", { exact: true }).first().isVisible().catch(() => false)) {
        did = (await chooseVerificationCode(tries)) ? "chose verification code" : "verification-code choice FAILED";
      } else if (await seen("#idSIButton9:visible")) {
        await page.locator("#idSIButton9:visible").first().click({ timeout: 5000 }); did = "yes/next (#idSIButton9)";
      } else if (await seen("#tilesHolder:visible")) {
        const tile = page.locator(`#tilesHolder [data-test-id="${creds.email.toLowerCase()}"]:visible`).first();
        if (await tile.isVisible().catch(() => false)) await tile.click({ timeout: 5000 });
        else await page.locator("#tilesHolder").getByText(creds.email, { exact: false }).first().click({ timeout: 5000 });
        did = "account tile";
      }
    } catch (e) { did += ` ERR ${String(e.message).split("\n")[0].slice(0, 160)}`; }
    steps.push(did);
    if (did.includes("FAILED")) break;
    if (did === "nothing" && steps.filter((s) => s === "nothing").length % 4 === 1) { await h.shot(`login-${steps.length}`); await h.html(`login-${steps.length}`); }
    await page.waitForTimeout(2500);
  }
  log(`login FAILED: ${steps.join(" | ")} || ${tries.join(" | ")}`);
  return { ok: false, steps, tries, url: redact(page.url()).slice(0, 150), text: await h.text(1500) };
}

// --- control endpoint -----------------------------------------------------------------------------------------
const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
let lastUse = Date.now();
const server = http.createServer(async (req, res) => {
  if (!req.url.startsWith(`/${TOKEN}/`)) { res.writeHead(404); res.end(); return; }
  lastUse = Date.now();
  const cmd = req.url.slice(TOKEN.length + 2);
  let body = "";
  for await (const c of req) body += c;
  let out;
  let quit = false;
  try {
    if (cmd === "eval") out = { ok: true, result: await new AsyncFunction("page", "context", "h", body)(page, context, h) };
    else if (cmd === "quit") { out = { ok: true, result: "closing" }; quit = true; }
    else out = { ok: false, error: `unknown command ${cmd}` };
  } catch (e) { out = { ok: false, error: String(e?.stack ?? e).slice(0, 4000) }; }
  lastUse = Date.now();
  res.writeHead(200, { "content-type": "application/json" });
  res.end(JSON.stringify(out, null, 1));
  if (quit) await shutdown("quit command");
});
async function shutdown(why) {
  log(`shutdown: ${why}`);
  server.close();
  await Promise.race([context.close(), new Promise((r) => setTimeout(r, 10000))]);
  process.exit(0);
}
setInterval(() => { if (Date.now() - lastUse > IDLE_MS) shutdown("idle"); }, 60000);
server.listen(PORT, "127.0.0.1", () => {
  log(`driver ready on 127.0.0.1:${PORT}`);
  if (!process.env.TOKEN) console.log(`driver ready: port ${PORT} token ${TOKEN} out ${OUT}`);
});
