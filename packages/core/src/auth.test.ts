import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const ORIGIN = "https://device.login.microsoftonline.com";

let dir: string;

/** auth.ts reads M365_SECRETS_FILE once, at load. */
async function load(secrets: Record<string, unknown> | null) {
  const file = join(dir, "secrets.json");
  if (secrets) writeFileSync(file, JSON.stringify(secrets));
  else rmSync(file, { force: true });
  vi.stubEnv("M365_SECRETS_FILE", file);
  vi.stubEnv("M365_CACHE_FILE", join(dir, "msal-cache.json"));
  vi.stubEnv("M365_BROWSER_PROFILE", join(dir, "browser-profile"));
  vi.stubEnv("M365_DEBUG", "");
  vi.stubEnv("M365_TRACE", "");
  vi.resetModules();
  return import("./auth.js");
}

describe("loginClientCertificates", () => {
  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), "m365-auth-test-"));
  });
  afterEach(() => {
    vi.unstubAllEnvs();
    rmSync(dir, { recursive: true, force: true });
  });

  it("is undefined without clientCertPath, or without secrets.json", async () => {
    const creds = { email: "a@b.c", password: "p", mfaSecret: "JBSWY3DPEHPK3PXP" };
    expect((await load(creds)).loginClientCertificates()).toBeUndefined();
    expect((await load({ ...creds, clientCertPath: "" })).loginClientCertificates()).toBeUndefined();
    expect((await load(null)).loginClientCertificates()).toBeUndefined();
  });

  it("presents the PKCS#12 file to the device-auth origin, with its passphrase", async () => {
    const cert = join(dir, "abs.p12");
    writeFileSync(cert, "x");
    const m = await load({ clientCertPath: cert, clientCertPassword: "pw" });
    expect(m.loginClientCertificates()).toEqual([{ origin: ORIGIN, pfxPath: cert, passphrase: "pw" }]);
  });

  it("reads a relative clientCertPath from secrets.json's directory, and needs no password", async () => {
    writeFileSync(join(dir, "rel.p12"), "x");
    const m = await load({ clientCertPath: "rel.p12" });
    expect(m.loginClientCertificates()).toEqual([
      { origin: ORIGIN, pfxPath: join(dir, "rel.p12"), passphrase: undefined },
    ]);
  });

  it("throws when clientCertPath names a missing file", async () => {
    const m = await load({ clientCertPath: "missing.p12", clientCertPassword: "pw" });
    expect(() => m.loginClientCertificates()).toThrow(/clientCertPath .* does not exist: .*missing\.p12/);
  });
});
