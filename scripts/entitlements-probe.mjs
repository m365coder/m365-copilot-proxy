// What Microsoft has assigned to this account: licences (SKUs and service plans), assigned plans and group
// memberships, from Graph. Silent token only (the cached refresh token), so it never starts a browser login.
// Diff two accounts' output to see what separates them (§47 F114: two non-premium accounts with no licence
// at all). The output names the account's groups: keep it out of the repo.
// Usage: node scripts/entitlements-probe.mjs > entitlements.json
import { readFileSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

const ROOT = pathToFileURL(process.cwd()).href;
const msal = await import(`${ROOT}/packages/core/node_modules/@azure/msal-node/dist/index.mjs`);
const { PublicClientApplication } = msal.default ?? msal;

const CACHE = process.env.M365_CACHE_FILE ?? join(homedir(), ".config", "opencode-m365", "msal-cache.json");
const app = new PublicClientApplication({ auth: { clientId: "c0ab8ce9-e9a0-42e7-b064-33d422df41f1", authority: "https://login.microsoftonline.com/common" } });
app.getTokenCache().deserialize(readFileSync(CACHE, "utf-8"));
const [account] = await app.getTokenCache().getAllAccounts();
let token;
try {
  token = (await app.acquireTokenSilent({ scopes: ["https://graph.microsoft.com/.default"], account })).accessToken;
  writeFileSync(CACHE, app.getTokenCache().serialize());
} catch (e) {
  console.log(JSON.stringify({ error: `no silent Graph token: ${String(e.errorCode ?? e.message).slice(0, 200)}` }));
  process.exit(1);
}

const g = async (u) => {
  const r = await fetch(`https://graph.microsoft.com/v1.0${u}`, { headers: { Authorization: `Bearer ${token}` } });
  let body; try { body = await r.json(); } catch { body = {}; }
  return r.status >= 400 ? { error: r.status, code: body.error?.code } : body;
};
const ld = await g("/me/licenseDetails");
const me = await g("/me?$select=usageLocation,preferredLanguage,accountEnabled,userType,assignedPlans");
const groups = await g("/me/transitiveMemberOf?$select=displayName,id&$top=999");
console.log(JSON.stringify({
  licences: ld.error ? ld : (ld.value ?? []).map((s) => ({ sku: s.skuPartNumber, plans: (s.servicePlans ?? []).map((p) => `${p.servicePlanName}=${p.provisioningStatus}`).sort() })),
  me: me.error ? me : { usageLocation: me.usageLocation, preferredLanguage: me.preferredLanguage, accountEnabled: me.accountEnabled, userType: me.userType, assignedPlans: (me.assignedPlans ?? []).map((p) => `${p.service}/${p.servicePlanId}=${p.capabilityStatus}`).sort() },
  groups: groups.error ? groups : (groups.value ?? []).map((x) => `${x["@odata.type"]?.replace("#microsoft.graph.", "")}|${x.displayName ?? ""}`).sort(),
}, null, 1));
process.exit(0);
