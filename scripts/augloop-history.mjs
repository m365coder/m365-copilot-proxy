// The Excel agent's chats in Copilot history, listed as Excel for the web's history pane lists
// them (GetChats, docs/m365-copilot-api.md §13). With temporary chat on, the proxy deletes its
// own AugLoop chats after their turns (hypotheses §48 F128); this shows what is there.
//
// Usage: node scripts/augloop-history.mjs                list them: ID, creation time, name
//        node scripts/augloop-history.mjs --delete ID…   delete those chats (DeleteConversation)
// Build first (`pnpm build`).
import { randomUUID } from "node:crypto";
import { decodeJwt, deleteExcelChats, getToken } from "../packages/core/dist/index.mjs";

const token = await getToken();
const args = process.argv.slice(2);

if (args[0] === "--delete") {
  const ids = args.slice(1);
  if (ids.length === 0) { console.log("--delete takes the conversation IDs to delete"); process.exit(2); }
  await deleteExcelChats(token, ids);
  console.log(`Deleted ${ids.length} chat(s)`);
  process.exit(0);
}

const { oid, tid } = decodeJwt(token);
const request = { source: "excel", traceId: "", productThreadType: "Excel", mergeWorkWebChats: true, threadType: "appchat" };
const res = await fetch(`https://substrate.office.com/m365Copilot/GetChats?request=${encodeURIComponent(JSON.stringify(request))}`, {
  headers: {
    Authorization: `Bearer ${token}`,
    "X-ClientRequestId": randomUUID(),
    "X-Scenario": "ExcelSidepaneWorkCopilot",
    "X-AnchorMailbox": `Oid:${oid}@${tid}`,
    "X-RoutingParameter-SessionKey": oid,
  },
});
if (!res.ok) { console.log(`GetChats: HTTP ${res.status}`); process.exit(1); }
const chats = (await res.json()).chats ?? [];
console.log(`${chats.length} Excel agent chat(s) in Copilot history`);
for (const c of chats) console.log(`${c.conversationId}  ${new Date(c.createTimeUtc).toISOString().slice(0, 16)}Z  ${JSON.stringify(String(c.chatName).slice(0, 60))}`);
