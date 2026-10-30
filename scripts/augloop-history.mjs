// The Office agents' chats in Copilot history, listed as each app's history pane lists them
// (GetChats, docs/m365-copilot-api.md §13). With temporary chat on, the proxy deletes its own
// AugLoop chats after their turns (hypotheses §48 F128); this shows what is there.
//
// Usage: node scripts/augloop-history.mjs [--app excel|word]                list them: ID, creation time, name
//        node scripts/augloop-history.mjs --app excel|word --delete ID…    delete those chats (DeleteConversation)
// Without --app it lists both apps' chats. Build first (`pnpm build`).
import { deleteAgentChats, getToken, listAgentChats } from "../packages/core/dist/index.mjs";

const args = process.argv.slice(2);
const appArg = args.includes("--app") ? args[args.indexOf("--app") + 1] : undefined;
if (appArg !== undefined && appArg !== "excel" && appArg !== "word") { console.log("--app takes excel or word"); process.exit(2); }
const token = await getToken();

if (args.includes("--delete")) {
  const ids = args.slice(args.indexOf("--delete") + 1).filter((a) => !a.startsWith("--"));
  if (!appArg || ids.length === 0) { console.log("--delete takes --app and the conversation IDs to delete"); process.exit(2); }
  await deleteAgentChats(token, ids, appArg);
  console.log(`Deleted ${ids.length} ${appArg} chat(s)`);
  process.exit(0);
}

for (const app of appArg ? [appArg] : ["excel", "word"]) {
  const chats = await listAgentChats(token, app);
  console.log(`${chats.length} ${app} agent chat(s) in Copilot history`);
  for (const c of chats) console.log(`${c.conversationId}  ${new Date(c.createTimeUtc).toISOString().slice(0, 16)}Z  ${c.copilotReferenceId ?? ""}  ${JSON.stringify(String(c.chatName).slice(0, 60))}`);
}
