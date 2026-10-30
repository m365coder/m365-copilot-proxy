import { chatDeletionsSettled } from "@m365-copilot/core";

/**
 * On shutdown (SIGTERM, SIGINT), wait for the Excel agent's temporary chats
 * that are due for deletion from Copilot history to be deleted, so the chats of
 * the last turns don't stay there. The wait is at most a few seconds.
 */
export default defineNitroPlugin((nitroApp) => {
  nitroApp.hooks.hook("close", () => chatDeletionsSettled());
});
