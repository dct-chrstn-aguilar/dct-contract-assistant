import { app, type Timer } from "@azure/functions";
import { cleanupConversationHistory } from "../conversationHistory.js";

// Azure Functions timers use UTC. 16:00 UTC is 00:00 the following day in Singapore.
export async function cleanupConversationHistoryTimer(_timer: Timer): Promise<void> {
  const result = await cleanupConversationHistory();
  console.info("Conversation history cleanup completed", result);
}

app.timer("cleanupConversationHistory", {
  schedule: "0 0 16 * * *",
  runOnStartup: false,
  handler: cleanupConversationHistoryTimer,
});
