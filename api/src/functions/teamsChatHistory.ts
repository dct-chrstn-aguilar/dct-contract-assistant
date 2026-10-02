import { app, type HttpRequest, type HttpResponseInit } from "@azure/functions";
import { handleTeamsChatHistory, teamsAuthorization } from "../teamsChat.js";

async function teamsChatHistory(request: HttpRequest): Promise<HttpResponseInit> {
  const result = await handleTeamsChatHistory(teamsAuthorization(request.headers), request.method === "DELETE" ? "DELETE" : "GET");
  return { ...result, headers: { "Cache-Control": "no-store", "X-DCT-API-Version": "teams-history-v1" } };
}

app.http("teamsChatHistory", { methods: ["GET", "DELETE"], authLevel: "anonymous", route: "teams/chat/history", handler: teamsChatHistory });
