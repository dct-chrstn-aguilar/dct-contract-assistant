import { app, type HttpRequest, type HttpResponseInit } from "@azure/functions";
import { handleTeamsChat, teamsAuthorization } from "../teamsChat.js";

async function teamsChat(request: HttpRequest): Promise<HttpResponseInit> {
  const result = await handleTeamsChat(
    teamsAuthorization(request.headers),
    await request.json().catch(() => null),
  );
  return {
    ...result,
    headers: {
      "Cache-Control": "no-store",
      "X-DCT-API-Version": "teams-service-principal-tasks-v1",
    },
  };
}

app.http("teamsChat", {
  methods: ["POST"],
  authLevel: "anonymous", // The handler verifies the signed, scoped Teams user token.
  route: "teams/chat",
  handler: teamsChat,
});
