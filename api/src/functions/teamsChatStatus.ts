import { app, type HttpRequest, type HttpResponseInit } from "@azure/functions";
import { handleTeamsChatStatus, teamsAuthorization } from "../teamsChat.js";

async function teamsChatStatus(request: HttpRequest): Promise<HttpResponseInit> {
  const result = await handleTeamsChatStatus(
    teamsAuthorization(request.headers),
    request.params.taskHandle,
  );
  return {
    ...result,
    headers: {
      "Cache-Control": "no-store",
      "X-DCT-API-Version": "teams-service-principal-tasks-v1",
    },
  };
}

app.http("teamsChatStatus", {
  methods: ["GET"],
  authLevel: "anonymous", // The handler verifies the signed, scoped Teams user token.
  route: "teams/chat/status/{taskHandle}",
  handler: teamsChatStatus,
});
