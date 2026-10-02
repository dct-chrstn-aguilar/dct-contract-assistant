import "dotenv/config";
import express, { type NextFunction, type Request, type Response } from "express";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { askFabricAgent } from "../api/src/fabricAgent.ts";
import { handleTeamsChat, handleTeamsChatHistory, handleTeamsChatStatus, teamsAuthorization } from "../api/src/teamsChat.ts";

const app = express();
const port = Number(process.env.PORT || 3001);
const currentDirectory = path.dirname(fileURLToPath(import.meta.url));
const distDirectory = path.resolve(currentDirectory, "../dist");

app.disable("x-powered-by");
app.use(express.json({ limit: "32kb" }));

app.post("/api/teams/chat", async (request, response) => {
  const result = await handleTeamsChat(teamsAuthorization({ get: (name) => request.header(name) }), request.body);
  response.setHeader("Cache-Control", "no-store");
  response.status(result.status).json(result.jsonBody);
});

app.get("/api/teams/chat/status/:taskHandle", async (request, response) => {
  const result = await handleTeamsChatStatus(
    teamsAuthorization({ get: (name) => request.header(name) }),
    request.params.taskHandle,
  );
  response.setHeader("Cache-Control", "no-store");
  response.status(result.status).json(result.jsonBody);
});

app.delete("/api/teams/chat/history", async (request, response) => {
  const result = await handleTeamsChatHistory(teamsAuthorization({ get: (name) => request.header(name) }), "DELETE");
  response.setHeader("Cache-Control", "no-store");
  response.status(result.status).send(result.jsonBody);
});

app.get("/api/health", (_request, response) => {
  response.json({
    ok: true,
    configured: Boolean(
      process.env.FABRIC_MCP_URL ||
        (process.env.FABRIC_WORKSPACE_ID && process.env.FABRIC_DATA_AGENT_ID),
    ),
    historyConfigured: Boolean(
      process.env.HISTORY_STORAGE_CONNECTION_STRING?.trim() &&
      process.env.HISTORY_TABLE_NAME?.trim(),
    ),
  });
});

app.post("/api/chat", async (request, response, next) => {
  try {
    const authorization = request.header("authorization");
    const browserToken = authorization?.match(/^Bearer\s+(.+)$/i)?.[1];
    const token =
      browserToken ||
      (process.env.NODE_ENV !== "production"
        ? process.env.FABRIC_ACCESS_TOKEN?.trim()
        : undefined);
    const question =
      typeof request.body?.question === "string" ? request.body.question.trim() : "";

    if (!token) {
      response.status(401).json({ error: "Sign in with Microsoft to ask a question." });
      return;
    }

    if (!question) {
      response.status(400).json({ error: "Enter a question to continue." });
      return;
    }

    if (question.length > 4_000) {
      response.status(400).json({ error: "Keep questions under 4,000 characters." });
      return;
    }

    response.json(await askFabricAgent(question, token));
  } catch (error) {
    next(error);
  }
});

if (process.env.NODE_ENV === "production") {
  app.use(express.static(distDirectory));
  app.get("/{*splat}", (_request, response) => {
    response.sendFile(path.join(distDirectory, "index.html"));
  });
}

app.use((error: unknown, _request: Request, response: Response, _next: NextFunction) => {
  void _next;
  const message = error instanceof Error ? error.message : "An unexpected error occurred.";
  const normalized = message.toLowerCase();
  const status = normalized.includes("401") || normalized.includes("unauthorized") ? 401 : 502;

  console.error("Fabric Data Agent request failed:", error);
  response.status(status).json({
    error:
      status === 401
        ? "Your Fabric session has expired or is not authorized. Sign in again and retry."
        : message,
  });
});

app.listen(port, () => {
  console.log(`DCT Contract Assistant API listening on http://localhost:${port}`);
});
