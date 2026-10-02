import { app, type HttpRequest, type HttpResponseInit, type InvocationContext } from "@azure/functions";
import { askFabricAgent } from "../fabricAgent.js";

type ChatBody = { question?: unknown };

async function chat(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  try {
    const authorization = request.headers.get("authorization");
    const token = authorization?.match(/^Bearer\s+(.+)$/i)?.[1];
    const body = (await request.json().catch(() => ({}))) as ChatBody;
    const question = typeof body.question === "string" ? body.question.trim() : "";

    if (!token) {
      return { status: 401, jsonBody: { error: "Sign in with Microsoft to ask a question." } };
    }
    if (!question) {
      return { status: 400, jsonBody: { error: "Enter a question to continue." } };
    }
    if (question.length > 4_000) {
      return { status: 400, jsonBody: { error: "Keep questions under 4,000 characters." } };
    }

    return { jsonBody: await askFabricAgent(question, token) };
  } catch (error) {
    context.error("Fabric Data Agent request failed", error);
    const message = error instanceof Error ? error.message : "An unexpected error occurred.";
    const unauthorized = /401|unauthorized/i.test(message);
    return {
      status: unauthorized ? 401 : 502,
      jsonBody: {
        error: unauthorized
          ? "Your Fabric session has expired or is not authorized. Sign in again and retry."
          : message,
      },
    };
  }
}

app.http("chat", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "chat",
  handler: chat,
});
