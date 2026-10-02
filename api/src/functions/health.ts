import { app } from "@azure/functions";

app.http("health", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "health",
  handler: async () => ({
    jsonBody: {
      ok: true,
      authentication: "service-principal",
      configured: Boolean(
        process.env.FABRIC_MCP_URL ||
          (process.env.FABRIC_WORKSPACE_ID && process.env.FABRIC_DATA_AGENT_ID),
      ) && Boolean(
        (process.env.FABRIC_SERVICE_PRINCIPAL_CLIENT_ID || process.env.TEAMS_CLIENT_ID) &&
        (process.env.FABRIC_SERVICE_PRINCIPAL_TENANT_ID || process.env.TEAMS_TENANT_ID) &&
        (process.env.FABRIC_SERVICE_PRINCIPAL_CLIENT_SECRET || process.env.TEAMS_CLIENT_SECRET)
      ),
      historyConfigured: Boolean(
        process.env.HISTORY_STORAGE_CONNECTION_STRING?.trim() &&
        process.env.HISTORY_TABLE_NAME?.trim(),
      ),
    },
  }),
});
