import {
  Client,
  StreamableHTTPClientTransport,
  type JSONValue,
} from "@modelcontextprotocol/client";
import {
  CreateTaskResultV2Schema,
  DetailedTaskV2Schema,
} from "@modelcontextprotocol/ext-tasks/core/v2";

type TextContent = { type: "text"; text: string };

type FabricTaskHandle = {
  taskId: string;
  protocolVersion: string;
  sessionId?: string;
};

type JsonRecord = Record<string, JSONValue>;

export type FabricAgentTaskResult =
  | { status: "working"; taskHandle: string; retryAfterMs: number }
  | { status: "completed"; answer: string; toolName: string };

export type FabricAgentTaskStatus =
  | { status: "working"; retryAfterMs: number }
  | { status: "completed"; answer: string };

const clientInfo = { name: "dct-contract-assistant", version: "0.1.0" };
const clientCapabilities = {
  extensions: { "io.modelcontextprotocol/tasks": {} },
};
const defaultPollIntervalMs = 2_000;
const maximumPollIntervalMs = 3_000;
const toolCacheLifetimeMs = 5 * 60_000;
type FabricTool = { toolName: string; questionProperty: string };
let cachedTool: { value: FabricTool; expiresAt: number } | undefined;

function getEndpoint(): string {
  const configuredUrl = process.env.FABRIC_MCP_URL?.trim();

  if (configuredUrl) {
    let endpoint: URL;

    try {
      endpoint = new URL(configuredUrl);
    } catch {
      throw new Error("FABRIC_MCP_URL must be a valid URL.");
    }

    const isFabricHost =
      endpoint.hostname === "api.fabric.microsoft.com" ||
      endpoint.hostname.endsWith(".api.fabric.microsoft.com");

    if (endpoint.protocol !== "https:" || !isFabricHost) {
      throw new Error("FABRIC_MCP_URL must use HTTPS and a Microsoft Fabric API host.");
    }

    return endpoint.toString();
  }

  const workspaceId = process.env.FABRIC_WORKSPACE_ID?.trim();
  const dataAgentId = process.env.FABRIC_DATA_AGENT_ID?.trim();

  if (!workspaceId || !dataAgentId) {
    throw new Error(
      "The Fabric connection is not configured. Set FABRIC_MCP_URL, or set FABRIC_WORKSPACE_ID and FABRIC_DATA_AGENT_ID.",
    );
  }

  return `https://api.fabric.microsoft.com/v1/mcp/workspaces/${encodeURIComponent(workspaceId)}/dataagents/${encodeURIComponent(dataAgentId)}/agent`;
}

function readAnswer(content: unknown): string {
  if (!Array.isArray(content)) return "";

  return content
    .filter(
      (block): block is TextContent =>
        typeof block === "object" &&
        block !== null &&
        "type" in block &&
        block.type === "text" &&
        "text" in block &&
        typeof block.text === "string",
    )
    .map((block) => block.text)
    .join("\n")
    .trim();
}

function readResult(result: unknown): string {
  const content =
    typeof result === "object" && result !== null && "content" in result
      ? result.content
      : undefined;
  const answer = readAnswer(content);

  if (!answer) {
    throw new Error("The Fabric Data Agent returned an empty response.");
  }

  return answer;
}

function encodeTaskHandle(handle: FabricTaskHandle): string {
  return Buffer.from(JSON.stringify(handle), "utf8").toString("base64url");
}

function decodeTaskHandle(value: string): FabricTaskHandle {
  if (!value || value.length > 2_048 || !/^[A-Za-z0-9_-]+$/.test(value)) {
    throw new Error("The Fabric task handle is invalid.");
  }

  try {
    const parsed = JSON.parse(Buffer.from(value, "base64url").toString("utf8")) as Partial<FabricTaskHandle>;
    if (typeof parsed.taskId !== "string" || !parsed.taskId) throw new Error("Missing task ID");
    if (typeof parsed.protocolVersion !== "string" || !parsed.protocolVersion) {
      throw new Error("Missing protocol version");
    }
    if (parsed.sessionId !== undefined && typeof parsed.sessionId !== "string") {
      throw new Error("Invalid session ID");
    }
    return {
      taskId: parsed.taskId,
      protocolVersion: parsed.protocolVersion,
      sessionId: parsed.sessionId,
    };
  } catch {
    throw new Error("The Fabric task handle is invalid.");
  }
}

function parseRpcMessages(text: string, contentType: string): unknown[] {
  if (contentType.includes("text/event-stream")) {
    return text
      .split(/\r?\n\r?\n/)
      .flatMap((event) => {
        const data = event
          .split(/\r?\n/)
          .filter((line) => line.startsWith("data:"))
          .map((line) => line.slice(5).trimStart())
          .join("\n");
        if (!data) return [];
        try {
          return [JSON.parse(data)];
        } catch {
          return [];
        }
      });
  }

  const parsed = JSON.parse(text) as unknown;
  return Array.isArray(parsed) ? parsed : [parsed];
}

function taskMetadata(protocolVersion: string): JsonRecord {
  return {
    "io.modelcontextprotocol/protocolVersion": protocolVersion,
    "io.modelcontextprotocol/clientInfo": clientInfo,
    "io.modelcontextprotocol/clientCapabilities": clientCapabilities,
  };
}

async function dispatchModernRequest(
  method: string,
  params: JsonRecord,
  accessToken: string,
  protocolVersion: string,
  sessionId?: string,
  timeoutMs = 15_000,
): Promise<unknown> {
  const requestId = `dct-${crypto.randomUUID()}`;
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  const headers: Record<string, string> = {
    Accept: "application/json, text/event-stream",
    Authorization: `Bearer ${accessToken}`,
    "Content-Type": "application/json",
    "Mcp-Method": method,
    "Mcp-Protocol-Version": protocolVersion,
  };

  if (sessionId) headers["Mcp-Session-Id"] = sessionId;
  if (method === "tools/call" && typeof params.name === "string") {
    headers["Mcp-Name"] = params.name;
  }

  try {
    const response = await fetch(getEndpoint(), {
      method: "POST",
      headers,
      signal: controller.signal,
      body: JSON.stringify({
        jsonrpc: "2.0",
        id: requestId,
        method,
        params: { ...params, _meta: taskMetadata(protocolVersion) },
      }),
    });
    const text = await response.text();

    if (!response.ok) {
      throw new Error(`Fabric MCP request failed with HTTP ${response.status}.`);
    }

    const message = parseRpcMessages(text, response.headers.get("content-type") ?? "")
      .find((candidate) =>
        typeof candidate === "object" &&
        candidate !== null &&
        "id" in candidate &&
        candidate.id === requestId,
      );

    if (!message || typeof message !== "object") {
      throw new Error("Fabric MCP returned no matching response.");
    }
    if ("error" in message) {
      const error = message.error;
      const errorMessage =
        typeof error === "object" && error !== null && "message" in error && typeof error.message === "string"
          ? error.message
          : "Fabric MCP request failed.";
      throw new Error(errorMessage);
    }
    if (!("result" in message)) {
      throw new Error("Fabric MCP returned an invalid response.");
    }

    return message.result;
  } catch (error) {
    if (error instanceof Error && error.name === "AbortError") {
      throw new Error("Fabric MCP did not acknowledge the request in time.");
    }
    throw error;
  } finally {
    clearTimeout(timeout);
  }
}

function readTool(result: unknown): FabricTool {
  if (!result || typeof result !== "object" || !("tools" in result) || !Array.isArray(result.tools)) {
    throw new Error("The published Fabric Data Agent did not return a tool list.");
  }

  const tool = result.tools.at(0);
  if (!tool || typeof tool !== "object" || !("name" in tool) || typeof tool.name !== "string") {
    throw new Error("The published Fabric Data Agent did not expose an MCP tool.");
  }

  const inputSchema = "inputSchema" in tool && typeof tool.inputSchema === "object" && tool.inputSchema !== null
    ? tool.inputSchema as { required?: unknown; properties?: unknown }
    : {};
  const required = Array.isArray(inputSchema.required)
    ? inputSchema.required.find((value): value is string => typeof value === "string")
    : undefined;
  const properties = inputSchema.properties && typeof inputSchema.properties === "object"
    ? Object.keys(inputSchema.properties)
    : [];
  const questionProperty =
    required ??
    properties.find((key) => /question|query|prompt|input/i.test(key)) ??
    properties.at(0);

  if (!questionProperty) {
    throw new Error("The Fabric Data Agent tool has no question input.");
  }

  return { toolName: tool.name, questionProperty };
}

function retryAfter(value: number | undefined): number {
  if (!Number.isFinite(value)) return defaultPollIntervalMs;
  return Math.min(maximumPollIntervalMs, Math.max(1_000, Math.round(value!)));
}

export async function startFabricAgentTask(
  question: string,
  accessToken: string,
): Promise<FabricAgentTaskResult> {
  const startedAt = performance.now();
  const endpoint = getEndpoint();
  const client = new Client(clientInfo, {
    capabilities: clientCapabilities,
    versionNegotiation: { mode: "auto", probe: { timeoutMs: 10_000 } },
  });
  const transport = new StreamableHTTPClientTransport(new URL(endpoint), {
    requestInit: { headers: { Authorization: `Bearer ${accessToken}` } },
  });

  try {
    await client.connect(transport, { timeout: 12_000 });
    const connectedAt = performance.now();
    const protocolVersion = transport.protocolVersion;
    const taskExtension = client.getServerCapabilities()?.extensions?.["io.modelcontextprotocol/tasks"];

    if (!protocolVersion || !taskExtension || typeof taskExtension !== "object") {
      throw new Error("The Fabric MCP endpoint does not advertise background task support.");
    }

    return await startTaskWithSession(question, accessToken, protocolVersion, transport.sessionId, {
      startedAt,
      connectMs: connectedAt - startedAt,
    });
  } finally {
    await client.close().catch(() => undefined);
  }
}

async function startTaskWithSession(
  question: string,
  accessToken: string,
  protocolVersion: string,
  sessionId: string | undefined,
  timing: { startedAt: number; connectMs: number },
): Promise<FabricAgentTaskResult> {
  const cached = cachedTool && cachedTool.expiresAt > Date.now() ? cachedTool.value : undefined;
  const discoveryStartedAt = performance.now();
  const tool = cached ?? readTool(await dispatchModernRequest(
    "tools/list",
    {},
    accessToken,
    protocolVersion,
    sessionId,
    10_000,
  ));
  if (!cached) cachedTool = { value: tool, expiresAt: Date.now() + toolCacheLifetimeMs };
  const discoveredAt = performance.now();
  const result = await dispatchModernRequest(
    "tools/call",
    { name: tool.toolName, arguments: { [tool.questionProperty]: question } },
    accessToken,
    protocolVersion,
    sessionId,
    15_000,
  );
  console.info("Fabric task start timing", {
    connectMs: Math.round(timing.connectMs),
    discoveryMs: Math.round(discoveredAt - discoveryStartedAt),
    taskStartMs: Math.round(performance.now() - discoveredAt),
    totalMs: Math.round(performance.now() - timing.startedAt),
    toolCacheHit: Boolean(cached),
  });
  const task = CreateTaskResultV2Schema.safeParse(result);
  if (task.success) {
    return {
      status: "working",
      taskHandle: encodeTaskHandle({ taskId: task.data.taskId, protocolVersion, sessionId }),
      retryAfterMs: retryAfter(task.data.pollIntervalMs),
    };
  }
  return { status: "completed", answer: readResult(result), toolName: tool.toolName };
}

export async function pollFabricAgentTask(
  taskHandle: string,
  accessToken: string,
): Promise<FabricAgentTaskStatus> {
  const startedAt = performance.now();
  const handle = decodeTaskHandle(taskHandle);
  const result = await dispatchModernRequest(
    "tasks/get",
    { taskId: handle.taskId },
    accessToken,
    handle.protocolVersion,
    handle.sessionId,
    15_000,
  );
  const task = DetailedTaskV2Schema.safeParse(result);

  if (!task.success) {
    throw new Error("Fabric returned an invalid task status.");
  }

  if (task.data.status === "completed") {
    console.info("Fabric task poll timing", { status: "completed", pollMs: Math.round(performance.now() - startedAt) });
    return { status: "completed", answer: readResult(task.data.result) };
  }
  if (task.data.status === "failed") {
    throw new Error(task.data.error.message || task.data.statusMessage || "The Fabric task failed.");
  }
  if (task.data.status === "cancelled") {
    throw new Error(task.data.statusMessage || "The Fabric task was cancelled.");
  }
  if (task.data.status === "input_required") {
    throw new Error("The Fabric task requires interactive input, which this assistant does not support.");
  }

  console.info("Fabric task poll timing", { status: "working", pollMs: Math.round(performance.now() - startedAt) });
  return {
    status: "working",
    retryAfterMs: retryAfter(task.data.pollIntervalMs),
  };
}

export async function askFabricAgent(question: string, accessToken: string) {
  const result = await startFabricAgentTask(question, accessToken);

  if (result.status === "completed") return result;

  let retryAfterMs = result.retryAfterMs;
  while (true) {
    await new Promise((resolve) => setTimeout(resolve, retryAfterMs));
    const status = await pollFabricAgentTask(result.taskHandle, accessToken);
    if (status.status === "completed") {
      return { answer: status.answer, toolName: "fabric-data-agent" };
    }
    retryAfterMs = status.retryAfterMs;
  }
}
