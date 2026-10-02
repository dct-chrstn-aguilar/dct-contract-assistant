import {
  pollFabricAgentTask,
  startFabricAgentTask,
  type FabricAgentTaskResult,
  type FabricAgentTaskStatus,
} from "./fabricAgent.js";
import { appendConversation, clearConversation, ConversationHistoryError, forgetTask, questionWithConversation, readConversation, rememberTask, taskForUser } from "./conversationHistory.js";
import { getFabricServicePrincipalTokenAndTeamsUser, getFabricServicePrincipalTokenForTeams, TeamsAuthError, type TeamsUserIdentity } from "./teamsAuth.js";

type TeamsAccess = { accessToken: string; user: TeamsUserIdentity };
type StartDependencies = ({ getAccess: (assertion: string) => Promise<TeamsAccess> } | { getToken: typeof getFabricServicePrincipalTokenForTeams }) & {
  start: (question: string, accessToken: string) => Promise<FabricAgentTaskResult>;
};

type PollDependencies = ({ getAccess: (assertion: string) => Promise<TeamsAccess> } | { getToken: typeof getFabricServicePrincipalTokenForTeams }) & {
  poll: (taskHandle: string, accessToken: string) => Promise<FabricAgentTaskStatus>;
};

// SWA injects its own Authorization bearer token. Never use it as Teams identity.
export function teamsAuthorization(headers: { get(name: string): string | null | undefined }) {
  return headers.get("X-DCT-Teams-Authorization");
}

function assertionFrom(authorization: string | null | undefined) {
  return authorization?.match(/^Bearer\s+(\S+)$/i)?.[1];
}

function downstreamError(error: unknown) {
  if (error instanceof TeamsAuthError) {
    return {
      status: error.status,
      jsonBody: { error: error.message },
    };
  }
  if (error instanceof ConversationHistoryError) {
    return { status: 503, jsonBody: { error: error.message } };
  }
  return {
    status: 502,
    jsonBody: { error: "Fabric could not answer this question. Check your data access or try again shortly." },
  };
}

export async function handleTeamsChat(
  authorization: string | null | undefined,
  body: unknown,
  dependencies: StartDependencies = {
    getAccess: getFabricServicePrincipalTokenAndTeamsUser,
    start: startFabricAgentTask,
  },
) {
  const assertion = assertionFrom(authorization);
  if (!assertion) {
    return { status: 401, jsonBody: { error: "Open this app in Teams to use your Teams account." } };
  }
  const question = body && typeof body === "object" && "question" in body && typeof body.question === "string"
    ? body.question.trim() : "";
  if (!question || question.length > 4_000) {
    return { status: 400, jsonBody: { error: "Enter a question between 1 and 4,000 characters." } };
  }
  try {
    if ("getToken" in dependencies) {
      const result = await dependencies.start(question, await dependencies.getToken(assertion));
      return { status: result.status === "working" ? 202 : 200, jsonBody: result };
    }
    const { accessToken, user } = await dependencies.getAccess(assertion);
    const result = await dependencies.start(questionWithConversation(question, await readConversation(user)), accessToken);
    if (result.status === "working") await rememberTask(user, result.taskHandle, question);
    else await appendConversation(user, question, result.answer);
    return { status: result.status === "working" ? 202 : 200, jsonBody: result };
  } catch (error) {
    return downstreamError(error);
  }
}

export async function handleTeamsChatStatus(
  authorization: string | null | undefined,
  taskHandle: string | undefined,
  dependencies: PollDependencies = {
    getAccess: getFabricServicePrincipalTokenAndTeamsUser,
    poll: pollFabricAgentTask,
  },
) {
  const assertion = assertionFrom(authorization);
  if (!assertion) {
    return { status: 401, jsonBody: { error: "Open this app in Teams to use your Teams account." } };
  }
  if (!taskHandle) {
    return { status: 400, jsonBody: { error: "The Fabric task handle is missing." } };
  }

  try {
    if ("getToken" in dependencies) {
      const result = await dependencies.poll(taskHandle, await dependencies.getToken(assertion));
      return { status: result.status === "working" ? 202 : 200, jsonBody: result };
    }
    const { accessToken, user } = await dependencies.getAccess(assertion);
    const question = await taskForUser(user, taskHandle);
    if (!question) return { status: 404, jsonBody: { error: "This task is unavailable or belongs to another Teams user." } };
    const result = await dependencies.poll(taskHandle, accessToken);
    if (result.status === "completed") {
      await appendConversation(user, question, result.answer);
      await forgetTask(taskHandle);
    }
    return { status: result.status === "working" ? 202 : 200, jsonBody: result };
  } catch (error) {
    return downstreamError(error);
  }
}

export async function handleTeamsChatHistory(authorization: string | null | undefined, method: "GET" | "DELETE") {
  const assertion = assertionFrom(authorization);
  if (!assertion) return { status: 401, jsonBody: { error: "Open this app in Teams to use your Teams account." } };
  try {
    const { user } = await getFabricServicePrincipalTokenAndTeamsUser(assertion);
    if (method === "DELETE") {
      await clearConversation(user);
      return { status: 204, jsonBody: undefined };
    }
    return { status: 200, jsonBody: { messages: await readConversation(user) } };
  } catch (error) { return downstreamError(error); }
}
