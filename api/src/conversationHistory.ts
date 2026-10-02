import { TableClient } from "@azure/data-tables";
import { createHash } from "node:crypto";
import type { TeamsUserIdentity } from "./teamsAuth.js";

const conversationRowKey = "conversation";
const taskPartitionKey = "tasks";
const maxTurns = 12;
const maxTurnCharacters = 6_000;
const maxContextCharacters = 18_000;
const taskLifetimeMs = 10 * 60_000;
const maxAppendRetries = 3;

export type ConversationTurn = { role: "user" | "assistant"; content: string };

export class ConversationHistoryError extends Error {}

function table(): TableClient {
  const connectionString = process.env.HISTORY_STORAGE_CONNECTION_STRING?.trim();
  const tableName = process.env.HISTORY_TABLE_NAME?.trim();
  if (!connectionString || !tableName) {
    throw new ConversationHistoryError("Conversation history is not configured. Contact your administrator.");
  }
  // The table is provisioned in Azure before deployment. Do not call createTable here:
  // a restricted connection string can read/write an existing table without create permission.
  return TableClient.fromConnectionString(connectionString, tableName);
}

function userKey(user: TeamsUserIdentity): string {
  return `user-${createHash("sha256").update(`${user.tenantId}:${user.objectId}`).digest("hex")}`;
}

function taskKey(taskHandle: string): string {
  return createHash("sha256").update(taskHandle).digest("hex");
}

function singaporeDateKey(value: Date = new Date()): string {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Singapore",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(value);
}

function isConversationFromToday(updatedAt: unknown): boolean {
  return typeof updatedAt === "string" && singaporeDateKey(new Date(updatedAt)) === singaporeDateKey();
}

function normaliseTurns(value: unknown): ConversationTurn[] {
  if (typeof value !== "string") return [];
  try {
    const parsed: unknown = JSON.parse(value);
    if (!Array.isArray(parsed)) return [];
    return parsed.filter((turn): turn is ConversationTurn =>
      Boolean(turn) && typeof turn === "object" &&
      ((turn as ConversationTurn).role === "user" || (turn as ConversationTurn).role === "assistant") &&
      typeof (turn as ConversationTurn).content === "string",
    ).map((turn) => ({ ...turn, content: turn.content.slice(0, maxTurnCharacters) })).slice(-maxTurns);
  } catch { return []; }
}

export async function readConversation(user: TeamsUserIdentity): Promise<ConversationTurn[]> {
  try {
    const entity = await table().getEntity<Record<string, unknown>>(userKey(user), conversationRowKey);
    if (!isConversationFromToday(entity.updatedAt)) {
      await clearConversation(user);
      return [];
    }
    return normaliseTurns(entity.turnsJson);
  } catch (error) {
    if ((error as { statusCode?: number }).statusCode === 404) return [];
    throw error;
  }
}

export async function appendConversation(user: TeamsUserIdentity, question: string, answer: string): Promise<void> {
  const client = table();
  const partitionKey = userKey(user);
  for (let attempt = 0; attempt < maxAppendRetries; attempt += 1) {
    let entity: Record<string, unknown> | undefined;
    try {
      entity = await client.getEntity<Record<string, unknown>>(partitionKey, conversationRowKey);
    } catch (error) {
      if ((error as { statusCode?: number }).statusCode !== 404) throw error;
    }
    const turns = entity && isConversationFromToday(entity.updatedAt)
      ? normaliseTurns(entity.turnsJson)
      : [];
    const nextTurns = [...turns,
      { role: "user" as const, content: question.slice(0, maxTurnCharacters) },
      { role: "assistant" as const, content: answer.slice(0, maxTurnCharacters) },
    ].slice(-maxTurns);
    try {
      if (entity) {
        await client.updateEntity({
          partitionKey,
          rowKey: conversationRowKey,
          turnsJson: JSON.stringify(nextTurns),
          updatedAt: new Date().toISOString(),
        }, "Replace", { etag: entity.etag as string });
      } else {
        await client.createEntity({
          partitionKey,
          rowKey: conversationRowKey,
          turnsJson: JSON.stringify(nextTurns),
          updatedAt: new Date().toISOString(),
        });
      }
      return;
    } catch (error) {
      const statusCode = (error as { statusCode?: number }).statusCode;
      if (statusCode !== 409 && statusCode !== 412) throw error;
      if (attempt === maxAppendRetries - 1) throw error;
    }
  }
}

export async function clearConversation(user: TeamsUserIdentity): Promise<void> {
  try { await table().deleteEntity(userKey(user), conversationRowKey); }
  catch (error) { if ((error as { statusCode?: number }).statusCode !== 404) throw error; }
}

export async function rememberTask(user: TeamsUserIdentity, taskHandle: string, question: string): Promise<void> {
  await table().upsertEntity({
    partitionKey: taskPartitionKey, rowKey: taskKey(taskHandle), ownerKey: userKey(user),
    question: question.slice(0, maxTurnCharacters), expiresAt: new Date(Date.now() + taskLifetimeMs).toISOString(),
  }, "Replace");
}

export async function taskForUser(user: TeamsUserIdentity, taskHandle: string): Promise<string | undefined> {
  try {
    const entity = await table().getEntity<Record<string, unknown>>(taskPartitionKey, taskKey(taskHandle));
    if (entity.ownerKey !== userKey(user) || typeof entity.question !== "string") return undefined;
    return typeof entity.expiresAt === "string" && Date.parse(entity.expiresAt) < Date.now() ? undefined : entity.question;
  } catch (error) {
    if ((error as { statusCode?: number }).statusCode === 404) return undefined;
    throw error;
  }
}

export async function forgetTask(taskHandle: string): Promise<void> {
  try { await table().deleteEntity(taskPartitionKey, taskKey(taskHandle)); }
  catch (error) { if ((error as { statusCode?: number }).statusCode !== 404) throw error; }
}

export async function cleanupConversationHistory(): Promise<{ conversations: number; tasks: number }> {
  const client = table();
  const today = singaporeDateKey();
  let conversations = 0;
  let tasks = 0;

  for await (const entity of client.listEntities<Record<string, unknown>>()) {
    if (!entity.partitionKey || !entity.rowKey) continue;
    if (entity.partitionKey === taskPartitionKey) {
      if (typeof entity.expiresAt === "string" && Date.parse(entity.expiresAt) < Date.now()) {
        await client.deleteEntity(entity.partitionKey, entity.rowKey);
        tasks += 1;
      }
    } else if (entity.rowKey === conversationRowKey &&
      typeof entity.updatedAt === "string" && singaporeDateKey(new Date(entity.updatedAt)) !== today) {
      await client.deleteEntity(entity.partitionKey, entity.rowKey);
      conversations += 1;
    }
  }
  return { conversations, tasks };
}

export function questionWithConversation(question: string, turns: ConversationTurn[]): string {
  if (turns.length === 0) return question;
  const history = turns.map((turn) => `${turn.role === "user" ? "User" : "Assistant"}: ${turn.content}`).join("\n").slice(-maxContextCharacters);
  return `Use this history only to resolve follow-up references. Answer the current question using Fabric data.\n\nConversation history:\n${history}\n\nCurrent question:\n${question}`;
}
