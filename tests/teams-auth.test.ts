import assert from "node:assert/strict";
import { before, test } from "node:test";
import { exportJWK, generateKeyPair, SignJWT } from "jose";
import { teamsClientIds, validateTeamsToken, TeamsAuthError, type TeamsSigningKeySet } from "../api/src/teamsAuth.ts";
import { handleTeamsChat, handleTeamsChatStatus, teamsAuthorization } from "../api/src/teamsChat.ts";

test("SWA platform Authorization cannot substitute for a Teams token", async () => {
  const headers = new Headers({ Authorization: "Bearer platform-token" });
  const never = async (): Promise<never> => { throw new Error("Must not be called"); };
  const result = await handleTeamsChat(teamsAuthorization(headers), { question: "test" }, {
    getToken: never, start: never,
  });
  assert.equal(result.status, 401);
  assert.equal(result.jsonBody.error, "Open this app in Teams to use your Teams account.");
});

test("Teams assertion survives platform Authorization replacement and is still verified", async () => {
  const headers = new Headers({
    Authorization: "Bearer platform-token",
    "X-DCT-Teams-Authorization": "Bearer teams-token",
  });
  let observed = "";
  const result = await handleTeamsChat(teamsAuthorization(headers), { question: "test" }, {
    getToken: async (assertion) => {
      observed = assertion;
      throw new TeamsAuthError(401, "Invalid session");
    },
    start: async () => { throw new Error("Must not be called"); },
  });
  assert.equal(observed, "teams-token");
  assert.equal(result.status, 401);
});

const config = {
  clientId: "54a740da-a7cc-40f3-8951-247f8ff1c307",
  tenantId: "a05fd237-40b5-40b4-bae8-c577a96df96c",
};
let privateKey: CryptoKey;
let keys: TeamsSigningKeySet;

before(async () => {
  const pair = await generateKeyPair("RS256");
  privateKey = pair.privateKey;
  keys = { keys: [{ ...await exportJWK(pair.publicKey), kid: "test-key", alg: "RS256" }] };
});

async function token(overrides: Record<string, unknown> = {}) {
  return new SignJWT({
    iss: `https://login.microsoftonline.com/${config.tenantId}/v2.0`,
    aud: config.clientId,
    tid: config.tenantId,
    oid: "test-user-a",
    scp: "access_as_user",
    azp: teamsClientIds[0],
    ver: "2.0",
    iat: Math.floor(Date.now() / 1000),
    exp: Math.floor(Date.now() / 1000) + 300,
    ...overrides,
  }).setProtectedHeader({ alg: "RS256", kid: "test-key" }).sign(privateKey);
}

test("accepts a signed, scoped Teams user token", async () => {
  const claims = await validateTeamsToken(await token(), config, keys);
  assert.equal(claims.oid, "test-user-a");
});

for (const [label, overrides] of Object.entries({
  audience: { aud: "another-application" },
  issuer: { iss: "https://attacker.example/v2.0" },
  tenant: { tid: "another-tenant" },
  scope: { scp: "User.Read" },
  "app-only token": { scp: undefined, roles: ["access_as_user"] },
  "missing expiry": { exp: undefined },
  expiry: { exp: 1 },
  client: { azp: "unapproved-client" },
  user: { oid: "" },
  version: { ver: "1.0" },
})) {
  test(`rejects incorrect ${label}`, async () => {
    await assert.rejects(validateTeamsToken(await token(overrides), config, keys),
      (error: unknown) => error instanceof TeamsAuthError && error.status === 401);
  });
}

test("rejects a token signed with a different key", async () => {
  const other = await generateKeyPair("RS256");
  const otherKeys = { keys: [{ ...await exportJWK(other.publicKey), kid: "test-key", alg: "RS256" }] };
  await assert.rejects(validateTeamsToken(await token(), config, otherKeys), TeamsAuthError);
});

test("validates the request before acquiring tokens or querying Fabric", async () => {
  const never = async (): Promise<never> => { throw new Error("Must not be called"); };
  const deps = { getToken: never, start: never };
  assert.equal((await handleTeamsChat(null, { question: "test" }, deps)).status, 401);
  assert.equal((await handleTeamsChat("Bearer test", null, deps)).status, 400);
  assert.equal((await handleTeamsChat("Bearer test", { question: "x".repeat(4001) }, deps)).status, 400);
});

test("authentication failure never queries Fabric", async () => {
  let queried = false;
  const result = await handleTeamsChat("Bearer invalid", { question: "test" }, {
    getToken: async () => { throw new TeamsAuthError(401, "Invalid session"); },
    start: async () => {
      queried = true;
      return { status: "completed" as const, answer: "no", toolName: "agent" };
    },
  });
  assert.equal(result.status, 401);
  assert.equal(queried, false);
});

test("simultaneous users use the shared service identity token without returning it", async () => {
  const observed: Record<string, string> = {};
  const deps = {
    getToken: async () => "shared-fabric-token",
    start: async (question: string, accessToken: string) => {
      observed[question] = accessToken;
      return { status: "completed" as const, answer: `answer-${question}`, toolName: "agent" };
    },
  };
  const results = await Promise.all([
    handleTeamsChat("Bearer alice", { question: "alice" }, deps),
    handleTeamsChat("Bearer bob", { question: "bob" }, deps),
  ]);
  assert.deepEqual(observed, { alice: "shared-fabric-token", bob: "shared-fabric-token" });
  assert.ok(results.every((result) => result.status === 200));
  assert.ok(!JSON.stringify(results).includes("shared-fabric-token"));
});

test("downstream errors do not leak assertions or provider details", async () => {
  const result = await handleTeamsChat("Bearer private-assertion", { question: "test" }, {
    getToken: async () => "private-fabric-token",
    start: async () => { throw new Error("private-fabric-token private-assertion provider detail"); },
  });
  assert.equal(result.status, 502);
  assert.ok(!JSON.stringify(result).includes("private-"));
});

test("returns a pollable handle when Fabric starts a background task", async () => {
  const result = await handleTeamsChat("Bearer teams-token", { question: "test" }, {
    getToken: async () => "fabric-token",
    start: async () => ({ status: "working", taskHandle: "opaque-handle", retryAfterMs: 2_000 }),
  });
  assert.equal(result.status, 202);
  assert.deepEqual(result.jsonBody, {
    status: "working",
    taskHandle: "opaque-handle",
    retryAfterMs: 2_000,
  });
});

test("polling revalidates Teams authentication before reading a Fabric task", async () => {
  let observedToken = "";
  const result = await handleTeamsChatStatus("Bearer teams-token", "opaque-handle", {
    getToken: async () => "fabric-token",
    poll: async (taskHandle, accessToken) => {
      observedToken = accessToken;
      assert.equal(taskHandle, "opaque-handle");
      return { status: "completed", answer: "done" };
    },
  });
  assert.equal(observedToken, "fabric-token");
  assert.equal(result.status, 200);
  assert.deepEqual(result.jsonBody, { status: "completed", answer: "done" });
});

test("polling rejects missing Teams authentication and task handles", async () => {
  const never = async (): Promise<never> => { throw new Error("Must not be called"); };
  const dependencies = { getToken: never, poll: never };
  assert.equal((await handleTeamsChatStatus(null, "handle", dependencies)).status, 401);
  assert.equal((await handleTeamsChatStatus("Bearer token", undefined, dependencies)).status, 400);
});
