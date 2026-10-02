import { ConfidentialClientApplication } from "@azure/msal-node";
import jwt, { type JwtPayload } from "jsonwebtoken";
import { createPublicKey, type JsonWebKey, type KeyObject } from "node:crypto";
import { get as httpsGet } from "node:https";

export class TeamsAuthError extends Error {
  constructor(public readonly status: number, message: string) {
    super(message);
  }
}

export type TeamsAuthConfig = {
  clientId: string;
  tenantId: string;
  fabricClientId: string;
  fabricTenantId: string;
  fabricClientSecret: string;
  fabricScope: string;
};

export type TeamsUserIdentity = { tenantId: string; objectId: string };

export const teamsClientIds = [
  "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
  "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
];
type SigningJwk = JsonWebKey & { kid?: string; alg?: string; use?: string };
export type TeamsSigningKeySet = { keys: SigningJwk[] };
type CachedKeySet = { document: TeamsSigningKeySet; expiresAt: number };
const keySets = new Map<string, Promise<CachedKeySet>>();

function readConfig(): TeamsAuthConfig {
  const clientId = process.env.TEAMS_CLIENT_ID?.trim() ?? "";
  const tenantId = process.env.TEAMS_TENANT_ID?.trim() ?? "";
  const fabricClientId = process.env.FABRIC_SERVICE_PRINCIPAL_CLIENT_ID?.trim() || clientId;
  const fabricTenantId = process.env.FABRIC_SERVICE_PRINCIPAL_TENANT_ID?.trim() || tenantId;
  const fabricClientSecret = process.env.FABRIC_SERVICE_PRINCIPAL_CLIENT_SECRET?.trim()
    || process.env.TEAMS_CLIENT_SECRET?.trim()
    || "";
  const fabricScope = process.env.FABRIC_SERVICE_PRINCIPAL_SCOPE?.trim()
    || process.env.TEAMS_FABRIC_SCOPE?.trim()
    || "https://api.fabric.microsoft.com/.default";
  const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  if (
    !guid.test(clientId) || !guid.test(tenantId) ||
    !guid.test(fabricClientId) || !guid.test(fabricTenantId) ||
    !fabricClientSecret || !fabricScope
  ) {
    throw new TeamsAuthError(503, "The Teams or Fabric service principal is not configured yet.");
  }
  return { clientId, tenantId, fabricClientId, fabricTenantId, fabricClientSecret, fabricScope };
}

function downloadKeySet(tenantId: string): Promise<CachedKeySet> {
  const url = `https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`;
  return new Promise((resolve, reject) => {
    const request = httpsGet(url, { headers: { Accept: "application/json" } }, (response) => {
      if (response.statusCode !== 200) {
        response.resume();
        reject(new Error(`Signing key endpoint returned HTTP ${response.statusCode ?? "unknown"}`));
        return;
      }
      const chunks: Buffer[] = [];
      let length = 0;
      response.on("data", (chunk: Buffer) => {
        length += chunk.length;
        if (length > 1_000_000) {
          request.destroy(new Error("Signing key response exceeded the size limit"));
          return;
        }
        chunks.push(chunk);
      });
      response.on("end", () => {
        try {
          const document = JSON.parse(Buffer.concat(chunks).toString("utf8")) as TeamsSigningKeySet;
          if (!document || !Array.isArray(document.keys) || document.keys.length === 0) {
            throw new Error("Signing key response was invalid");
          }
          resolve({ document, expiresAt: Date.now() + 10 * 60_000 });
        } catch (error) {
          reject(error);
        }
      });
    });
    request.setTimeout(10_000, () => request.destroy(new Error("Signing key request timed out")));
    request.on("error", reject);
  });
}

async function cachedKeySet(tenantId: string, forceRefresh = false): Promise<CachedKeySet> {
  let pending = forceRefresh ? undefined : keySets.get(tenantId);
  if (pending) {
    const cached = await pending;
    if (cached.expiresAt > Date.now()) return cached;
  }
  pending = downloadKeySet(tenantId);
  keySets.set(tenantId, pending);
  try {
    return await pending;
  } catch (error) {
    keySets.delete(tenantId);
    throw error;
  }
}

function signingKey(document: TeamsSigningKeySet, kid: string): KeyObject | undefined {
  const jwk = document.keys.find((candidate) =>
    candidate.kid === kid &&
    candidate.kty === "RSA" &&
    (candidate.use === undefined || candidate.use === "sig") &&
    (candidate.alg === undefined || candidate.alg === "RS256"));
  return jwk ? createPublicKey({ key: jwk, format: "jwk" }) : undefined;
}

async function resolveSigningKey(token: string, tenantId: string, override?: TeamsSigningKeySet) {
  const decoded = jwt.decode(token, { complete: true });
  if (!decoded || typeof decoded === "string" || decoded.header.alg !== "RS256" || !decoded.header.kid) {
    throw new jwt.JsonWebTokenError("invalid token header");
  }
  if (override) {
    const key = signingKey(override, decoded.header.kid);
    if (!key) throw new jwt.JsonWebTokenError("signing key not found");
    return key;
  }
  let cached = await cachedKeySet(tenantId);
  let key = signingKey(cached.document, decoded.header.kid);
  if (!key) {
    cached = await cachedKeySet(tenantId, true);
    key = signingKey(cached.document, decoded.header.kid);
  }
  if (!key) throw new jwt.JsonWebTokenError("signing key not found");
  return key;
}

export async function validateTeamsToken(
  token: string,
  config: Pick<TeamsAuthConfig, "clientId" | "tenantId">,
  keySetOverride?: TeamsSigningKeySet,
) {
  let payload: JwtPayload;
  try {
    const key = await resolveSigningKey(token, config.tenantId, keySetOverride);
    const verified = jwt.verify(token, key, {
      algorithms: ["RS256"],
      issuer: `https://login.microsoftonline.com/${config.tenantId}/v2.0`,
      audience: config.clientId,
      clockTolerance: 30,
    });
    if (typeof verified === "string") throw new jwt.JsonWebTokenError("invalid token payload");
    payload = verified;
  } catch (error) {
    const message = error instanceof Error ? error.message.toLowerCase() : "";
    const reason = error instanceof jwt.TokenExpiredError ? "expired"
      : error instanceof jwt.NotBeforeError ? "not-active"
      : message.includes("audience") ? "aud"
      : message.includes("issuer") ? "iss"
      : message.includes("signature") ? "invalid-signature"
      : message.includes("signing key") ? "signing-key-not-found"
      : message.includes("header") || message.includes("malformed") ? "invalid-token-format"
      : message.includes("algorithm") ? "algorithm"
      : "signing-key-service";
    throw new TeamsAuthError(
      401,
      `Your Teams session could not be verified (validation: ${reason}). Reopen this app in Teams and try again.`,
    );
  }

  let invalidClaim: string | undefined;
  const now = Math.floor(Date.now() / 1000);
  if (typeof payload.exp !== "number") invalidClaim = "expiry";
  else if (typeof payload.iat !== "number" || payload.iat > now + 300) invalidClaim = "issued-at";
  else if (payload.ver !== "2.0") invalidClaim = "version";
  else if (payload.tid !== config.tenantId) invalidClaim = "tenant";
  else if (typeof payload.oid !== "string" || !payload.oid) invalidClaim = "user";
  else if (typeof payload.scp !== "string" || !payload.scp.split(" ").includes("access_as_user")) invalidClaim = "scope";
  else if (typeof payload.azp !== "string" || !teamsClientIds.includes(payload.azp)) invalidClaim = "client";
  if (invalidClaim) {
    throw new TeamsAuthError(
      401,
      `Your Teams session could not be verified (validation: ${invalidClaim}). Reopen this app in Teams and try again.`,
    );
  }
  return payload;
}

async function getFabricServicePrincipalToken(config: TeamsAuthConfig): Promise<string> {
  // Fabric receives one application identity; the Teams assertion only authorizes this API call.
  const client = new ConfidentialClientApplication({
    auth: {
      clientId: config.fabricClientId,
      authority: `https://login.microsoftonline.com/${config.fabricTenantId}`,
      clientSecret: config.fabricClientSecret,
    },
  });
  try {
    const result = await client.acquireTokenByClientCredential({
      scopes: [config.fabricScope],
    });
    if (!result?.accessToken) throw new Error("No downstream token");
    return result.accessToken;
  } catch {
    // Never expose token-service error payloads, assertions, or credentials to clients/logs.
    throw new TeamsAuthError(502, "The assistant service identity could not connect to Fabric.");
  }
}

export async function getFabricServicePrincipalTokenAndTeamsUser(assertion: string): Promise<{
  accessToken: string;
  user: TeamsUserIdentity;
}> {
  const config = readConfig();
  const payload = await validateTeamsToken(assertion, config);
  return {
    accessToken: await getFabricServicePrincipalToken(config),
    user: { tenantId: config.tenantId, objectId: payload.oid as string },
  };
}

export async function getFabricServicePrincipalTokenForTeams(assertion: string): Promise<string> {
  return (await getFabricServicePrincipalTokenAndTeamsUser(assertion)).accessToken;
}
