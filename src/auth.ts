import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type Configuration,
} from "@azure/msal-browser";

const defaultClientId = "54a740da-a7cc-40f3-8951-247f8ff1c307";
const defaultTenantId = "a05fd237-40b5-40b4-bae8-c577a96df96c";

// Entra application and tenant IDs are public identifiers. Defaults keep branch
// preview builds usable when GitHub repository variables are unavailable.
export const clientId = import.meta.env.VITE_ENTRA_CLIENT_ID?.trim() || defaultClientId;
export const tenantId = import.meta.env.VITE_ENTRA_TENANT_ID?.trim() || defaultTenantId;
export const isAuthConfigured = Boolean(
  clientId && tenantId && !clientId.startsWith("00000000") && !tenantId.startsWith("00000000"),
);

const config: Configuration = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "sessionStorage",
  },
};

export const fabricScopes = [
  import.meta.env.VITE_FABRIC_SCOPE?.trim() || "https://api.fabric.microsoft.com/.default",
];
export const msalInstance = new PublicClientApplication(config);

export async function getFabricToken(account: AccountInfo): Promise<string> {
  let result: AuthenticationResult;

  try {
    result = await msalInstance.acquireTokenSilent({ scopes: fabricScopes, account });
  } catch (error) {
    if (!(error instanceof InteractionRequiredAuthError)) throw error;
    result = await msalInstance.acquireTokenPopup({ scopes: fabricScopes, account });
  }

  return result.accessToken;
}
