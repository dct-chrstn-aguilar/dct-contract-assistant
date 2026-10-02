# DCT Assistant as a Teams personal tab with a shared Fabric identity

This branch embeds the existing DCT interface as a Teams personal tab. Teams SSO silently identifies the caller, while the backend uses one Microsoft Entra service principal to query the published Fabric Data Agent. Testers do not grant `DataAgent.Execute.All` and do not receive a **Connect to Fabric** prompt.

## Security model

- The API validates every Teams token's signature, issuer, tenant, audience, expiry, `access_as_user` scope, and authorized Teams client.
- The Teams token authorizes access to this API only. It is not exchanged for a user Fabric token.
- The backend requests a Fabric token with the service principal's client ID and secret.
- Fabric evaluates every question as the same service principal. All testers therefore receive the same Fabric, Data Agent, semantic-model, and source permissions granted to that identity.
- Do not use this mode when Fabric must enforce different data permissions for different testers. Use delegated OBO authentication for that requirement.

Service-principal authentication for Fabric Data Agents is a preview feature. It requires paid F2-or-higher Fabric capacity or P1-or-higher Power BI Premium capacity. Managed identities are not currently supported for this Data Agent authentication mode.

## Deployment values

| Setting | Value |
| --- | --- |
| Preview website | `https://proud-mushroom-09f4b1800-1.eastasia.6.azurestaticapps.net` |
| Production website | `https://proud-mushroom-09f4b1800.6.azurestaticapps.net` |
| Teams/Entra client ID | `54a740da-a7cc-40f3-8951-247f8ff1c307` |
| Tenant ID | `a05fd237-40b5-40b4-bae8-c577a96df96c` |
| Fabric MCP endpoint | `https://api.fabric.microsoft.com/v1/mcp/workspaces/f67b0ccc-1f83-430c-93b4-0cbfc45c7f3a/dataagents/5071fab0-9cf5-40d0-9e0c-26dc4d9398a2/agent` |
| Fabric token scope | `https://api.fabric.microsoft.com/.default` |

The existing Entra application can serve as both the Teams SSO API and the Fabric service principal. A separate service principal is also supported through the `FABRIC_SERVICE_PRINCIPAL_*` settings.

## 1. Keep Teams tab SSO configured

Run the helper for the URL used by the Teams package:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\configure-teams-sso.ps1 -TabUrl "https://proud-mushroom-09f4b1800-1.eastasia.6.azurestaticapps.net"
```

This exposes `access_as_user` for the DCT API and pre-authorizes the Teams desktop/mobile and web clients. It does not request a delegated Fabric permission. Users should receive their Teams token silently.

## 2. Enable service principals in Fabric

A Fabric tenant administrator must open **Fabric Admin portal > Tenant settings > Developer settings** and enable **Service principals can use Fabric APIs**. Scope the setting to a security group containing the DCT service principal where possible.

This is a Fabric tenant control. It does not grant tenant-wide delegated consent to the Teams application.

## 3. Grant the service principal only the required Fabric access

In the workspace that hosts the published Data Agent:

1. Open **Manage access**.
2. Add the service principal by its application name.
3. Assign **Member** or **Contributor** as required by the Fabric Data Agent preview.
4. Grant the service principal read access to every attached source: lakehouse, warehouse, semantic model, mirrored database, or ontology.
5. For a Power BI semantic model, grant at least **Read** permission and ensure any model security design behaves correctly for a service identity.

Sharing only the Data Agent is insufficient when the service principal cannot read an underlying source. KQL-backed Data Agents do not currently support service-principal authentication.

## 4. Create and store the service-principal credential

If reusing the DCT application, create a credential using the organization's expiry policy:

```powershell
az ad app credential reset `
  --id "54a740da-a7cc-40f3-8951-247f8ff1c307" `
  --append `
  --display-name "DCT Fabric service principal" `
  --years 1 `
  --query password `
  -o tsv
```

The output is a secret. Put it directly in the Static Web App environment settings. Never place it in GitHub variables, source control, screenshots, or a `VITE_*` variable.

Configure the preview environment and production environment separately:

| Name | Value |
| --- | --- |
| `TEAMS_CLIENT_ID` | `54a740da-a7cc-40f3-8951-247f8ff1c307` |
| `TEAMS_TENANT_ID` | `a05fd237-40b5-40b4-bae8-c577a96df96c` |
| `FABRIC_SERVICE_PRINCIPAL_CLIENT_ID` | Service principal application ID; may be the Teams client ID |
| `FABRIC_SERVICE_PRINCIPAL_TENANT_ID` | `a05fd237-40b5-40b4-bae8-c577a96df96c` |
| `FABRIC_SERVICE_PRINCIPAL_CLIENT_SECRET` | Credential value, not its secret ID |
| `FABRIC_SERVICE_PRINCIPAL_SCOPE` | `https://api.fabric.microsoft.com/.default` |
| `FABRIC_MCP_URL` | Published endpoint shown above |

For backward compatibility, if the three `FABRIC_SERVICE_PRINCIPAL_*` identity settings are absent, the API reuses `TEAMS_CLIENT_ID`, `TEAMS_TENANT_ID`, and `TEAMS_CLIENT_SECRET`. Explicit Fabric settings are recommended.

## 5. Build and deploy

```powershell
npm ci
npm --prefix api ci
npm run build
npm --prefix api run build
npm run lint
npm test
```

Push this branch to update the pull-request preview. The frontend sends the signed Teams token in `X-DCT-Teams-Authorization` because Static Web Apps replaces the standard `Authorization` header. The API validates that Teams token and uses its own service-principal token only for the Fabric MCP call.

The Teams package does not need to be rebuilt when only backend authentication changes and its preview URL remains unchanged. Reload the Teams tab after deployment so it loads the latest frontend bundle.

## Acceptance checks

- The app opens in Teams without another Microsoft sign-in prompt.
- Asking a question does not display **Connect to Fabric**.
- `/api/health` returns `authentication: "service-principal"` and `configured: true`.
- A user without a valid DCT Teams SSO token cannot call `/api/teams/chat`.
- A valid tester can query only the data granted to the shared service principal.
- Removing the service principal from the workspace or source causes Fabric queries to fail.
- No token or client secret appears in browser responses, logs, or source control.

## References

- [Fabric Data Agent service-principal authentication](https://learn.microsoft.com/en-us/fabric/data-science/data-agent-service-principal)
- [Teams tab SSO setup](https://learn.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/tab-sso-register-aad)
- [Teams SSO token validation](https://learn.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/tab-sso-code)
- [Fabric Data Agent MCP](https://learn.microsoft.com/en-us/fabric/data-science/data-agent-mcp-server)
