# DCT Contract Assistant

A React + Vite chat experience for a published Microsoft Fabric Data Agent. The Teams personal tab validates users through silent Teams SSO and queries Fabric through one shared service principal.

## Architecture

- **React + Vite** renders the chat experience.
- **MSAL Browser** supports the standalone website's delegated sign-in flow.
- An **Azure Functions API** validates Teams callers and obtains a server-side Fabric service-principal token. An Express adapter provides the same API during local development.
- The **Fabric Data Agent MCP endpoint** discovers its tool at runtime, so the app does not hard-code the tool or question argument name.

The standalone browser flow does not require a client secret. The Teams path uses a server-only service-principal credential; no client secret or Fabric token is exposed to the browser.

## Microsoft Teams personal tab

The Teams app embeds this same React interface, rather than a conversational bot. It requests a Teams token silently and sends it to `/api/teams/chat`. The API validates the caller and obtains an app-only Fabric token. All Teams users therefore share the Fabric permissions granted to that service principal.

See [the Teams tab deployment guide](docs/teams-tab-setup.md) for the exact Entra, Azure, and Developer Portal settings. The Azure Bot resource and bot OAuth connection are not needed for this experience.

## Prerequisites

- Node.js 20.19+ or 22.12+
- Azure CLI (for the recommended local development flow)
- A Fabric capacity that supports Data Agents
- A configured and **published** Fabric Data Agent
- Fabric tenant approval for service principals to use Fabric APIs
- Workspace and source access for the shared service principal
- An Entra app registration for the browser sign-in flow or deployment

## 1. Configure Microsoft Entra ID for the standalone browser (optional)

This section applies only when opening the website outside Teams. Teams uses silent tab SSO plus the shared service principal described in the [Teams setup guide](docs/teams-tab-setup.md).

1. Create an app registration in the same Entra tenant as Fabric.
2. Under **Authentication**, add a **Single-page application** platform.
3. Add `http://localhost:5173` as a redirect URI.
4. Under **API permissions**, add the delegated Fabric/Data Agent permission required by your tenant and grant admin consent. The app requests `https://api.fabric.microsoft.com/.default`, which includes the delegated Fabric permissions consented on the registration.

> Fabric Data Agent MCP is a preview feature. If your tenant exposes the permission through the Power BI Service API, use `DataAgent.Execute.All` and grant admin consent.

## 2. Configure the app

Copy the environment template:

```powershell
Copy-Item .env.example .env
```

Set these values in `.env`:

```dotenv
VITE_ENTRA_CLIENT_ID=<application-client-id>
VITE_ENTRA_TENANT_ID=<directory-tenant-id>
FABRIC_MCP_URL=https://api.fabric.microsoft.com/v1/mcp/workspaces/f67b0ccc-1f83-430c-93b4-0cbfc45c7f3a/dataagents/5071fab0-9cf5-40d0-9e0c-26dc4d9398a2/agent
PORT=3001
```

The default token scope follows the direct MCP guidance: `https://api.fabric.microsoft.com/.default`. If your tenant registration uses the Power BI API audience, set `VITE_FABRIC_SCOPE=https://analysis.windows.net/powerbi/api/DataAgent.Execute.All` as directed by your Fabric administrator.

The MCP URL is available in the Data Agent's **Settings > Model Context Protocol** page. The configured endpoint is:

```text
https://api.fabric.microsoft.com/v1/mcp/workspaces/f67b0ccc-1f83-430c-93b4-0cbfc45c7f3a/dataagents/5071fab0-9cf5-40d0-9e0c-26dc4d9398a2/agent
```

## 3. Run locally

### Recommended: Azure CLI authentication (Windows)

This option uses your Azure CLI sign-in and does not require an Entra app registration.

1. Install the Azure CLI if it is not already installed:

   ```powershell
   winget install --exact --id Microsoft.AzureCLI
   ```

2. Close and reopen PowerShell after installation, then verify that both Node.js and the Azure CLI are available:

   ```powershell
   node --version
   az version
   ```

   If Azure CLI is installed but `az` is not recognized, refresh `PATH` in the current PowerShell window:

   ```powershell
   $env:Path = [Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [Environment]::GetEnvironmentVariable("Path", "User")
   az version
   ```

3. Sign in with the Microsoft account that can access the published Fabric Data Agent:

   ```powershell
   az login --allow-no-subscriptions
   ```

4. From the repository root, allow scripts for the current PowerShell process and start the app:

   ```powershell
   Set-ExecutionPolicy -Scope Process Bypass
   .\start-local.ps1
   ```

The launcher installs npm dependencies when needed, obtains a short-lived Fabric access token, and starts both the Vite UI and Express API. Open [http://localhost:5173](http://localhost:5173), and press `Ctrl+C` to stop the app. Run the launcher again if the access token expires.

### Alternative: browser MSAL authentication

Use this option to test the browser sign-in flow. First complete the Entra configuration above and make sure `.env` contains valid `VITE_ENTRA_CLIENT_ID`, `VITE_ENTRA_TENANT_ID`, and Fabric endpoint values. Then run:

```powershell
Copy-Item .env.example .env # Skip this if .env is already configured
npm install
npm run dev
```

Open [http://localhost:5173](http://localhost:5173). Both the Vite UI and API server start together.

## Production

Build the React app and start the Express server:

```powershell
npm run build
$env:NODE_ENV="production"
npm start
```

In production, Express serves the generated `dist` directory and the API from the same origin. Configure the production origin as an SPA redirect URI in the Entra app registration.

### Deploy with the linked Azure Static Web App

The existing GitHub Actions workflow is configured to deploy `dist` and the `api` Azure Functions package. Before pushing to `main`:

1. Add repository variables named `VITE_ENTRA_CLIENT_ID` and `VITE_ENTRA_TENANT_ID` under **GitHub > Settings > Secrets and variables > Actions > Variables**.
2. If needed, add the optional repository variable `VITE_FABRIC_SCOPE`.
3. In the Azure portal, add `FABRIC_MCP_URL` and the server-only `FABRIC_SERVICE_PRINCIPAL_*` settings from the Teams setup guide to the linked Static Web App's **Environment variables**.
4. Add the deployed Static Web App URL as an SPA redirect URI in the Entra app registration.

The browser-visible Entra IDs are intentionally GitHub variables; the MCP URL is configured as an Azure environment variable. None of these values is a client secret.

## Useful checks

```bash
npm run build
npm run lint
```

The API health endpoint is `GET /api/health`. It reports whether the Fabric IDs are present without disclosing them.

## Troubleshooting

- **`az` is not recognized:** close and reopen the terminal after installing Azure CLI, or use the `PATH` refresh command in the local setup instructions.
- **Azure CLI says to run `az login`:** run `az login --allow-no-subscriptions`, complete the browser sign-in, and then run `.\start-local.ps1` again.
- **PowerShell blocks a script:** run `Set-ExecutionPolicy -Scope Process Bypass`; this changes the policy only for the current PowerShell process.
- **Sign-in fails:** confirm the exact browser origin is registered as an SPA redirect URI.
- **Teams 401:** verify the Teams SSO resource, pre-authorized Teams clients, deployed URL, and `TEAMS_CLIENT_ID`/`TEAMS_TENANT_ID` settings.
- **Teams Fabric failure:** verify the Fabric service-principal tenant setting, credential, workspace role, and read access to every attached source.
- **Agent endpoint returns an error:** publish the Data Agent first; staging agents do not expose the MCP runtime.
- **Empty or poor answers:** improve the Data Agent description, instructions, source descriptions, and example queries in Fabric.

The integration follows Microsoft's current [Fabric Data Agent MCP guidance](https://learn.microsoft.com/en-us/fabric/data-science/data-agent-mcp-server).
