import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { MsalProvider } from "@azure/msal-react";
import { msalInstance } from "./auth";
import App from "./App";
import { initializeTeams } from "./teams";
import "./styles.css";

const teamsHost = await initializeTeams();
await msalInstance.initialize();
const redirectResult = teamsHost.kind === "browser" ? await msalInstance.handleRedirectPromise() : null;
if (redirectResult?.account) msalInstance.setActiveAccount(redirectResult.account);
if (!msalInstance.getActiveAccount()) {
  msalInstance.setActiveAccount(msalInstance.getAllAccounts()[0] ?? null);
}

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <MsalProvider instance={msalInstance}>
      <App teamsHost={teamsHost} />
    </MsalProvider>
  </StrictMode>,
);
