import { app, authentication } from "@microsoft/teams-js";

export type TeamsTheme = "default" | "dark" | "contrast";

export type TeamsHost =
  | { kind: "browser" }
  | { kind: "teams"; name: string; theme?: TeamsTheme; error?: string };

// Mirrors the Teams client theme onto <html data-theme> so the tab matches Teams.
export function applyTeamsTheme(theme: string | undefined) {
  const root = document.documentElement;
  if (theme === "dark" || theme === "contrast") root.dataset.theme = "dark";
  else if (theme === "default") root.dataset.theme = "light";
  else delete root.dataset.theme;
}

export const teamsAccessMessage =
  "We couldn't connect your Teams account automatically. Please contact the app administrator to check access.";

export async function initializeTeams(): Promise<TeamsHost> {
  const explicitlyInTeams = new URLSearchParams(window.location.search).get("host") === "teams";
  if (!explicitlyInTeams && window.parent === window) return { kind: "browser" };

  try {
    await app.initialize();
    const context = await app.getContext();
    app.notifyAppLoaded();
    app.notifySuccess();
    const theme = context.app.theme as TeamsTheme | undefined;
    applyTeamsTheme(theme);
    app.registerOnThemeChangeHandler(applyTeamsTheme);
    return { kind: "teams", name: context.user?.displayName || "Teams account", theme };
  } catch {
    // Never launch the website's MSAL popup as a fallback in a Teams-marked tab.
    return explicitlyInTeams
      ? { kind: "teams", name: "Teams account", error: "Open this page from the DCT app in Teams, or reload the tab." }
      : { kind: "browser" };
  }
}

export async function getTeamsToken(): Promise<string> {
  try {
    // Teams is pre-authorized for this API, so testers should not see a sign-in dialog.
    return await authentication.getAuthToken({ silent: true });
  } catch {
    throw new Error(teamsAccessMessage);
  }
}
