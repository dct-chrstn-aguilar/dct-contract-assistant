import { useEffect, useRef, useState, type FormEvent, type KeyboardEvent, type ReactNode } from "react";
import { useMsal } from "@azure/msal-react";
import {
  Check,
  ChevronsUpDown,
  CircleAlert,
  Copy,
  LogIn,
  LogOut,
  RotateCcw,
} from "lucide-react";
import Sidebar, { type Section } from "./Sidebar";
import AssetIcon, { type AssetIconName } from "./AssetIcon";
import { useTheme } from "./theme";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { fabricScopes, getFabricToken, isAuthConfigured } from "./auth";
import { getTeamsToken, type TeamsHost } from "./teams";

type Message = {
  id: string;
  role: "user" | "assistant";
  content: string;
  failed?: boolean;
  responseTimeMs?: number;
};

type ChatResponse = {
  answer?: string;
  error?: string;
  status?: "working" | "completed";
  taskHandle?: string;
  retryAfterMs?: number;
};

type PendingTask = {
  taskHandle?: string;
  question?: string;
  startedAt: number;
};

const maximumTaskWaitMs = 9 * 60_000;
const teamsMessagesStorageKey = "dct-contract-assistant.messages";
const teamsPendingTaskStorageKey = "dct-contract-assistant.pending-task";

function readStoredValue<T>(key: string, fallback: T): T {
  try {
    const value = localStorage.getItem(key);
    return value ? JSON.parse(value) as T : fallback;
  } catch {
    return fallback;
  }
}

const wait = (milliseconds: number) =>
  new Promise<void>((resolve) => window.setTimeout(resolve, milliseconds));

function formatElapsed(milliseconds: number): string {
  const seconds = milliseconds / 1_000;
  return seconds < 10 ? `${seconds.toFixed(1)}s` : `${Math.round(seconds)}s`;
}

async function readChatResponse(response: Response): Promise<ChatResponse> {
  const text = await response.text();

  if (!text.trim()) {
    return { error: `The assistant backend returned HTTP ${response.status} without a response.` };
  }

  try {
    return JSON.parse(text) as ChatResponse;
  } catch {
    return {
      error: response.ok
        ? "The assistant backend returned an invalid response."
        : text.trim(),
    };
  }
}

// Every section asks the same Fabric data agent; the section only changes the starting prompts.
const suggestionsBySection: Record<Section, { icon: AssetIconName; title: string; prompt: string }[]> = {
  chat: [
    { icon: "building-2", title: "Customer record", prompt: "Show me a customer entry" },
    { icon: "contact-round", title: "Account manager", prompt: "Who is the account manager for <company>?" },
    { icon: "mail", title: "Email & phone", prompt: "Show all available contact details for <company>" },
    { icon: "file-text", title: "Contract summary", prompt: "Show contract description for <company>" },
  ],
  jira: [
    { icon: "list-todo", title: "Open tickets", prompt: "List the open Jira tickets for <company>" },
    { icon: "clock", title: "Latest activity", prompt: "Show the most recent Jira issues raised for <company>" },
    { icon: "user-round", title: "Ticket owners", prompt: "Who is assigned to the open Jira tickets for <company>?" },
    { icon: "circle-check", title: "Resolved", prompt: "Which Jira tickets for <company> were resolved this month?" },
  ],
  sales: [
    { icon: "contact-round", title: "Account manager", prompt: "Who is the account manager for <company>?" },
    { icon: "users", title: "Rep portfolio", prompt: "Which customers are managed by the same account manager as <company>?" },
    { icon: "mail", title: "Sales contacts", prompt: "Show all available contact details for <company>" },
    { icon: "file-text", title: "Contract terms", prompt: "Show contract description for <company>" },
  ],
};

const sectionCopy: Record<Section, { lede: string; placeholder: string; strip: string }> = {
  chat: {
    lede: "Ask about customer records, account managers, contact details or contract terms.",
    placeholder: "Ask anything…",
    strip: "Searches DCT customer, contact and contract records",
  },
  jira: {
    lede: "Ask about Jira tickets linked to DCT customers: what's open, who owns it, and what changed.",
    placeholder: "Ask about Jira tickets…",
    strip: "Searches Jira issues connected to DCT customers",
  },
  sales: {
    lede: "Look up account managers, customer portfolios and sales contacts.",
    placeholder: "Ask about accounts and sales contacts…",
    strip: "Searches DCT account ownership and contact records",
  },
};

const sidebarStorageKey = "dct-contract-assistant.sidebar-open";

function initialSidebarOpen(): boolean {
  // Phones start closed so the drawer doesn't cover the chat on first load.
  if (window.matchMedia("(max-width: 767px)").matches) return false;
  try {
    return localStorage.getItem(sidebarStorageKey) !== "false";
  } catch {
    return true;
  }
}

const isMobileViewport = () => window.matchMedia("(max-width: 767px)").matches;

const companyToken = "<company>";

// Renders the <company> placeholder as a fill-in chip so users see what to replace.
function PromptText({ prompt }: { prompt: string }) {
  const [before, after] = prompt.split(companyToken);
  if (after === undefined) return <>{prompt}</>;
  return <>{before}<span className="token">company</span>{after}</>;
}

// Long Fabric queries can run for minutes; staged copy tells users it hasn't stalled.
function thinkingLabel(milliseconds: number): string {
  if (milliseconds < 4_000) return "Reading your question";
  if (milliseconds < 15_000) return "Searching contract records";
  if (milliseconds < 45_000) return "Putting the answer together";
  return "Still working. Complex questions can take a few minutes";
}

const localAuthBypass =
  import.meta.env.DEV && import.meta.env.VITE_LOCAL_AUTH_BYPASS === "true";

function friendlyError(error: unknown): string {
  if (error instanceof Error) {
    if (/popup|interaction.*progress/i.test(error.message)) {
      return "The sign-in window was closed. Please sign in and try again.";
    }
    return error.message;
  }
  return "Something went wrong while contacting the data agent.";
}

export default function App({ teamsHost }: { teamsHost: TeamsHost }) {
  const inTeams = teamsHost.kind === "teams";
  const useLocalAuth = localAuthBypass && !inTeams;
  const baseCanAsk = inTeams ? !teamsHost.error : useLocalAuth || isAuthConfigured;
  const { instance, accounts } = useMsal();
  const account = instance.getActiveAccount() ?? accounts[0];
  const [messages, setMessages] = useState<Message[]>(() =>
    inTeams ? readStoredValue<Message[]>(teamsMessagesStorageKey, []) : [],
  );
  const [pendingTask, setPendingTask] = useState<PendingTask | null>(() => {
    if (!inTeams) return null;
    const storedTask = readStoredValue<PendingTask | null>(teamsPendingTaskStorageKey, null);
    if (storedTask) return storedTask;

    // Recover conversations left by older builds that saved the user message
    // before Teams unloaded the iframe, but did not yet save a task handle.
    const lastMessage = messages.at(-1);
    return lastMessage?.role === "user"
      ? { question: lastMessage.content, startedAt: Date.now() }
      : null;
  });
  const [draft, setDraft] = useState("");
  const [isSending, setIsSending] = useState(Boolean(pendingTask));
  const [elapsedMilliseconds, setElapsedMilliseconds] = useState(() =>
    pendingTask ? Math.max(0, Date.now() - pendingTask.startedAt) : 0,
  );
  const [accountMenuOpen, setAccountMenuOpen] = useState(false);
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [sidebarOpen, setSidebarOpen] = useState(initialSidebarOpen);
  const [section, setSection] = useState<Section>("chat");
  const { theme, toggleTheme } = useTheme();
  const accountAreaRef = useRef<HTMLDivElement>(null);
  const sidebarOpenButtonRef = useRef<HTMLButtonElement>(null);
  const bottomRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const requestStartedAtRef = useRef<number | null>(pendingTask?.startedAt ?? null);
  const activeRequestRef = useRef(false);
  const initialPendingTaskRef = useRef(pendingTask);
  const canAsk = baseCanAsk;

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, isSending]);

  // Grow the composer with its content up to the CSS max-height.
  useEffect(() => {
    const input = inputRef.current;
    if (!input) return;
    input.style.height = "auto";
    input.style.height = `${input.scrollHeight}px`;
  }, [draft]);

  useEffect(() => {
    if (!accountMenuOpen) return;
    const close = (event: PointerEvent | globalThis.KeyboardEvent) => {
      if (event instanceof globalThis.KeyboardEvent ? event.key === "Escape" : !accountAreaRef.current?.contains(event.target as Node)) {
        setAccountMenuOpen(false);
      }
    };
    document.addEventListener("pointerdown", close);
    document.addEventListener("keydown", close);
    return () => {
      document.removeEventListener("pointerdown", close);
      document.removeEventListener("keydown", close);
    };
  }, [accountMenuOpen]);

  useEffect(() => {
    if (!isSending || requestStartedAtRef.current === null) return;

    const updateElapsed = () => {
      if (requestStartedAtRef.current !== null) {
        setElapsedMilliseconds(Date.now() - requestStartedAtRef.current);
      }
    };
    updateElapsed();
    const interval = window.setInterval(updateElapsed, 250);
    return () => window.clearInterval(interval);
  }, [isSending]);

  useEffect(() => {
    if (!inTeams) return;
    localStorage.setItem(teamsMessagesStorageKey, JSON.stringify(messages));
  }, [inTeams, messages]);

  useEffect(() => {
    if (!inTeams) return;
    if (pendingTask) {
      localStorage.setItem(teamsPendingTaskStorageKey, JSON.stringify(pendingTask));
    } else {
      localStorage.removeItem(teamsPendingTaskStorageKey);
    }
  }, [inTeams, pendingTask]);

  const signIn = async () => {
    if (!isAuthConfigured) return undefined;
    try {
      const result = await instance.loginPopup({ scopes: fabricScopes });
      instance.setActiveAccount(result.account);
      return result.account;
    } catch (error) {
      setMessages((current) => [
        ...current,
        {
          id: crypto.randomUUID(),
          role: "assistant",
          content: friendlyError(error),
          failed: true,
        },
      ]);
      return undefined;
    }
  };

  const signOut = async () => {
    setAccountMenuOpen(false);
    await instance.logoutPopup({ account });
  };

  const resetChat = async () => {
    if (isSending) return;
    if (inTeams) {
      try {
        const token = await getTeamsToken();
        const response = await fetch("/api/teams/chat/history", {
          method: "DELETE",
          headers: { "X-DCT-Teams-Authorization": `Bearer ${token}` },
        });
        if (!response.ok) {
          const data = await readChatResponse(response);
          throw new Error(data.error || "The saved conversation could not be cleared.");
        }
      } catch (error) {
        setMessages((current) => [...current, { id: crypto.randomUUID(), role: "assistant", content: friendlyError(error), failed: true }]);
        return;
      }
    }
    setMessages([]);
    setTimeout(() => inputRef.current?.focus(), 0);
  };

  const selectSuggestion = (prompt: string) => {
    setDraft(prompt);
    window.setTimeout(() => {
      const input = inputRef.current;
      if (!input) return;

      input.focus();
      const companyPlaceholder = prompt.indexOf("<company>");
      if (companyPlaceholder >= 0) {
        input.setSelectionRange(companyPlaceholder, companyPlaceholder + "<company>".length);
      }
    }, 0);
  };

  const sendQuestion = async (question: string, activeAccount = account): Promise<string> => {
    const token = inTeams
      ? await getTeamsToken()
      : activeAccount
        ? await getFabricToken(activeAccount)
        : "";
    const headers: Record<string, string> = { "Content-Type": "application/json" };
    if (token) headers[inTeams ? "X-DCT-Teams-Authorization" : "Authorization"] = `Bearer ${token}`;

    let response = await fetch(inTeams ? "/api/teams/chat" : "/api/chat", {
      method: "POST",
      headers,
      body: JSON.stringify({ question }),
    });
    let data = await readChatResponse(response);

    if (inTeams && response.status === 202) {
      const taskHandle = data.taskHandle;
      if (!taskHandle) {
        throw new Error(data.error || "Fabric did not return a background task handle.");
      }

      const task = { taskHandle, question, startedAt: requestStartedAtRef.current ?? Date.now() };
      setPendingTask(task);

      const deadline = task.startedAt + maximumTaskWaitMs;
      while (response.status === 202 && Date.now() < deadline) {
        const retryAfterMs = Math.min(Math.max(data.retryAfterMs ?? 2_000, 1_000), 10_000);
        await wait(retryAfterMs);
        response = await fetch(`/api/teams/chat/status/${encodeURIComponent(taskHandle)}`, {
          method: "GET",
          headers,
        });
        data = await readChatResponse(response);
      }

      if (response.status === 202) {
        throw new Error("Fabric is still processing this question. Please try again shortly.");
      }
    }

    if (!response.ok || !data.answer) {
      throw new Error(data.error || "The data agent could not answer this question.");
    }
    return data.answer;
  };

  const ask = async (question: string) => {
    const trimmed = question.trim();
    if (!trimmed || isSending || !canAsk) return;

    const activeAccount = inTeams || useLocalAuth ? undefined : account ?? (await signIn());
    if (!inTeams && !useLocalAuth && !activeAccount) return;

    const userMessage: Message = {
      id: crypto.randomUUID(),
      role: "user",
      content: trimmed,
    };
    setMessages((current) => [...current, userMessage]);
    setDraft("");
    const startedAt = Date.now();
    const startingTask = { question: trimmed, startedAt };
    requestStartedAtRef.current = startedAt;
    activeRequestRef.current = true;
    if (inTeams) {
      localStorage.setItem(teamsPendingTaskStorageKey, JSON.stringify(startingTask));
      setPendingTask(startingTask);
    }
    setElapsedMilliseconds(0);
    setIsSending(true);

    try {
      const answer = await sendQuestion(trimmed, activeAccount);
      setMessages((current) => [
        ...current,
        {
          id: crypto.randomUUID(),
          role: "assistant",
          content: answer,
          responseTimeMs: Date.now() - startedAt,
        },
      ]);
    } catch (error) {
      setMessages((current) => [
        ...current,
        {
          id: crypto.randomUUID(),
          role: "assistant",
          content: friendlyError(error),
          failed: true,
          responseTimeMs: Date.now() - startedAt,
        },
      ]);
    } finally {
      activeRequestRef.current = false;
      requestStartedAtRef.current = null;
      setPendingTask(null);
      setIsSending(false);
      setTimeout(() => inputRef.current?.focus(), 0);
    }
  };

  const sendQuestionRef = useRef(sendQuestion);
  sendQuestionRef.current = sendQuestion;

  useEffect(() => {
    const taskToResume = initialPendingTaskRef.current;
    if (!inTeams || !taskToResume) return;

    let mounted = true;

    const resume = async () => {
      if (activeRequestRef.current) return;
      activeRequestRef.current = true;
      requestStartedAtRef.current = taskToResume.startedAt;
      setIsSending(true);

      try {
        let answer: string;

        if (!taskToResume.taskHandle) {
          if (!taskToResume.question) {
            throw new Error("The interrupted question could not be restored. Please ask it again.");
          }
          answer = await sendQuestionRef.current(taskToResume.question);
        } else {
          const token = await getTeamsToken();
          const headers = {
            "Content-Type": "application/json",
            "X-DCT-Teams-Authorization": `Bearer ${token}`,
          };
          const deadline = taskToResume.startedAt + maximumTaskWaitMs;
          let data: ChatResponse = { status: "working", retryAfterMs: 1_000 };
          let response: Response | undefined;

          while (Date.now() < deadline) {
            await wait(Math.min(Math.max(data.retryAfterMs ?? 2_000, 1_000), 10_000));
            response = await fetch(`/api/teams/chat/status/${encodeURIComponent(taskToResume.taskHandle)}`, {
              method: "GET",
              headers,
            });
            data = await readChatResponse(response);
            if (response.status !== 202) break;
          }

          if (!response || response.status === 202) {
            throw new Error("Fabric is still processing this question. Please try again shortly.");
          }
          if (!response.ok || !data.answer) {
            throw new Error(data.error || "The data agent could not answer this question.");
          }
          answer = data.answer;
        }

        if (mounted) {
          setMessages((current) => [...current, {
            id: crypto.randomUUID(),
            role: "assistant",
            content: answer,
            responseTimeMs: Date.now() - taskToResume.startedAt,
          }]);
        }
      } catch (error) {
        if (mounted) {
          setMessages((current) => [...current, {
            id: crypto.randomUUID(),
            role: "assistant",
            content: friendlyError(error),
            failed: true,
            responseTimeMs: Date.now() - taskToResume.startedAt,
          }]);
        }
      } finally {
        if (mounted) {
          activeRequestRef.current = false;
          requestStartedAtRef.current = null;
          setPendingTask(null);
          setIsSending(false);
          setTimeout(() => inputRef.current?.focus(), 0);
        }
      }
    };

    // Deferring one tick prevents React Strict Mode's setup/cleanup probe from
    // abandoning the only recovery attempt before it starts.
    const resumeTimer = window.setTimeout(() => void resume(), 0);
    return () => {
      mounted = false;
      window.clearTimeout(resumeTimer);
    };
  }, [inTeams]);

  const handleSubmit = (event: FormEvent) => {
    event.preventDefault();
    void ask(draft);
  };

  const handleKeyDown = (event: KeyboardEvent<HTMLTextAreaElement>) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      void ask(draft);
    }
  };


  const retry = (failedId: string) => {
    if (isSending) return;
    const failedIndex = messages.findIndex((message) => message.id === failedId);
    const question = messages.slice(0, failedIndex).reverse().find((message) => message.role === "user");
    if (!question) return;
    setMessages((current) => current.filter((message) => message.id !== failedId && message.id !== question.id));
    void ask(question.content);
  };

  const copyAnswer = async (message: Message) => {
    try {
      await navigator.clipboard.writeText(message.content);
      setCopiedId(message.id);
      window.setTimeout(() => setCopiedId((current) => (current === message.id ? null : current)), 1_800);
    } catch {
      // Clipboard access can be blocked inside some Teams clients; fail quietly.
    }
  };

  const displayName = inTeams ? teamsHost.name : account?.name || account?.username || "";
  const initials = (displayName || "User")
    .split(/[\s@.]+/)
    .slice(0, 2)
    .map((part: string) => part[0])
    .join("")
    .toUpperCase();
  const firstName = displayName.split(/[\s@.]+/)[0] || "";
  const hasConversation = messages.length > 0;
  const lastMessage = messages.at(-1);
  const lastFailedId = lastMessage?.failed ? lastMessage.id : undefined;
  const showSignInHint = !inTeams && !useLocalAuth && !account && isAuthConfigured;
  const copy = sectionCopy[section];

  const setSidebar = (open: boolean) => {
    const closing = sidebarOpen && !open;
    setSidebarOpen(open);
    setAccountMenuOpen(false);
    if (closing) {
      // Return focus to the control that reopens the drawer instead of leaving
      // focus inside the newly inert sidebar.
      window.setTimeout(() => sidebarOpenButtonRef.current?.focus(), 0);
    }
    try {
      if (!isMobileViewport()) localStorage.setItem(sidebarStorageKey, String(open));
    } catch {
      // Remembering the drawer state is a convenience only.
    }
  };

  // On phones the drawer overlays the chat, so close it once the user has picked something.
  const closeSidebarOnMobile = () => {
    if (isMobileViewport()) setSidebar(false);
  };

  const startNewChat = () => {
    closeSidebarOnMobile();
    void resetChat();
  };

  const selectSection = (next: Section) => {
    setSection(next);
    closeSidebarOnMobile();
    window.setTimeout(() => inputRef.current?.focus(), 0);
  };

  const composer = (
    <div className={`composer-wrap ${isSending ? "composer-wrap--busy" : ""}`}>
      {showSignInHint && (
        <p className="composer-hint">
          <AssetIcon name="user-round" size={14} />
          Sign in with Microsoft to ask questions using your Fabric permissions.
        </p>
      )}
      <div className="composer-glow">
        <span className="composer-glow__halo" aria-hidden="true" />
        <span className="composer-glow__ring" aria-hidden="true" />
        <div className="composer-shell">
          {!hasConversation && (
            <p className="composer-strip">
              <AssetIcon name="zap" size={13} />
              {copy.strip}
            </p>
          )}
          <form className="composer" onSubmit={handleSubmit}>
            <textarea
              ref={inputRef}
              value={draft}
              onChange={(event) => setDraft(event.target.value)}
              onKeyDown={handleKeyDown}
              placeholder={hasConversation ? "Ask a follow-up…" : copy.placeholder}
              rows={1}
              maxLength={4_000}
              disabled={isSending || !canAsk}
              aria-label="Ask the DCT Contract Assistant"
            />
            <div className="composer__bar">
              <span className="composer__note">
                <kbd>Enter</kbd> to send <span aria-hidden="true">·</span> <kbd>Shift</kbd> + <kbd>Enter</kbd> new line
              </span>
              <button className="send-button" type="submit" disabled={!draft.trim() || isSending || !canAsk} aria-label="Send message">
                <AssetIcon name="arrow-up" size={18} />
              </button>
            </div>
          </form>
        </div>
      </div>
      <p className="disclaimer">AI can make mistakes. Verify important customer and contract details.</p>
    </div>
  );

  const identityRow = (label: string, detail: string, avatar: ReactNode) => (
    <>
      <span className="identity__avatar">{avatar}</span>
      <span className="identity__text">
        <strong>{label}</strong>
        <small>{detail}</small>
      </span>
    </>
  );

  const accountControl = inTeams ? (
    <div className="identity">{identityRow(teamsHost.name, "Signed in with Teams", initials)}</div>
  ) : useLocalAuth ? (
    <div className="identity">{identityRow("Local session", "Azure CLI", <AssetIcon name="user-round" size={16} />)}</div>
  ) : !account ? (
    <button type="button" onClick={() => void signIn()} disabled={!isAuthConfigured} className="sign-in-button">
      <LogIn size={16} /> Sign in with Microsoft
    </button>
  ) : (
    <div className="account-area" ref={accountAreaRef}>
      <button
        type="button"
        onClick={() => setAccountMenuOpen((open) => !open)}
        aria-expanded={accountMenuOpen}
        aria-haspopup="menu"
        className="identity identity--button"
      >
        {identityRow(account.name || account.username, account.username, initials)}
        <ChevronsUpDown size={15} className="identity__chevron" />
      </button>
      {accountMenuOpen && (
        <div role="menu" className="account-menu">
          <button type="button" role="menuitem" onClick={() => void signOut()} className="account-menu__item">
            <LogOut size={15} /> Sign out
          </button>
        </div>
      )}
    </div>
  );

  return (
    <div className={`app-shell ${hasConversation ? "app-shell--chat" : "app-shell--home"}`}>
      <Sidebar
        open={sidebarOpen}
        onClose={() => setSidebar(false)}
        onNewChat={startNewChat}
        newChatDisabled={isSending}
        section={section}
        onSelectSection={selectSection}
        footer={accountControl}
      />

      <main className="stage">
        {/* Visible only while the drawer is closed, so it can always be reopened. */}
        <div className={`floating-controls ${sidebarOpen ? "" : "floating-controls--visible"}`} inert={sidebarOpen}>
          <button
            ref={sidebarOpenButtonRef}
            type="button"
            onClick={() => setSidebar(true)}
            aria-label="Open sidebar"
            aria-controls="app-sidebar"
            aria-expanded={sidebarOpen}
            title="Open sidebar"
            className="floating-button"
          >
            <AssetIcon name="panel-left-open" size={18} />
          </button>
          <button type="button" onClick={startNewChat} disabled={isSending} aria-label="New chat" title="New chat" className="floating-button">
            <AssetIcon name="square-pen" size={17} />
          </button>
        </div>

        <button
          type="button"
          onClick={toggleTheme}
          aria-label={theme === "dark" ? "Switch to light mode" : "Switch to dark mode"}
          title={theme === "dark" ? "Light mode" : "Dark mode"}
          className="floating-button theme-toggle"
        >
          {theme === "dark" ? <AssetIcon name="sun" size={17} /> : <AssetIcon name="moon" size={17} />}
        </button>

        <span className="stage__floor" aria-hidden="true" />
        {inTeams && teamsHost.error && (
          <div className="notice" role="alert">
            <CircleAlert size={17} />
            <span>{teamsHost.error}</span>
          </div>
        )}
        {!inTeams && !isAuthConfigured && !useLocalAuth && (
          <div className="notice" role="alert">
            <CircleAlert size={17} />
            <span>Add your Entra and Fabric IDs to <code>.env</code> to connect this assistant.</span>
          </div>
        )}

        {!hasConversation ? (
          <section className="home" aria-labelledby="home-title">
            <div className="orb orb--hero" aria-hidden="true"><span /></div>
            <h1 id="home-title">
              <span className="home__hi">Hi, {firstName || "there"}</span>
              <span className="home__ask">How can I help today?</span>
            </h1>
            <p className="home__lede">
              {copy.lede}
            </p>

            {composer}

            <ul className="starters" aria-label="Example questions">
                {suggestionsBySection[section].map(({ icon, title, prompt }) => (
                <li key={title}>
                  <button onClick={() => selectSuggestion(prompt)} disabled={isSending} aria-label={`Use example prompt: ${prompt}`}>
                    <strong><AssetIcon name={icon} size={16} />{title}</strong>
                    <span><PromptText prompt={prompt} /></span>
                  </button>
                </li>
              ))}
            </ul>
          </section>
        ) : (
          <>
            <section className="thread" aria-live="polite" aria-label="Conversation">
              <div className="thread__inner">
                {messages.map((message) => (
                  message.role === "user" ? (
                    <article key={message.id} className="turn turn--user">
                      <p>{message.content}</p>
                    </article>
                  ) : (
                    <article key={message.id} className={`turn turn--assistant ${message.failed ? "turn--error" : ""}`}>
                      {message.failed ? (
                        <span className="turn__mark turn__mark--error" aria-hidden="true"><CircleAlert size={15} /></span>
                      ) : (
                        <span className="orb orb--mini" aria-hidden="true"><span /></span>
                      )}
                      <div className="turn__body">
                        {message.failed ? (
                          <div className="error-card">
                            <strong>That didn&apos;t go through</strong>
                            <p>{message.content}</p>
                            {message.id === lastFailedId && (
                              <button className="chip-button" onClick={() => retry(message.id)} disabled={isSending || !canAsk}>
                                <RotateCcw size={14} /> Try again
                              </button>
                            )}
                          </div>
                        ) : (
                          <div className="markdown">
                            <ReactMarkdown remarkPlugins={[remarkGfm]}>{message.content}</ReactMarkdown>
                          </div>
                        )}
                        <div className="turn__meta">
                          {message.responseTimeMs !== undefined && (
                            <span>{message.failed ? "Stopped after" : "Answered in"} {formatElapsed(message.responseTimeMs)}</span>
                          )}
                          {!message.failed && (
                            <button className="meta-button" onClick={() => void copyAnswer(message)} aria-label="Copy answer">
                              {copiedId === message.id ? <><Check size={13} /> Copied</> : <><Copy size={13} /> Copy</>}
                            </button>
                          )}
                        </div>
                      </div>
                    </article>
                  )
                ))}
                {isSending && (
                  <article className="turn turn--assistant" aria-busy="true">
                    <span className="orb orb--mini orb--live" aria-hidden="true"><span /></span>
                    <div className="turn__body">
                      <div className="thinking">
                        <span className="thinking__label">{thinkingLabel(elapsedMilliseconds)}</span>
                        <span className="thinking__time">{formatElapsed(elapsedMilliseconds)}</span>
                      </div>
                    </div>
                  </article>
                )}
                <div ref={bottomRef} />
              </div>
            </section>
            <div className="dock">{composer}</div>
          </>
        )}
      </main>
    </div>
  );
}
