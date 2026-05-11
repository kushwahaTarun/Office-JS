import { useState, useRef, useEffect, useCallback } from "react";
import { useAuth0 } from "@auth0/auth0-react";
import { runAgent } from "../services/agent";
import { cn } from "../lib/utils";

type Message = {
  id: string;
  role: "user" | "assistant";
  content: string;
};

const QUICK_ACTIONS = [
  { label: "Build financial models", icon: "??" },
  { label: "Explain this workbook", icon: "??" },
  { label: "Transform messy data", icon: "?" },
  { label: "Debug my formulas", icon: "???" },
];

/* -- Sparkle / brand star icon -- */
function StarIcon({ className }: { className?: string }) {
  return (
    <svg
      className={className}
      width="14"
      height="14"
      viewBox="0 0 16 16"
      fill="none"
    >
      <path
        d="M8 1l1.5 4.5L14 7l-4.5 1.5L8 13l-1.5-4.5L2 7l4.5-1.5z"
        fill="currentColor"
      />
    </svg>
  );
}

/* -- Single message bubble -- */
function MessageBubble({ msg, dark }: { msg: Message; dark: boolean }) {
  const isUser = msg.role === "user";
  return (
    <div
      className={cn(
        "msg-enter flex max-w-[88%] flex-col",
        isUser ? "self-end" : "self-start",
      )}
    >
      {!isUser && (
        <div className="mb-1 ml-1 flex items-center gap-1.5">
          <div className="flex h-5 w-5 items-center justify-center rounded-full bg-gradient-to-br from-brand-500 to-brand-700 shadow-sm">
            <StarIcon className="text-white" />
          </div>
          <span
            className={cn(
              "text-[10px] font-medium",
              dark ? "text-gray-400" : "text-gray-400",
            )}
          >
            Fluid AI
          </span>
        </div>
      )}
      <div
        className={cn(
          "rounded-2xl px-3.5 py-2.5 text-[13px] leading-relaxed whitespace-pre-wrap break-words transition-all duration-200",
          isUser
            ? "rounded-br-sm bg-gradient-to-br from-brand-500 to-brand-700 text-white shadow-md shadow-brand-500/20"
            : dark
              ? "rounded-bl-sm border border-white/8 bg-white/6 text-gray-100 shadow-sm"
              : "rounded-bl-sm border border-gray-100 bg-white text-gray-800 shadow-sm",
        )}
      >
        {msg.content}
      </div>
    </div>
  );
}

/* -- Typing indicator bubble -- */
function TypingBubble({ dark }: { dark: boolean }) {
  return (
    <div className="msg-enter self-start flex max-w-[88%] flex-col">
      <div className="mb-1 ml-1 flex items-center gap-1.5">
        <div className="flex h-5 w-5 items-center justify-center rounded-full bg-gradient-to-br from-brand-500 to-brand-700 shadow-sm">
          <StarIcon className="pulse-glow text-white" />
        </div>
        <span className="text-[10px] font-medium text-gray-400">Fluid AI</span>
      </div>
      <div
        className={cn(
          "flex items-center gap-1.5 rounded-2xl rounded-bl-sm px-4 py-3 shadow-sm",
          dark
            ? "border border-white/8 bg-white/6"
            : "border border-gray-100 bg-white",
        )}
      >
        <span
          className={cn(
            "typing-dot",
            dark ? "bg-brand-400" : "bg-brand-400/70",
          )}
        />
        <span
          className={cn(
            "typing-dot",
            dark ? "bg-brand-400" : "bg-brand-400/70",
          )}
        />
        <span
          className={cn(
            "typing-dot",
            dark ? "bg-brand-400" : "bg-brand-400/70",
          )}
        />
      </div>
    </div>
  );
}

/* -- Main component -- */
export default function ChatComponent() {
  const { user, logout, getIdTokenClaims, getAccessTokenSilently } = useAuth0();
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [dark, setDark] = useState(false);
  const [tenant, setTenant] = useState<string | null>(null);
  const [showUserMenu, setShowUserMenu] = useState(false);
  const [attachedFile, setAttachedFile] = useState<File | null>(null);

  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const abortControllerRef = useRef<AbortController | null>(null);
  const userMenuRef = useRef<HTMLDivElement>(null);

  /* close user menu on outside click */
  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (
        userMenuRef.current &&
        !userMenuRef.current.contains(e.target as Node)
      ) {
        setShowUserMenu(false);
      }
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  useEffect(() => {
    getIdTokenClaims().then((claims) => {
      if (claims) {
        const c = claims as Record<string, unknown>;
        const tv =
          c["tenant_value"] ??
          (c["fluidai_app_metadata"] as Record<string, unknown> | undefined)?.[
            "tenant"
          ];
        setTenant(
          Array.isArray(tv)
            ? (tv[0] ?? null)
            : typeof tv === "string"
              ? tv
              : null,
        );
      }
    });

    const setupSelectionListener = () => {
      try {
        if (
          typeof Office !== "undefined" &&
          Office.context &&
          Office.context.document &&
          Office.context.host === Office.HostType.Excel
        ) {
          Office.context.document.addHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            () => {
              console.log("Selection changed");
            },
            (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Selection handler registered successfully");
              } else {
                console.warn(
                  "Failed to register selection handler:",
                  result.error,
                );
              }
            },
          );
        }
      } catch (err) {
        console.warn("Could not add selection handler:", err);
      }
    };

    const timeoutId = setTimeout(setupSelectionListener, 100);
    return () => clearTimeout(timeoutId);
  }, [getIdTokenClaims]);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  const sendMessage = useCallback(
    async (text: string) => {
      if (!text.trim() || isLoading) return;

      const userMessage: Message = {
        id: Date.now().toString(),
        role: "user",
        content: text.trim(),
      };

      setMessages((prev) => [...prev, userMessage]);
      setInput("");
      if (textareaRef.current) textareaRef.current.style.height = "auto";
      setIsLoading(true);

      const controller = new AbortController();
      abortControllerRef.current = controller;

      const assistantId = (Date.now() + 1).toString();
      setMessages((prev) => [
        ...prev,
        { id: assistantId, role: "assistant", content: "" },
      ]);

      try {
        let freshToken: string | null = null;
        try {
          freshToken = await getAccessTokenSilently();
        } catch {
          /* fallback */
        }

        if (!freshToken) {
          throw new Error(
            "Auth token not available. Please sign out and sign in again.",
            {
              cause: "CAUGHT_ERROR: AUTH ERROR",
            },
          );
        }

        let resolvedTenant: string | null = null;
        try {
          const payload = JSON.parse(
            atob(
              freshToken.split(".")[1].replace(/-/g, "+").replace(/_/g, "/"),
            ),
          );
          const meta = payload["fluidai_app_metadata"] as
            | Record<string, unknown>
            | undefined;
          const tv = payload["tenant_value"] ?? meta?.["tenant"];
          resolvedTenant = Array.isArray(tv)
            ? (tv[0] ?? null)
            : typeof tv === "string"
              ? tv
              : null;
        } catch {
          resolvedTenant = tenant;
        }

        if (!resolvedTenant) {
          throw new Error(
            "Tenant not available. Please sign out and sign in again.",
            {
              cause: "CAUGHT_ERROR: AUTH ERROR",
            },
          );
        }

        const answer = await runAgent(
          text.trim(),
          freshToken,
          resolvedTenant,
          (status) => {
            setMessages((prev) =>
              prev.map((msg) =>
                msg.id === assistantId ? { ...msg, content: status } : msg,
              ),
            );
          },
          controller.signal,
        );

        setMessages((prev) =>
          prev.map((msg) =>
            msg.id === assistantId ? { ...msg, content: answer } : msg,
          ),
        );
      } catch (err) {
        const error = err as Error;
        const isCancelled = error.name === "AbortError";
        setMessages((prev) =>
          prev.map((msg) =>
            msg.id === assistantId
              ? {
                  ...msg,
                  content: isCancelled
                    ? "Stopped."
                    : error.message || "Something went wrong.",
                }
              : msg,
          ),
        );
      } finally {
        abortControllerRef.current = null;
        setIsLoading(false);
      }
    },
    [isLoading, getAccessTokenSilently, tenant],
  );

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage(input);
    }
  };

  const handleInputChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setInput(e.target.value);
    if (textareaRef.current) {
      textareaRef.current.style.height = "auto";
      textareaRef.current.style.height = `${Math.min(textareaRef.current.scrollHeight, 100)}px`;
    }
  };

  const avatarLetter = user?.name?.[0]?.toUpperCase() ?? "U";

  const canSend = input.trim().length > 0 && !isLoading;

  return (
    <div
      className={cn(
        "flex h-full w-full flex-col overflow-hidden font-sans transition-colors duration-300",
        dark ? "bg-surface-950 text-gray-100" : "bg-surface-50 text-gray-900",
      )}
    >
      {/* -- Top header bar -- */}
      <header
        className={cn(
          "flex flex-shrink-0 items-center justify-between px-3.5 py-2.5 border-b transition-colors duration-300",
          dark
            ? "border-white/8 bg-surface-900/80 backdrop-blur-md"
            : "border-gray-100 bg-white/80 backdrop-blur-md shadow-sm",
        )}
      >
        {/* Left: brand + beta badge */}
        <div className="flex items-center gap-2">
          <div className="flex h-7 w-7 items-center justify-center rounded-lg bg-gradient-to-br from-brand-500 to-brand-700 shadow-md shadow-brand-500/30">
            <StarIcon className="text-white" />
          </div>
          <span
            className={cn(
              "text-[11px] font-semibold tracking-wide",
              dark ? "text-gray-200" : "text-gray-700",
            )}
          >
            Fluid AI
          </span>
          <span
            className={cn(
              "rounded-full px-1.5 py-0.5 text-[9px] font-semibold uppercase tracking-widest",
              dark
                ? "bg-brand-500/20 text-brand-400 ring-1 ring-brand-500/30"
                : "bg-brand-50 text-brand-600 ring-1 ring-brand-200",
            )}
          >
            Beta
          </span>
        </div>

        {/* Right: new chat / theme toggle / avatar */}
        <div className="flex items-center gap-1">
          {/* New chat */}
          <button
            onClick={() => setMessages([])}
            title="New chat"
            className={cn(
              "focus-brand flex h-7 w-7 items-center justify-center rounded-lg transition-all duration-150",
              dark
                ? "text-gray-400 hover:bg-white/8 hover:text-gray-200"
                : "text-gray-400 hover:bg-gray-100 hover:text-gray-700",
            )}
          >
            <svg width="14" height="14" viewBox="0 0 16 16" fill="none">
              <circle
                cx="8"
                cy="8"
                r="6.3"
                stroke="currentColor"
                strokeWidth="1.4"
              />
              <path
                d="M8 5.5v5M5.5 8h5"
                stroke="currentColor"
                strokeWidth="1.4"
                strokeLinecap="round"
              />
            </svg>
          </button>

          {/* Theme toggle */}
          <button
            onClick={() => setDark((d) => !d)}
            title={dark ? "Light mode" : "Dark mode"}
            className={cn(
              "focus-brand flex h-7 w-7 items-center justify-center rounded-lg transition-all duration-150",
              dark
                ? "text-gray-400 hover:bg-white/8 hover:text-gray-200"
                : "text-gray-400 hover:bg-gray-100 hover:text-gray-700",
            )}
          >
            {dark ? (
              <svg width="13" height="13" viewBox="0 0 16 16" fill="none">
                <circle
                  cx="8"
                  cy="8"
                  r="3"
                  stroke="currentColor"
                  strokeWidth="1.4"
                />
                <path
                  d="M8 1v2M8 13v2M1 8h2M13 8h2M3.05 3.05l1.42 1.42M11.53 11.53l1.42 1.42M3.05 12.95l1.42-1.42M11.53 4.47l1.42-1.42"
                  stroke="currentColor"
                  strokeWidth="1.4"
                  strokeLinecap="round"
                />
              </svg>
            ) : (
              <svg width="13" height="13" viewBox="0 0 16 16" fill="none">
                <path
                  d="M13.5 10.5A6 6 0 015.5 2.5a6 6 0 000 11 6 6 0 008-3z"
                  stroke="currentColor"
                  strokeWidth="1.4"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            )}
          </button>

          {/* Avatar / user menu */}
          <div ref={userMenuRef} className="relative">
            <button
              onClick={() => setShowUserMenu((v) => !v)}
              title={`Signed in as ${user?.email}`}
              className="focus-brand flex h-7 w-7 items-center justify-center overflow-hidden rounded-full ring-2 ring-brand-200 transition-all duration-150 hover:ring-brand-400"
            >
              {user?.picture ? (
                <img
                  src={user.picture}
                  alt={user.name ?? "User"}
                  className="h-full w-full object-cover"
                  onError={(e) => {
                    (e.target as HTMLImageElement).style.display = "none";
                  }}
                />
              ) : (
                <div className="flex h-full w-full items-center justify-center bg-gradient-to-br from-amber-400 to-orange-500 text-[10px] font-bold text-white">
                  {avatarLetter}
                </div>
              )}
            </button>

            {/* Dropdown */}
            {showUserMenu && (
              <div
                className={cn(
                  "absolute right-0 top-9 z-50 min-w-[180px] rounded-xl p-1 shadow-2xl ring-1 transition-all duration-150",
                  dark ? "glass-dark ring-white/10" : "glass ring-black/5",
                )}
                style={{ animation: "fade-up 0.15s ease-out forwards" }}
              >
                <div
                  className={cn(
                    "rounded-lg px-3 py-2 text-[11px]",
                    dark ? "text-gray-300" : "text-gray-500",
                  )}
                >
                  <p
                    className="font-semibold text-[12px]"
                    style={{ color: dark ? "#f1f5f9" : "#111827" }}
                  >
                    {user?.name ?? "User"}
                  </p>
                  <p className="truncate mt-0.5">{user?.email}</p>
                </div>
                <div
                  className={cn(
                    "my-1 h-px",
                    dark ? "bg-white/8" : "bg-gray-100",
                  )}
                />
                <button
                  onClick={() =>
                    logout({
                      logoutParams: { returnTo: window.location.origin },
                    })
                  }
                  className={cn(
                    "flex w-full items-center gap-2 rounded-lg px-3 py-2 text-[12px] font-medium transition-colors duration-150",
                    dark
                      ? "text-gray-300 hover:bg-white/8 hover:text-red-400"
                      : "text-gray-600 hover:bg-red-50 hover:text-red-600",
                  )}
                >
                  <svg width="12" height="12" viewBox="0 0 16 16" fill="none">
                    <path
                      d="M10 3h3a1 1 0 011 1v8a1 1 0 01-1 1h-3M7 11l3-3-3-3M10 8H2"
                      stroke="currentColor"
                      strokeWidth="1.5"
                      strokeLinecap="round"
                      strokeLinejoin="round"
                    />
                  </svg>
                  Sign out
                </button>
              </div>
            )}
          </div>
        </div>
      </header>

      {/* -- Messages area -- */}
      <div
        className={cn(
          "scrollbar-thin flex flex-1 flex-col gap-3 overflow-y-auto px-3.5 py-4 min-h-0",
          dark ? "bg-surface-950" : "bg-surface-50",
        )}
      >
        {messages.length === 0 ? (
          /* Empty / welcome state */
          <div className="flex flex-1 flex-col items-center justify-center gap-5 py-8 text-center">
            {/* Glowing icon */}
            <div className="relative">
              <div
                className={cn(
                  "absolute inset-0 rounded-2xl blur-xl",
                  dark ? "bg-brand-500/30" : "bg-brand-400/20",
                )}
              />
              <div className="relative flex h-12 w-12 items-center justify-center rounded-2xl bg-gradient-to-br from-brand-500 to-brand-700 shadow-xl shadow-brand-500/30">
                <StarIcon className="pulse-glow text-white" />
              </div>
            </div>

            <div className="flex flex-col gap-1">
              <p
                className={cn(
                  "text-sm font-semibold",
                  dark ? "text-gray-200" : "text-gray-700",
                )}
              >
                How can I help?
              </p>
              <p
                className={cn(
                  "text-[11px]",
                  dark ? "text-gray-500" : "text-gray-400",
                )}
              >
                Ask me anything about your spreadsheet
              </p>
            </div>

            {/* Quick action chips */}
            <div className="flex w-full flex-col gap-2">
              {QUICK_ACTIONS.map(({ label, icon }) => (
                <button
                  key={label}
                  onClick={() => sendMessage(label)}
                  className={cn(
                    "focus-brand group flex items-center gap-2.5 rounded-xl px-3.5 py-2.5 text-left text-[12.5px] font-medium ring-1 transition-all duration-200 hover:-translate-y-0.5",
                    dark
                      ? "bg-white/4 text-gray-200 ring-white/8 hover:bg-white/8 hover:ring-brand-500/30 hover:shadow-lg hover:shadow-brand-500/10"
                      : "bg-white text-gray-700 ring-gray-100 hover:bg-brand-50 hover:ring-brand-200 hover:shadow-md hover:shadow-brand-500/8",
                  )}
                >
                  <span className="text-base">{icon}</span>
                  {label}
                  <svg
                    className={cn(
                      "ml-auto h-3.5 w-3.5 -translate-x-1 opacity-0 transition-all duration-200 group-hover:translate-x-0 group-hover:opacity-100",
                      dark ? "text-brand-400" : "text-brand-500",
                    )}
                    viewBox="0 0 16 16"
                    fill="none"
                  >
                    <path
                      d="M3 8h10M9 4l4 4-4 4"
                      stroke="currentColor"
                      strokeWidth="1.5"
                      strokeLinecap="round"
                      strokeLinejoin="round"
                    />
                  </svg>
                </button>
              ))}
            </div>
          </div>
        ) : (
          <>
            {messages.map((msg) =>
              msg.content ? (
                <MessageBubble key={msg.id} msg={msg} dark={dark} />
              ) : null,
            )}
            {isLoading && <TypingBubble dark={dark} />}
            <div ref={messagesEndRef} />
          </>
        )}
      </div>

      {/* -- Input area -- */}
      <div
        className={cn(
          "flex-shrink-0 px-3 pb-3 pt-1.5 transition-colors duration-300",
          dark ? "bg-surface-950" : "bg-surface-50",
        )}
      >
        <div
          className={cn(
            "rounded-2xl ring-1 transition-all duration-200 focus-within:ring-2",
            dark
              ? "bg-surface-800 ring-white/10 focus-within:ring-brand-500/50"
              : "bg-white ring-gray-200 focus-within:ring-brand-400/60 shadow-sm",
          )}
        >
          {/* Attachment tag */}
          {attachedFile && (
            <div
              className={cn(
                "mx-3 mt-2.5 flex items-center gap-1.5 rounded-lg px-2.5 py-1.5 text-[11px] ring-1",
                dark
                  ? "bg-white/6 text-gray-400 ring-white/10"
                  : "bg-gray-50 text-gray-500 ring-gray-200",
              )}
            >
              <svg width="10" height="10" viewBox="0 0 16 16" fill="none">
                <path
                  d="M4 2h6l4 4v8a1 1 0 01-1 1H4a1 1 0 01-1-1V3a1 1 0 011-1z"
                  stroke="currentColor"
                  strokeWidth="1.4"
                />
                <path
                  d="M10 2v4h4"
                  stroke="currentColor"
                  strokeWidth="1.4"
                  strokeLinecap="round"
                />
              </svg>
              <span className="min-w-0 flex-1 truncate">
                {attachedFile.name}
              </span>
              <button
                onClick={() => setAttachedFile(null)}
                className="ml-1 flex-shrink-0 rounded p-0.5 opacity-60 transition-opacity hover:opacity-100"
                aria-label="Remove"
              >
                <svg width="10" height="10" viewBox="0 0 10 10" fill="none">
                  <path
                    d="M1.5 1.5l7 7M8.5 1.5l-7 7"
                    stroke="currentColor"
                    strokeWidth="1.4"
                    strokeLinecap="round"
                  />
                </svg>
              </button>
            </div>
          )}

          {/* Textarea */}
          <textarea
            ref={textareaRef}
            value={input}
            onChange={handleInputChange}
            onKeyDown={handleKeyDown}
            placeholder="Ask anything about your spreadsheet "
            rows={1}
            disabled={isLoading}
            className={cn(
              "block w-full resize-none bg-transparent px-3.5 pt-3 pb-1 text-[13px] leading-relaxed outline-none transition-colors duration-200 placeholder:text-gray-400 disabled:cursor-not-allowed",
              dark
                ? "text-gray-100 placeholder:text-gray-600"
                : "text-gray-800",
            )}
            style={{ maxHeight: "100px", overflowY: "auto" }}
          />

          {/* Footer actions */}
          <div className="flex items-center justify-between px-2.5 pb-2.5 pt-1">
            {/* Upload button */}
            <input
              ref={fileInputRef}
              type="file"
              accept="image/*,video/*,.pdf,.xlsx,.xls,.csv"
              style={{ display: "none" }}
              onChange={(e) => {
                const file = e.target.files?.[0] ?? null;
                setAttachedFile(file);
                e.target.value = "";
              }}
            />
            <button
              onClick={() => fileInputRef.current?.click()}
              title="Attach file"
              className={cn(
                "focus-brand flex h-7 w-7 items-center justify-center rounded-lg transition-all duration-150",
                dark
                  ? "text-gray-500 hover:bg-white/8 hover:text-gray-300"
                  : "text-gray-400 hover:bg-gray-100 hover:text-gray-600",
              )}
            >
              <svg width="14" height="14" viewBox="0 0 16 16" fill="none">
                <path
                  d="M2 11.5V13a1 1 0 001 1h10a1 1 0 001-1v-1.5"
                  stroke="currentColor"
                  strokeWidth="1.4"
                  strokeLinecap="round"
                />
                <path
                  d="M8 2v8M5 5l3-3 3 3"
                  stroke="currentColor"
                  strokeWidth="1.4"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            </button>

            {/* Send / stop button */}
            {isLoading ? (
              <button
                onClick={() => abortControllerRef.current?.abort()}
                aria-label="Stop generation"
                className="flex h-7 w-7 items-center justify-center rounded-lg bg-gradient-to-br from-red-500 to-rose-600 text-white shadow-md shadow-red-500/25 transition-all duration-150 hover:-translate-y-0.5 hover:shadow-lg"
              >
                <svg width="10" height="10" viewBox="0 0 16 16" fill="none">
                  <rect
                    x="3"
                    y="3"
                    width="10"
                    height="10"
                    rx="2"
                    fill="currentColor"
                  />
                </svg>
              </button>
            ) : (
              <button
                onClick={() => sendMessage(input)}
                disabled={!canSend}
                aria-label="Send message"
                className={cn(
                  "flex h-7 w-7 items-center justify-center rounded-lg transition-all duration-200",
                  canSend
                    ? "bg-gradient-to-br from-brand-500 to-brand-700 text-white shadow-md shadow-brand-500/25 hover:-translate-y-0.5 hover:shadow-lg hover:shadow-brand-500/35"
                    : dark
                      ? "bg-white/8 text-gray-600 cursor-not-allowed"
                      : "bg-gray-100 text-gray-300 cursor-not-allowed",
                )}
              >
                <svg width="13" height="13" viewBox="0 0 16 16" fill="none">
                  <path
                    d="M8 13V3M3 8l5-5 5 5"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />
                </svg>
              </button>
            )}
          </div>
        </div>

        {/* Hint text */}
        <p
          className={cn(
            "mt-1.5 text-center text-[10px]",
            dark ? "text-gray-700" : "text-gray-300",
          )}
        >
          Press{" "}
          <kbd className="rounded px-1 py-0.5 font-mono text-[9px] ring-1 ring-current">
            Enter
          </kbd>{" "}
          to send &middot;{" "}
          <kbd className="rounded px-1 py-0.5 font-mono text-[9px] ring-1 ring-current">
            Shift+Enter
          </kbd>{" "}
          for new line
        </p>
      </div>
    </div>
  );
}
