import { useState, useRef, useEffect } from "react";
import { useAuth0 } from "@auth0/auth0-react";
import { runAgent } from "../services/agent";

type Message = {
  id: string;
  role: "user" | "assistant";
  content: string;
};

const QUICK_ACTIONS = [
  "Build financial models",
  "Explain this workbook",
  "Transform messy data",
  "Debug my formulas",
];

export default function ChatComponent() {
  const { user, logout, getIdTokenClaims, getAccessTokenSilently } = useAuth0();
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [dark, setDark] = useState(false);
  const [tenant, setTenant] = useState<string | null>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const abortControllerRef = useRef<AbortController | null>(null);
  const [attachedFile, setAttachedFile] = useState<File | null>(null);

  useEffect(() => {
    getIdTokenClaims().then((claims) => {
      if (claims) {
        const c = claims as Record<string, unknown>;
        const tv =
          c["tenant_value"] ??
          (c["fluidai_app_metadata"] as Record<string, unknown> | undefined)?.["tenant"];
        setTenant(Array.isArray(tv) ? (tv[0] ?? null) : typeof tv === "string" ? tv : null);
      }
    });

    // Selection listener - Office.js is already initialized by main.tsx
    // So we can directly check if context is available
    const setupSelectionListener = () => {
      try {
        // Check if Office context is available and we're in Excel
        if (
          typeof Office !== "undefined" &&
          Office.context &&
          Office.context.document &&
          Office.context.host === Office.HostType.Excel
        ) {
          Office.context.document.addHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            () => {
              // Trigger a silent context refresh or just let the next agent run pick it up
              // For now, we don't need to do much because runAgent calls getWorkbookContext() at the start.
              console.log("Selection changed");
            },
            (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Selection handler registered successfully");
              } else {
                console.warn("Failed to register selection handler:", result.error);
              }
            }
          );
        }
      } catch (err) {
        console.warn("Could not add selection handler:", err);
      }
    };

    // Try to set up listener after a small delay to ensure Office.js is fully ready
    const timeoutId = setTimeout(setupSelectionListener, 100);

    return () => {
      clearTimeout(timeoutId);
      // Note: Office.js doesn't provide a clean way to remove event handlers in task pane add-ins
    };
  }, [getIdTokenClaims]);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  const sendMessage = async (text: string) => {
    if (!text.trim() || isLoading) return;

    const userMessage: Message = {
      id: Date.now().toString(),
      role: "user",
      content: text.trim(),
    };

    setMessages((prev) => [...prev, userMessage]);
    setInput("");
    if (textareaRef.current) {
      textareaRef.current.style.height = "auto";
    }
    setIsLoading(true);
    const controller = new AbortController();
    abortControllerRef.current = controller;

    const assistantId = (Date.now() + 1).toString();
    setMessages((prev) => [...prev, { id: assistantId, role: "assistant", content: "" }]);

    try {
      let freshToken: string | null = null;
      try {
        freshToken = await getAccessTokenSilently();
      } catch {
        // fall back to cached token if silent refresh fails
      }

      if (!freshToken) {
        throw new Error("Auth token not available. Please sign out and sign in again.", {
          cause: "CAUGHT_ERROR: AUTH ERROR",
        });
      }

      let resolvedTenant: string | null = null;
      try {
        const payload = JSON.parse(atob(freshToken.split(".")[1].replace(/-/g, "+").replace(/_/g, "/")));
        const meta = payload["fluidai_app_metadata"] as Record<string, unknown> | undefined;
        const tv = payload["tenant_value"] ?? meta?.["tenant"];
        resolvedTenant = Array.isArray(tv) ? (tv[0] ?? null) : typeof tv === "string" ? tv : null;
      } catch {
        // fall back to state if decode fails
        resolvedTenant = tenant;
      }

      if (!resolvedTenant) {
        throw new Error("Tenant not available. Please sign out and sign in again.", {
          cause: "CAUGHT_ERROR: AUTH ERROR",
        });
      }

      const answer = await runAgent(
        text.trim(),
        freshToken,
        resolvedTenant,
        (status) => {
          setMessages((prev) =>
            prev.map((msg) =>
              msg.id === assistantId ? { ...msg, content: status } : msg
            )
          );
        },
        controller.signal
      );
      setMessages((prev) =>
        prev.map((msg) =>
          msg.id === assistantId ? { ...msg, content: answer } : msg
        )
      );
    } catch (err) {
      const error = err as Error;
      const isCancelled = error.name === "AbortError";
      setMessages((prev) =>
        prev.map((msg) =>
          msg.id === assistantId
            ? { ...msg, content: isCancelled ? "Stopped." : error.message || "Something went wrong." }
            : msg
        )
      );
    } finally {
      abortControllerRef.current = null;
      setIsLoading(false);
    }
  };

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

  return (
    <div className={`chat-root${dark ? " dark" : ""}`}>
      <div className="chat-main">

        {/* Rounded card — messages area */}
        <div className="chat-card">
          <div className="chat-card-bar">
            <div className="chat-card-bar-left">
              <span className="chat-beta">Beta</span>
            </div>
            <div className="chat-card-bar-right">
              {/* New chat */}
              <button
                className="chat-icon-btn"
                onClick={() => setMessages([])}
                title="New chat"
              >
                <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                  <circle cx="8" cy="8" r="6.3" stroke="currentColor" strokeWidth="1.4" />
                  <path d="M8 5.5v5M5.5 8h5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
                </svg>
              </button>

              {/* Dark / light mode toggle */}
              <button
                className="chat-icon-btn"
                onClick={() => setDark((d) => !d)}
                title={dark ? "Switch to light mode" : "Switch to dark mode"}
              >
                {dark ? (
                  /* Sun icon */
                  <svg width="15" height="15" viewBox="0 0 16 16" fill="none">
                    <circle cx="8" cy="8" r="3" stroke="currentColor" strokeWidth="1.4" />
                    <path d="M8 1v2M8 13v2M1 8h2M13 8h2M3.05 3.05l1.42 1.42M11.53 11.53l1.42 1.42M3.05 12.95l1.42-1.42M11.53 4.47l1.42-1.42" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
                  </svg>
                ) : (
                  /* Moon icon */
                  <svg width="15" height="15" viewBox="0 0 16 16" fill="none">
                    <path d="M13.5 10.5A6 6 0 015.5 2.5a6 6 0 000 11 6 6 0 008-3z" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
                  </svg>
                )}
              </button>

              {/* Avatar / sign out */}
              <button
                className="chat-avatar-btn"
                onClick={() =>
                  logout({ logoutParams: { returnTo: window.location.origin } })
                }
                title={`Sign out (${user?.email})`}
              >
                {user?.picture ? (
                  <img
                    src={user.picture}
                    alt={user.name ?? "User"}
                    className="chat-avatar-img"
                    onError={(e) => {
                      (e.target as HTMLImageElement).style.display = "none";
                    }}
                  />
                ) : (
                  <div className="chat-avatar-fallback">{avatarLetter}</div>
                )}
              </button>
            </div>
          </div>

          <div className="chat-messages">
            {messages.length === 0 ? (
              <div className="chat-empty">
                <div className="chat-empty-icon">✦</div>
                <p className="chat-empty-text">Take actions in Excel</p>
                <div className="chat-quick-actions">
                  {QUICK_ACTIONS.map((action) => (
                    <button
                      key={action}
                      className="chat-quick-action"
                      onClick={() => sendMessage(action)}
                    >
                      {action}
                    </button>
                  ))}
                </div>
              </div>
            ) : (
              <>
                {messages.map((msg) =>
                  msg.content ? (
                    <div key={msg.id} className={`chat-msg chat-msg-${msg.role}`}>
                      <div className="chat-msg-bubble">{msg.content}</div>
                    </div>
                  ) : null
                )}
                {isLoading && (
                  <div className="chat-msg chat-msg-assistant">
                    <div className="chat-msg-bubble chat-typing">
                      <svg className="chat-typing-icon" width="14" height="14" viewBox="0 0 16 16" fill="none">
                        <path d="M8 1l1.5 4.5L14 7l-4.5 1.5L8 13l-1.5-4.5L2 7l4.5-1.5z" fill="currentColor" opacity="0.8" />
                      </svg>
                      <span />
                      <span />
                      <span />
                    </div>
                  </div>
                )}
                <div ref={messagesEndRef} />
              </>
            )}
          </div>
        </div>

        {/* Input card */}
        <div className="chat-input-card">
          {attachedFile && (
            <div className="chat-attachment-tag">
              <svg width="12" height="12" viewBox="0 0 16 16" fill="none">
                <path d="M4 2h6l4 4v8a1 1 0 01-1 1H4a1 1 0 01-1-1V3a1 1 0 011-1z" stroke="currentColor" strokeWidth="1.4" />
                <path d="M10 2v4h4" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
              </svg>
              <span>{attachedFile.name}</span>
              <button
                className="chat-attachment-remove"
                onClick={() => setAttachedFile(null)}
                aria-label="Remove attachment"
              >×</button>
            </div>
          )}
          <textarea
            ref={textareaRef}
            className="chat-textarea"
            value={input}
            onChange={handleInputChange}
            onKeyDown={handleKeyDown}
            placeholder="Reply"
            rows={1}
            disabled={isLoading}
          />
          <div className="chat-input-footer">
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
              className="chat-footer-icon-btn"
              title="Upload media"
              onClick={() => fileInputRef.current?.click()}
            >
              <svg width="15" height="15" viewBox="0 0 16 16" fill="none">
                <path d="M2 11.5V13a1 1 0 001 1h10a1 1 0 001-1v-1.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
                <path d="M8 2v8M5 5l3-3 3 3" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </button>
            {isLoading ? (
              <button
                className="chat-send-btn"
                onClick={() => abortControllerRef.current?.abort()}
                aria-label="Stop"
              >
                <svg width="12" height="12" viewBox="0 0 16 16" fill="none">
                  <rect x="3" y="3" width="10" height="10" rx="1.5" fill="currentColor" />
                </svg>
              </button>
            ) : (
              <button
                className="chat-send-btn"
                onClick={() => sendMessage(input)}
                disabled={!input.trim()}
                aria-label="Send"
              >
                <svg width="14" height="14" viewBox="0 0 16 16" fill="none">
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

      </div>
    </div>
  );
}
