import { useState, useRef, useEffect } from "react";
import { useAuth0 } from "@auth0/auth0-react";

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
  const { user, logout } = useAuth0();
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [dark, setDark] = useState(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  const getExcelContext = async (): Promise<string> => {
    try {
      return await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(["values", "address"]);
        await context.sync();
        const value = range.values[0][0];
        return value !== null && value !== ""
          ? `[Selected cell ${range.address}: ${value}]`
          : "";
      });
    } catch {
      return "";
    }
  };

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

    const excelContext = await getExcelContext();

    // TODO: Replace with your actual API call
    // const res = await fetch("/api/chat", {
    //   method: "POST",
    //   headers: { "Content-Type": "application/json" },
    //   body: JSON.stringify({ message: text, context: excelContext }),
    // });
    // const data = await res.json();

    setTimeout(() => {
      const assistantMessage: Message = {
        id: (Date.now() + 1).toString(),
        role: "assistant",
        content: excelContext
          ? `${excelContext}\n\nI can help you with that. Connect your AI backend to get real responses.`
          : "I can help you with that. Connect your AI backend to get real responses.",
      };
      setMessages((prev) => [...prev, assistantMessage]);
      setIsLoading(false);
    }, 800);
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
                {messages.map((msg) => (
                  <div key={msg.id} className={`chat-msg chat-msg-${msg.role}`}>
                    <div className="chat-msg-bubble">{msg.content}</div>
                  </div>
                ))}
                {isLoading && (
                  <div className="chat-msg chat-msg-assistant">
                    <div className="chat-msg-bubble chat-typing">
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
            <div className="chat-input-footer-left">
              <button className="chat-footer-icon-btn" title="Attach">
                <svg width="15" height="15" viewBox="0 0 16 16" fill="none">
                  <path
                    d="M13 8.5V11a5 5 0 01-10 0V5a3 3 0 016 0v5.5a1.5 1.5 0 01-3 0V6"
                    stroke="currentColor"
                    strokeWidth="1.4"
                    strokeLinecap="round"
                  />
                </svg>
              </button>
              <button className="chat-footer-icon-btn" title="Excel context">
                <svg width="15" height="15" viewBox="0 0 16 16" fill="none">
                  <rect x="2" y="2" width="12" height="12" rx="2" stroke="currentColor" strokeWidth="1.4" />
                  <path d="M5 8h6M8 5v6" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
                </svg>
              </button>
            </div>
            <button
              className="chat-send-btn"
              onClick={() => sendMessage(input)}
              disabled={!input.trim() || isLoading}
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
          </div>
        </div>

      </div>
    </div>
  );
}
