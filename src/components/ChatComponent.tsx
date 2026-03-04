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
    // Example:
    // const res = await fetch("/api/chat", {
    //   method: "POST",
    //   headers: { "Content-Type": "application/json" },
    //   body: JSON.stringify({ message: text, context: excelContext }),
    // });
    // const data = await res.json();
    // assistantContent = data.reply;

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
      textareaRef.current.style.height = `${Math.min(textareaRef.current.scrollHeight, 120)}px`;
    }
  };

  const avatarLetter = user?.name?.[0]?.toUpperCase() ?? "U";

  return (
    <div className="chat-root">
      {/* Header */}
      <div className="chat-header">
        <div className="chat-header-left">
          <span className="chat-title">Fluid AI GPT</span>
          <span className="chat-beta">Beta</span>
        </div>
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

      {/* Messages */}
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

      {/* Input */}
      <div className="chat-input-wrapper">
        <div className="chat-input-box">
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
          <button
            className="chat-send-btn"
            onClick={() => sendMessage(input)}
            disabled={!input.trim() || isLoading}
            aria-label="Send"
          >
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
              <path
                d="M8 13V3M3 8l5-5 5 5"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            </svg>
          </button>
        </div>
      </div>
    </div>
  );
}
