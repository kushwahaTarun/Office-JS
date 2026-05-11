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
  { label: "Build financial models", icon: "??", desc: "Create P&L, forecasts, budgets" },
  { label: "Explain this workbook",  icon: "??", desc: "Summarise structure & formulas" },
  { label: "Transform messy data",   icon: "?", desc: "Clean, dedupe, reformat data"  },
  { label: "Debug my formulas",      icon: "???", desc: "Find & fix errors instantly"    },
];

function StarIcon({ className }: { className?: string }) {
  return (
    <svg className={className} width="12" height="12" viewBox="0 0 16 16" fill="none">
      <path d="M8 1l1.5 4.5L14 7l-4.5 1.5L8 13l-1.5-4.5L2 7l4.5-1.5z" fill="currentColor" />
    </svg>
  );
}

/* -- Message bubble -- */
function MessageBubble({ msg, dark }: { msg: Message; dark: boolean }) {
  const isUser = msg.role === "user";
  return (
    <div className={cn("msg-enter flex max-w-[90%] flex-col gap-1", isUser ? "self-end" : "self-start")}>
      {!isUser && (
        <div className="ml-1 flex items-center gap-1.5">
          <div className="flex h-4 w-4 items-center justify-center rounded-full bg-gradient-to-br from-brand-500 to-brand-700 shadow-sm shadow-brand-500/30">
            <StarIcon className="text-white" />
          </div>
          <span className="text-[10px] font-medium text-gray-500">Fluid AI</span>
        </div>
      )}
      <div
        className={cn(
          "rounded-2xl px-3.5 py-2.5 text-[13px] leading-relaxed whitespace-pre-wrap break-words",
          isUser
            ? [
                "rounded-br-[4px]",
                "bg-gradient-to-br from-brand-500 to-brand-700",
                "text-white shadow-lg shadow-brand-500/25",
                "ring-1 ring-white/10",
              ].join(" ")
            : dark
            ? [
                "rounded-bl-[4px]",
                "bg-white/[0.05] text-gray-100",
                "ring-1 ring-white/[0.08] shadow-md",
              ].join(" ")
            : [
                "rounded-bl-[4px]",
                "bg-white text-gray-800",
                "ring-1 ring-gray-100/80 shadow-sm",
              ].join(" ")
        )}
      >
        {msg.content}
      </div>
    </div>
  );
}

/* -- Typing indicator -- */
function TypingBubble({ dark }: { dark: boolean }) {
  return (
    <div className="msg-enter self-start flex max-w-[90%] flex-col gap-1">
      <div className="ml-1 flex items-center gap-1.5">
        <div className="flex h-4 w-4 items-center justify-center rounded-full bg-gradient-to-br from-brand-500 to-brand-700 shadow-sm shadow-brand-500/30">
          <StarIcon className="pulse-glow text-white" />
        </div>
        <span className="text-[10px] font-medium text-gray-500">thinking </span>
      </div>
      <div
        className={cn(
          "flex items-center gap-1.5 rounded-2xl rounded-bl-[4px] px-4 py-3",
          dark
            ? "bg-white/[0.05] ring-1 ring-white/[0.08] shadow-md"
            : "bg-white ring-1 ring-gray-100/80 shadow-sm"
        )}
      >
        <span className="typing-dot bg-brand-400" />
        <span className="typing-dot bg-brand-400" />
        <span className="typing-dot bg-brand-400" />
      </div>
    </div>
  );
}

/* -- Main component -- */
export default function ChatComponent() {
  const { user, logout, getIdTokenClaims, getAccessTokenSilently } = useAuth0();
  const [messages,     setMessages]     = useState<Message[]>([]);
  const [input,        setInput]        = useState("");
  const [isLoading,    setIsLoading]    = useState(false);
  const [dark,         setDark]         = useState(false);
  const [tenant,       setTenant]       = useState<string | null>(null);
  const [showMenu,     setShowMenu]     = useState(false);
  const [attachedFile, setAttachedFile] = useState<File | null>(null);
  const [inputFocused, setInputFocused] = useState(false);

  const messagesEndRef     = useRef<HTMLDivElement>(null);
  const textareaRef        = useRef<HTMLTextAreaElement>(null);
  const fileInputRef       = useRef<HTMLInputElement>(null);
  const abortControllerRef = useRef<AbortController | null>(null);
  const userMenuRef        = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const h = (e: MouseEvent) => {
      if (userMenuRef.current && !userMenuRef.current.contains(e.target as Node))
        setShowMenu(false);
    };
    document.addEventListener("mousedown", h);
    return () => document.removeEventListener("mousedown", h);
  }, []);

  useEffect(() => {
    getIdTokenClaims().then((claims) => {
      if (claims) {
        const c = claims as Record<string, unknown>;
        const tv = c["tenant_value"] ??
          (c["fluidai_app_metadata"] as Record<string, unknown> | undefined)?.["tenant"];
        setTenant(Array.isArray(tv) ? (tv[0] ?? null) : typeof tv === "string" ? tv : null);
      }
    });
    const id = setTimeout(() => {
      try {
        if (
          typeof Office !== "undefined" &&
          Office.context?.document &&
          Office.context.host === Office.HostType.Excel
        ) {
          Office.context.document.addHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            () => console.log("Selection changed"),
            (r) => {
              if (r.status !== Office.AsyncResultStatus.Succeeded)
                console.warn("Selection handler failed:", r.error);
            }
          );
        }
      } catch (err) { console.warn("Office handler:", err); }
    }, 100);
    return () => clearTimeout(id);
  }, [getIdTokenClaims]);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  const sendMessage = useCallback(async (text: string) => {
    if (!text.trim() || isLoading) return;

    const userMsg: Message = { id: Date.now().toString(), role: "user", content: text.trim() };
    setMessages((p) => [...p, userMsg]);
    setInput("");
    if (textareaRef.current) textareaRef.current.style.height = "auto";
    setIsLoading(true);

    const controller = new AbortController();
    abortControllerRef.current = controller;
    const aid = (Date.now() + 1).toString();
    setMessages((p) => [...p, { id: aid, role: "assistant", content: "" }]);

    try {
      let freshToken: string | null = null;
      try { freshToken = await getAccessTokenSilently(); } catch { /* fallback */ }

      if (!freshToken)
        throw new Error("Auth token not available. Please sign out and sign in again.", { cause: "CAUGHT_ERROR: AUTH ERROR" });

      let resolvedTenant: string | null = null;
      try {
        const payload = JSON.parse(atob(freshToken.split(".")[1].replace(/-/g, "+").replace(/_/g, "/")));
        const meta = payload["fluidai_app_metadata"] as Record<string, unknown> | undefined;
        const tv = payload["tenant_value"] ?? meta?.["tenant"];
        resolvedTenant = Array.isArray(tv) ? (tv[0] ?? null) : typeof tv === "string" ? tv : null;
      } catch { resolvedTenant = tenant; }

      if (!resolvedTenant)
        throw new Error("Tenant not available. Please sign out and sign in again.", { cause: "CAUGHT_ERROR: AUTH ERROR" });

      const answer = await runAgent(
        text.trim(), freshToken, resolvedTenant,
        (status) => setMessages((p) => p.map((m) => m.id === aid ? { ...m, content: status } : m)),
        controller.signal
      );
      setMessages((p) => p.map((m) => m.id === aid ? { ...m, content: answer } : m));
    } catch (err) {
      const e = err as Error;
      setMessages((p) => p.map((m) =>
        m.id === aid ? { ...m, content: e.name === "AbortError" ? "Stopped." : e.message || "Something went wrong." } : m
      ));
    } finally {
      abortControllerRef.current = null;
      setIsLoading(false);
    }
  }, [isLoading, getAccessTokenSilently, tenant]);

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); sendMessage(input); }
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

  /* -- Header bg & border based on dark/light -- */
  const headerCls = dark
    ? "border-white/[0.06] bg-surface-900/90 backdrop-blur-xl"
    : "border-gray-100/80 bg-white/85 backdrop-blur-xl shadow-sm";

  const rootBg = dark ? "bg-surface-950" : "bg-[#f5f5f7]";

  return (
    <div className={cn("flex h-full w-full flex-col overflow-hidden font-sans transition-colors duration-300", rootBg, dark ? "text-gray-100" : "text-gray-900")}>

      {/* ---------- Header ---------- */}
      <header className={cn("relative flex flex-shrink-0 items-center justify-between border-b px-3.5 py-2 transition-colors duration-300", headerCls)}>
        {/* Subtle top shimmer line */}
        {dark && (
          <div
            className="pointer-events-none absolute inset-x-0 top-0 h-px"
            style={{ background: "linear-gradient(90deg, transparent, rgba(99,102,241,0.4), transparent)" }}
          />
        )}

        {/* Brand */}
        <div className="flex items-center gap-2">
          <div className={cn(
            "relative flex h-6 w-6 items-center justify-center rounded-lg bg-gradient-to-br from-brand-500 to-brand-700 shadow-md shadow-brand-500/30 transition-all duration-300",
            isLoading && "animate-[pulse-glow_1.5s_ease-in-out_infinite]"
          )}>
            <StarIcon className="text-white" />
            {/* live indicator dot */}
            <span className={cn(
              "absolute -right-0.5 -top-0.5 h-1.5 w-1.5 rounded-full ring-1",
              dark ? "ring-surface-900" : "ring-white",
              isLoading ? "bg-amber-400" : "bg-emerald-400"
            )} />
          </div>
          <span className={cn("text-[11px] font-semibold tracking-wide", dark ? "text-gray-200" : "text-gray-700")}>
            Fluid AI
          </span>
          <span className={cn(
            "rounded-full px-1.5 py-[2px] text-[9px] font-bold uppercase tracking-widest",
            dark ? "bg-brand-500/15 text-brand-400 ring-1 ring-brand-500/25" : "bg-brand-50 text-brand-600 ring-1 ring-brand-100"
          )}>
            Beta
          </span>
        </div>

        {/* Controls */}
        <div className="flex items-center gap-0.5">
          {/* New chat */}
          <button
            onClick={() => setMessages([])}
            title="New chat"
            className={cn(
              "focus-brand flex h-7 w-7 items-center justify-center rounded-lg transition-all duration-200",
              dark ? "text-gray-500 hover:bg-white/[0.07] hover:text-gray-200" : "text-gray-400 hover:bg-gray-100 hover:text-gray-700"
            )}
          >
            <svg width="13" height="13" viewBox="0 0 16 16" fill="none">
              <circle cx="8" cy="8" r="6.3" stroke="currentColor" strokeWidth="1.4" />
              <path d="M8 5.5v5M5.5 8h5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
            </svg>
          </button>

          {/* Theme */}
          <button
            onClick={() => setDark((d) => !d)}
            title={dark ? "Light mode" : "Dark mode"}
            className={cn(
              "focus-brand flex h-7 w-7 items-center justify-center rounded-lg transition-all duration-200",
              dark ? "text-gray-500 hover:bg-white/[0.07] hover:text-yellow-300" : "text-gray-400 hover:bg-gray-100 hover:text-brand-500"
            )}
          >
            {dark
              ? <svg width="13" height="13" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="3" stroke="currentColor" strokeWidth="1.4" /><path d="M8 1v2M8 13v2M1 8h2M13 8h2M3.05 3.05l1.42 1.42M11.53 11.53l1.42 1.42M3.05 12.95l1.42-1.42M11.53 4.47l1.42-1.42" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" /></svg>
              : <svg width="13" height="13" viewBox="0 0 16 16" fill="none"><path d="M13.5 10.5A6 6 0 015.5 2.5a6 6 0 000 11 6 6 0 008-3z" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" /></svg>
            }
          </button>

          {/* Avatar + menu */}
          <div ref={userMenuRef} className="relative ml-0.5">
            <button
              onClick={() => setShowMenu((v) => !v)}
              title={user?.email}
              className={cn(
                "focus-brand flex h-[26px] w-[26px] overflow-hidden rounded-full ring-2 transition-all duration-200",
                showMenu
                  ? "ring-brand-500 shadow-md shadow-brand-500/30"
                  : dark ? "ring-white/20 hover:ring-brand-400" : "ring-gray-200 hover:ring-brand-400"
              )}
            >
              {user?.picture
                ? <img src={user.picture} alt={user.name ?? "User"} className="h-full w-full object-cover" onError={(e) => { (e.target as HTMLImageElement).style.display = "none"; }} />
                : <div className="flex h-full w-full items-center justify-center bg-gradient-to-br from-amber-400 to-orange-500 text-[10px] font-bold text-white">{avatarLetter}</div>
              }
            </button>

            {showMenu && (
              <div
                className={cn("absolute right-0 top-9 z-50 min-w-[188px] rounded-2xl p-1.5 shadow-2xl ring-1", dark ? "glass-dark ring-white/[0.08]" : "glass ring-black/[0.06]")}
                style={{ animation: "fade-up 0.18s cubic-bezier(0.16,1,0.3,1) forwards" }}
              >
                <div className="px-3 py-2">
                  <p className={cn("truncate text-[12px] font-semibold", dark ? "text-white" : "text-gray-900")}>{user?.name ?? "User"}</p>
                  <p className="truncate text-[11px] text-gray-500">{user?.email}</p>
                </div>
                <div className={cn("my-1 h-px", dark ? "bg-white/[0.06]" : "bg-gray-100")} />
                <button
                  onClick={() => logout({ logoutParams: { returnTo: window.location.origin } })}
                  className={cn(
                    "flex w-full items-center gap-2 rounded-xl px-3 py-2 text-[12px] font-medium transition-all duration-150",
                    dark ? "text-gray-400 hover:bg-red-500/10 hover:text-red-400" : "text-gray-500 hover:bg-red-50 hover:text-red-600"
                  )}
                >
                  <svg width="12" height="12" viewBox="0 0 16 16" fill="none">
                    <path d="M10 3h3a1 1 0 011 1v8a1 1 0 01-1 1h-3M7 11l3-3-3-3M10 8H2" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                  </svg>
                  Sign out
                </button>
              </div>
            )}
          </div>
        </div>
      </header>

      {/* ---------- Messages ---------- */}
      <div className={cn(
        "scrollbar-thin relative flex flex-1 flex-col gap-3 overflow-y-auto px-3.5 py-4 min-h-0",
        dark ? "bg-surface-950" : "bg-[#f5f5f7]"
      )}>
        {messages.length === 0 ? (
          /* -- Empty / welcome state -- */
          <div className="flex flex-1 flex-col items-center justify-center gap-5 px-1 text-center">

            {/* Animated logo cluster */}
            <div className="relative" style={{ width: 76, height: 76 }}>
              {/* ambient glow */}
              <div className={cn(
                "absolute inset-0 rounded-3xl blur-2xl",
                dark ? "bg-brand-500/30" : "bg-brand-400/18"
              )} />
              {/* orbit rings */}
              <div className="orbit-ring" style={{ width: 68, height: 68, top: "50%", left: "50%", borderColor: dark ? "rgba(99,102,241,0.2)" : "rgba(99,102,241,0.15)" }}>
                <div className="orbit-dot" style={{ animationDuration: "10s" }} />
              </div>
              <div className="orbit-ring" style={{ width: 48, height: 48, top: "50%", left: "50%", borderColor: dark ? "rgba(167,139,250,0.18)" : "rgba(167,139,250,0.12)" }}>
                <div className="orbit-dot-2" style={{ "--orbit-r": "24px", animationDuration: "14s" } as React.CSSProperties} />
              </div>
              {/* icon */}
              <div className="float absolute left-1/2 top-1/2 z-10 flex h-11 w-11 -translate-x-1/2 -translate-y-1/2 items-center justify-center rounded-[14px] bg-gradient-to-br from-brand-500 to-brand-700 shadow-xl shadow-brand-500/35">
                <svg width="20" height="20" viewBox="0 0 16 16" fill="none">
                  <path d="M8 1l1.5 4.5L14 7l-4.5 1.5L8 13l-1.5-4.5L2 7l4.5-1.5z" fill="white" />
                </svg>
              </div>
            </div>

            <div className="flex flex-col gap-1">
              <p className={cn("text-[15px] font-semibold leading-snug", dark ? "text-white" : "text-gray-800")}>
                How can I help?
              </p>
              <p className={cn("text-[11.5px] leading-relaxed", dark ? "text-gray-500" : "text-gray-400")}>
                Ask anything about your spreadsheet
              </p>
            </div>

            {/* Quick action cards */}
            <div className="flex w-full flex-col gap-2">
              {QUICK_ACTIONS.map(({ label, icon, desc }, i) => (
                <button
                  key={label}
                  onClick={() => sendMessage(label)}
                  className={cn(
                    "action-card focus-brand group relative flex items-center gap-3 rounded-xl px-3.5 py-2.5 text-left ring-1 transition-all duration-250",
                    "opacity-0",
                    dark
                      ? "bg-white/[0.04] ring-white/[0.07] hover:bg-white/[0.07] hover:ring-brand-500/25 hover:shadow-lg hover:shadow-brand-500/10"
                      : "bg-white ring-gray-200/80 hover:bg-brand-50/60 hover:ring-brand-200 hover:shadow-md hover:shadow-brand-500/[0.07]"
                  )}
                  style={{ animation: `fade-up 0.4s ${0.05 * i + 0.05}s cubic-bezier(0.16,1,0.3,1) forwards` }}
                >
                  <span className="flex h-8 w-8 flex-shrink-0 items-center justify-center rounded-lg bg-gradient-to-br from-brand-500/10 to-violet-500/10 text-[16px] transition-transform duration-200 group-hover:scale-110">
                    {icon}
                  </span>
                  <div className="min-w-0 flex-1">
                    <p className={cn("text-[12.5px] font-medium leading-none", dark ? "text-gray-200" : "text-gray-700")}>
                      {label}
                    </p>
                    <p className={cn("mt-0.5 text-[10.5px] leading-snug", dark ? "text-gray-600" : "text-gray-400")}>
                      {desc}
                    </p>
                  </div>
                  <svg
                    className={cn(
                      "h-3.5 w-3.5 flex-shrink-0 -translate-x-1 opacity-0 transition-all duration-200 group-hover:translate-x-0 group-hover:opacity-100",
                      dark ? "text-brand-400" : "text-brand-500"
                    )}
                    viewBox="0 0 16 16" fill="none"
                  >
                    <path d="M3 8h10M9 4l4 4-4 4" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                  </svg>
                </button>
              ))}
            </div>
          </div>
        ) : (
          <>
            {messages.map((msg) =>
              msg.content ? <MessageBubble key={msg.id} msg={msg} dark={dark} /> : null
            )}
            {isLoading && <TypingBubble dark={dark} />}
            <div ref={messagesEndRef} />
          </>
        )}
      </div>

      {/* ---------- Input area ---------- */}
      <div className={cn("flex-shrink-0 px-3 pb-3 pt-2 transition-colors duration-300", dark ? "bg-surface-950" : "bg-[#f5f5f7]")}>
        <div
          className={cn(
            "relative overflow-hidden rounded-2xl transition-all duration-200",
            dark ? "bg-surface-800 input-focus-ring-dark" : "bg-white input-focus-ring",
          )}
        >
          {/* Animated top border glow when focused */}
          {inputFocused && (
            <div
              className="pointer-events-none absolute inset-x-0 top-0 h-px"
              style={{ background: "linear-gradient(90deg, transparent, rgba(99,102,241,0.7), rgba(167,139,250,0.7), transparent)" }}
            />
          )}

          {/* Attachment chip */}
          {attachedFile && (
            <div className={cn(
              "mx-3 mt-2.5 flex items-center gap-1.5 rounded-lg px-2.5 py-1.5 text-[11px] ring-1",
              dark ? "bg-white/[0.06] text-gray-400 ring-white/[0.08]" : "bg-gray-50 text-gray-500 ring-gray-200"
            )}>
              <svg width="10" height="10" viewBox="0 0 16 16" fill="none">
                <path d="M4 2h6l4 4v8a1 1 0 01-1 1H4a1 1 0 01-1-1V3a1 1 0 011-1z" stroke="currentColor" strokeWidth="1.4" />
                <path d="M10 2v4h4" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
              </svg>
              <span className="min-w-0 flex-1 truncate">{attachedFile.name}</span>
              <button
                onClick={() => setAttachedFile(null)}
                className="ml-1 flex-shrink-0 rounded p-0.5 opacity-50 transition-opacity hover:opacity-100"
                aria-label="Remove"
              >
                <svg width="9" height="9" viewBox="0 0 10 10" fill="none">
                  <path d="M1.5 1.5l7 7M8.5 1.5l-7 7" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
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
            onFocus={() => setInputFocused(true)}
            onBlur={() => setInputFocused(false)}
            placeholder="Ask anything about your spreadsheet "
            rows={1}
            disabled={isLoading}
            className={cn(
              "block w-full resize-none bg-transparent px-3.5 pt-3 pb-1 text-[13px] leading-relaxed outline-none transition-colors duration-200 disabled:cursor-not-allowed",
              dark ? "text-gray-100 placeholder:text-gray-600" : "text-gray-800 placeholder:text-gray-400"
            )}
            style={{ maxHeight: "100px", overflowY: "auto" }}
          />

          {/* Footer */}
          <div className="flex items-center justify-between px-2.5 pb-2.5 pt-1">
            <input
              ref={fileInputRef}
              type="file"
              accept="image/*,video/*,.pdf,.xlsx,.xls,.csv"
              style={{ display: "none" }}
              onChange={(e) => { setAttachedFile(e.target.files?.[0] ?? null); e.target.value = ""; }}
            />
            <button
              onClick={() => fileInputRef.current?.click()}
              title="Attach file"
              className={cn(
                "focus-brand flex h-7 w-7 items-center justify-center rounded-lg transition-all duration-200",
                dark ? "text-gray-600 hover:bg-white/[0.07] hover:text-gray-300" : "text-gray-400 hover:bg-gray-100 hover:text-gray-600"
              )}
            >
              <svg width="14" height="14" viewBox="0 0 16 16" fill="none">
                <path d="M2 11.5V13a1 1 0 001 1h10a1 1 0 001-1v-1.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
                <path d="M8 2v8M5 5l3-3 3 3" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </button>

            {isLoading ? (
              <button
                onClick={() => abortControllerRef.current?.abort()}
                aria-label="Stop"
                className="flex h-7 w-7 items-center justify-center rounded-lg bg-gradient-to-br from-red-500 to-rose-600 text-white shadow-md shadow-red-500/25 transition-all duration-200 hover:-translate-y-0.5 hover:shadow-lg"
              >
                <svg width="9" height="9" viewBox="0 0 16 16" fill="none">
                  <rect x="3" y="3" width="10" height="10" rx="2" fill="currentColor" />
                </svg>
              </button>
            ) : (
              <button
                onClick={() => sendMessage(input)}
                disabled={!canSend}
                aria-label="Send"
                className={cn(
                  "flex h-7 w-7 items-center justify-center rounded-lg transition-all duration-200",
                  canSend
                    ? "bg-gradient-to-br from-brand-500 to-brand-700 text-white shadow-md shadow-brand-500/30 hover:-translate-y-0.5 hover:shadow-lg hover:shadow-brand-500/40"
                    : dark ? "bg-white/[0.07] text-gray-600 cursor-not-allowed" : "bg-gray-100 text-gray-300 cursor-not-allowed"
                )}
              >
                <svg width="13" height="13" viewBox="0 0 16 16" fill="none">
                  <path d="M8 13V3M3 8l5-5 5 5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                </svg>
              </button>
            )}
          </div>
        </div>

        <p className={cn("mt-1.5 text-center text-[9.5px]", dark ? "text-gray-700" : "text-gray-300")}>
          <kbd className={cn("rounded px-1 py-px font-mono ring-1", dark ? "ring-white/10 text-gray-500" : "ring-gray-200 text-gray-400")}>?</kbd>{" "}
          to send &middot;{" "}
          <kbd className={cn("rounded px-1 py-px font-mono ring-1", dark ? "ring-white/10 text-gray-500" : "ring-gray-200 text-gray-400")}>??</kbd>{" "}
          new line
        </p>
      </div>
    </div>
  );
}
