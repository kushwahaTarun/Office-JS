import { useAuth0 } from "@auth0/auth0-react";
import LoginButton from "./components/LoginButton";
import ChatComponent from "./components/ChatComponent";

function App() {
  const { isAuthenticated, isLoading, error } = useAuth0();

  if (isLoading) {
    return (
      <div className="flex h-full w-full items-center justify-center bg-surface-50">
        <div className="flex flex-col items-center gap-4">
          {/* Animated logo */}
          <div className="relative flex h-12 w-12 items-center justify-center rounded-2xl bg-gradient-to-br from-brand-500 to-brand-700 shadow-lg shadow-brand-500/30">
            <svg
              className="pulse-glow"
              width="22"
              height="22"
              viewBox="0 0 16 16"
              fill="none"
            >
              <path
                d="M8 1l1.5 4.5L14 7l-4.5 1.5L8 13l-1.5-4.5L2 7l4.5-1.5z"
                fill="white"
              />
            </svg>
          </div>
          {/* Dots */}
          <div className="flex items-center gap-1.5">
            <span className="typing-dot bg-brand-400" />
            <span className="typing-dot bg-brand-400" />
            <span className="typing-dot bg-brand-400" />
          </div>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex h-full w-full items-center justify-center bg-surface-50 p-6">
        <div className="glass flex flex-col items-center gap-3 rounded-2xl p-8 text-center shadow-xl shadow-black/5">
          <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-red-100">
            <svg width="22" height="22" viewBox="0 0 16 16" fill="none">
              <circle cx="8" cy="8" r="6.3" stroke="#ef4444" strokeWidth="1.4" />
              <path d="M8 5v3.5M8 10.5v.5" stroke="#ef4444" strokeWidth="1.6" strokeLinecap="round" />
            </svg>
          </div>
          <p className="text-sm font-semibold text-gray-800">Something went wrong</p>
          <p className="text-xs text-gray-500">{error.message}</p>
        </div>
      </div>
    );
  }

  if (isAuthenticated) {
    return <ChatComponent />;
  }

  return (
    <div className="relative flex h-full w-full items-center justify-center overflow-hidden bg-gradient-to-br from-indigo-50 via-white to-violet-50">
      {/* Animated background blobs */}
      <div className="blob blob-delay-0 h-64 w-64 bg-brand-500" style={{ top: '-5%', left: '-10%' }} />
      <div className="blob blob-delay-1 h-48 w-48 bg-violet-400" style={{ top: '55%', right: '-8%' }} />
      <div className="blob blob-delay-2 h-40 w-40 bg-indigo-300" style={{ bottom: '5%', left: '20%' }} />

      {/* Grid dot pattern overlay */}
      <div
        className="pointer-events-none absolute inset-0 opacity-[0.035]"
        style={{
          backgroundImage:
            'radial-gradient(circle, #6366f1 1px, transparent 1px)',
          backgroundSize: '22px 22px',
        }}
      />

      {/* Login card */}
      <div
        className="glass relative z-10 flex w-full max-w-[300px] flex-col items-center gap-5 rounded-3xl p-8 shadow-2xl shadow-brand-500/10"
        style={{ animation: 'fade-up 0.5s ease-out forwards' }}
      >
        {/* Brand icon */}
        <div className="flex h-14 w-14 items-center justify-center rounded-2xl bg-gradient-to-br from-brand-500 to-brand-700 shadow-lg shadow-brand-500/40">
          <svg width="26" height="26" viewBox="0 0 16 16" fill="none">
            <path
              d="M8 1l1.5 4.5L14 7l-4.5 1.5L8 13l-1.5-4.5L2 7l4.5-1.5z"
              fill="white"
            />
          </svg>
        </div>

        {/* Title */}
        <div className="flex flex-col items-center gap-1.5 text-center">
          <h2 className="text-shimmer text-lg font-bold tracking-tight">
            Fluid GPT for Excel
          </h2>
          <p className="text-xs leading-relaxed text-gray-500">
            Sign in to start taking AI-powered<br />actions in your spreadsheet
          </p>
        </div>

        {/* Divider */}
        <div className="h-px w-full bg-gradient-to-r from-transparent via-gray-200 to-transparent" />

        {/* Login button */}
        <LoginButton />

        {/* Footer note */}
        <p className="text-center text-[10px] text-gray-400">
          Secured by Auth0 &middot; End-to-end encrypted
        </p>
      </div>
    </div>
  );
}

export default App;