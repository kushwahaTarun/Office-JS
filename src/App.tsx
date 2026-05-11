import { useAuth0 } from "@auth0/auth0-react";
import LoginButton from "./components/LoginButton";
import ChatComponent from "./components/ChatComponent";

/* Tiny twinkle particle */
function Particle({
  x,
  y,
  size,
  delay,
}: {
  x: string;
  y: string;
  size: number;
  delay: string;
}) {
  return (
    <div
      className="particle"
      style={{
        left: x,
        top: y,
        width: size,
        height: size,
        animationDelay: delay,
      }}
    />
  );
}

const PARTICLES = [
  { x: "12%", y: "18%", size: 3, delay: "0s" },
  { x: "28%", y: "72%", size: 2, delay: "0.8s" },
  { x: "55%", y: "12%", size: 3.5, delay: "1.4s" },
  { x: "80%", y: "35%", size: 2, delay: "0.3s" },
  { x: "68%", y: "78%", size: 4, delay: "1.9s" },
  { x: "6%", y: "55%", size: 2.5, delay: "0.6s" },
  { x: "90%", y: "62%", size: 3, delay: "1.1s" },
  { x: "42%", y: "88%", size: 2, delay: "2.2s" },
];

function App() {
  const { isAuthenticated, isLoading, error } = useAuth0();

  /* ── Loading state ── */
  if (isLoading) {
    return (
      <div className="aurora-bg relative flex h-full w-full items-center justify-center overflow-hidden">
        <div className="dot-grid pointer-events-none absolute inset-0 opacity-[0.04]" />
        <div className="flex flex-col items-center gap-5">
          {/* Orbit rings */}
          <div className="relative flex items-center justify-center">
            <div className="orbit-ring" style={{ width: 80, height: 80 }}>
              <div className="orbit-dot" />
            </div>
            <div className="orbit-ring" style={{ width: 110, height: 110 }}>
              <div
                className="orbit-dot-2"
                style={{ animationDuration: "18s" }}
              />
            </div>
            {/* Center icon */}
            <div className="relative z-10 flex h-12 w-12 items-center justify-center rounded-2xl bg-gradient-to-br from-brand-500 to-brand-700 shadow-2xl">
              <div className="glow-halo h-10 w-10 bg-brand-500/50" />
              <svg
                className="pulse-glow relative z-10"
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
          </div>
          <div className="flex items-center gap-2">
            <span className="typing-dot bg-brand-400" />
            <span className="typing-dot bg-brand-400" />
            <span className="typing-dot bg-brand-400" />
          </div>
        </div>
      </div>
    );
  }

  /* ── Error state ── */
  if (error) {
    return (
      <div className="aurora-bg relative flex h-full w-full items-center justify-center overflow-hidden p-6">
        <div className="dot-grid pointer-events-none absolute inset-0 opacity-[0.04]" />
        <div
          className="glass-card relative z-10 flex flex-col items-center gap-3 rounded-2xl p-8 text-center"
          style={{
            animation: "fade-up 0.4s cubic-bezier(0.16,1,0.3,1) forwards",
            maxWidth: 280,
          }}
        >
          <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-red-500/15 ring-1 ring-red-500/30">
            <svg width="20" height="20" viewBox="0 0 16 16" fill="none">
              <circle
                cx="8"
                cy="8"
                r="6.3"
                stroke="#f87171"
                strokeWidth="1.4"
              />
              <path
                d="M8 5v3.5M8 10.5v.5"
                stroke="#f87171"
                strokeWidth="1.6"
                strokeLinecap="round"
              />
            </svg>
          </div>
          <p className="text-sm font-semibold text-white">
            Something went wrong
          </p>
          <p className="text-xs text-gray-400">{error.message}</p>
        </div>
      </div>
    );
  }

  if (isAuthenticated) return <ChatComponent />;

  /* ── Login screen ── */
  return (
    <div className="aurora-bg relative flex h-full w-full items-center justify-center overflow-hidden">
      {/* Dot grid */}
      <div className="dot-grid pointer-events-none absolute inset-0 opacity-[0.045]" />

      {/* Scan line */}
      <div className="scan-line" />

      {/* Ambient blobs */}
      <div
        className="blob h-72 w-72 bg-brand-600"
        style={{ top: "-10%", left: "-12%" }}
      />
      <div
        className="blob blob-delay-1 h-56 w-56 bg-violet-600"
        style={{ bottom: "0%", right: "-8%" }}
      />
      <div
        className="blob blob-delay-2 h-40 w-40 bg-indigo-400"
        style={{ top: "50%", left: "10%" }}
      />

      {/* Star particles */}
      {PARTICLES.map((p, i) => (
        <Particle key={i} {...p} />
      ))}

      {/* ── Card ── */}
      <div
        className="grad-border relative z-10"
        style={{
          borderRadius: 28,
          animation: "fade-up 0.55s cubic-bezier(0.16,1,0.3,1) forwards",
        }}
      >
        <div className="glass-card relative flex w-[290px] flex-col items-center gap-6 overflow-hidden rounded-[27px] p-7">
          {/* inner noise texture */}
          <div className="noise pointer-events-none absolute inset-0 rounded-[27px]" />

          {/* ── Logo with orbit rings ── */}
          <div
            className="relative flex items-center justify-center"
            style={{ width: 88, height: 88, animationDelay: "0.1s" }}
          >
            {/* Outer orbit ring */}
            <div
              className="orbit-ring"
              style={{
                width: 80,
                height: 80,
                borderColor: "rgba(99,102,241,0.20)",
              }}
            >
              <div className="orbit-dot" style={{ animationDuration: "10s" }} />
            </div>
            {/* Inner orbit ring */}
            <div
              className="orbit-ring"
              style={{
                width: 58,
                height: 58,
                borderColor: "rgba(167,139,250,0.18)",
              }}
            >
              <div
                className="orbit-dot-2"
                style={
                  {
                    animationName: "orbit-rev",
                    "--orbit-r": "29px",
                    animationDuration: "14s",
                  } as React.CSSProperties
                }
              />
            </div>
            {/* Glow halo */}
            <div
              className="glow-halo h-11 w-11 bg-brand-500/40"
              style={{ animationDuration: "2.5s" }}
            />
            {/* Logo box */}
            <div className="float relative z-10 flex h-11 w-11 items-center justify-center rounded-[14px] bg-gradient-to-br from-brand-500 via-brand-600 to-brand-700 shadow-2xl shadow-brand-500/50">
              <svg width="22" height="22" viewBox="0 0 16 16" fill="none">
                <path
                  d="M8 1l1.5 4.5L14 7l-4.5 1.5L8 13l-1.5-4.5L2 7l4.5-1.5z"
                  fill="white"
                />
              </svg>
            </div>
          </div>

          {/* ── Copy ── */}
          <div
            className="flex flex-col items-center gap-1.5 text-center"
            style={{
              animation: "fade-up 0.55s 0.12s cubic-bezier(0.16,1,0.3,1) both",
            }}
          >
            <h2 className="text-shimmer text-[18px] font-bold leading-tight tracking-tight">
              Fluid GPT for Excel
            </h2>
            <p className="text-[12px] leading-relaxed text-gray-400">
              AI-powered actions, right inside
              <br />
              your spreadsheet
            </p>
          </div>

          {/* ── Feature pills ── */}
          <div
            className="flex flex-wrap justify-center gap-1.5"
            style={{
              animation: "fade-up 0.55s 0.22s cubic-bezier(0.16,1,0.3,1) both",
            }}
          >
            {["Formulas", "Charts", "Data Analysis", "Automation"].map((f) => (
              <span
                key={f}
                className="rounded-full bg-brand-500/12 px-2.5 py-0.5 text-[10px] font-medium text-brand-300 ring-1 ring-brand-500/20"
              >
                {f}
              </span>
            ))}
          </div>

          {/* ── Divider ── */}
          <div className="h-px w-full bg-gradient-to-r from-transparent via-white/10 to-transparent" />

          {/* ── Login button ── */}
          <div
            className="w-full"
            style={{
              animation: "fade-up 0.55s 0.30s cubic-bezier(0.16,1,0.3,1) both",
            }}
          >
            <LoginButton />
          </div>

          {/* ── Footer ── */}
          <p
            className="text-center text-[10px] text-gray-600"
            style={{ animation: "fade-in 0.5s 0.45s ease both" }}
          >
            :lock: Secured by Auth0 &middot; End-to-end encrypted
          </p>
        </div>
      </div>
    </div>
  );
}

export default App;
