import { useAuth0 } from "@auth0/auth0-react";

const LoginButton = () => {
  const { loginWithPopup } = useAuth0();
  return (
    <button
      className="btn-shimmer btn-glow focus-brand group relative flex w-full items-center justify-center gap-2.5 rounded-xl bg-gradient-to-r from-brand-500 via-brand-600 to-brand-700 px-5 py-3 text-[13px] font-semibold text-white shadow-lg shadow-brand-500/40 transition-all duration-300 hover:-translate-y-0.5 hover:shadow-xl hover:shadow-brand-500/50 active:translate-y-0 active:shadow-md"
      onClick={() => loginWithPopup()}
    >
      <svg
        className="transition-transform duration-300 group-hover:scale-110 group-hover:rotate-[-8deg]"
        width="15"
        height="15"
        viewBox="0 0 16 16"
        fill="none"
      >
        <rect
          x="2"
          y="7"
          width="9"
          height="8"
          rx="1.2"
          stroke="white"
          strokeWidth="1.4"
        />
        <path
          d="M5 7V5a3 3 0 016 0v2"
          stroke="white"
          strokeWidth="1.4"
          strokeLinecap="round"
        />
      </svg>
      Sign in to continue
      <svg
        className="ml-auto h-3.5 w-3.5 -translate-x-1 opacity-0 transition-all duration-300 group-hover:translate-x-0 group-hover:opacity-100"
        viewBox="0 0 16 16"
        fill="none"
      >
        <path
          d="M3 8h10M9 4l4 4-4 4"
          stroke="white"
          strokeWidth="1.5"
          strokeLinecap="round"
          strokeLinejoin="round"
        />
      </svg>
    </button>
  );
};

export default LoginButton;
