import { useAuth0 } from "@auth0/auth0-react";

const LoginButton = () => {
  const { loginWithPopup } = useAuth0();
  return (
    <button
      className="btn-shimmer focus-brand group relative flex w-full items-center justify-center gap-2 rounded-xl bg-gradient-to-r from-brand-500 to-brand-700 px-5 py-2.5 text-sm font-semibold text-white shadow-lg shadow-brand-500/30 transition-all duration-200 hover:-translate-y-0.5 hover:shadow-xl hover:shadow-brand-500/40 active:translate-y-0 active:shadow-md"
      onClick={() => loginWithPopup()}
    >
      {/* Lock icon */}
      <svg
        className="transition-transform duration-200 group-hover:scale-110"
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
    </button>
  );
};

export default LoginButton;
