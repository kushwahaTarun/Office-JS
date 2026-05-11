import { useAuth0 } from "@auth0/auth0-react";

const LoginButton = () => {
  const { loginWithPopup } = useAuth0();
  return (
    <button className="login-btn" onClick={() => loginWithPopup()}>
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
        <rect x="2" y="7" width="9" height="8" rx="1.2" stroke="currentColor" strokeWidth="1.4"/>
        <path d="M5 7V5a3 3 0 016 0v2" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
      </svg>
      Sign in
    </button>
  );
};

export default LoginButton;
