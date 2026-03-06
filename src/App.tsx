import { useAuth0 } from "@auth0/auth0-react";
import LoginButton from "./components/LoginButton";
import ChatComponent from "./components/ChatComponent";

function App() {
  const { isAuthenticated, isLoading, error } = useAuth0();

  if (isLoading) {
    return (
      <div className="app-container">
        <div className="loading-state">
          <div className="loading-text">Loading...</div>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="app-container">
        <div className="error-state">
          <div className="error-title">Oops!</div>
          <div className="error-message">Something went wrong</div>
          <div className="error-sub-message">{error.message}</div>
        </div>
      </div>
    );
  }

  return (
    <>
      {isAuthenticated ? (
        <ChatComponent />
      ) : (
        <div className="login-screen">
          <div className="login-card">
            <div className="login-icon">
              <svg width="28" height="28" viewBox="0 0 16 16" fill="none">
                <path d="M8 1l1.5 4.5L14 7l-4.5 1.5L8 13l-1.5-4.5L2 7l4.5-1.5z" fill="#6366f1"/>
              </svg>
            </div>
            <h2 className="login-title">Fluid GPT for Excel</h2>
            <p className="login-subtitle">Sign in to start taking AI-powered actions in your spreadsheet</p>
            <LoginButton />
          </div>
        </div>
      )}
    </>
  );
}

export default App;
