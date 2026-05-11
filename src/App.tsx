import { useAuth0 } from "@auth0/auth0-react";
import LoginButton from "./components/LoginButton";
import ChatComponent from "./components/ChatComponent";

function App() {
  const { isAuthenticated, isLoading, error } = useAuth0();

  if (isLoading) {
    return (
      <div className="app-container">
        <div className="loading-state">
          <div className="brand-mark" aria-hidden="true">
            <span />
          </div>
          <div className="loading-text">Preparing workspace</div>
          <div className="loading-subtext">Connecting Fluid GPT to Excel</div>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="app-container">
        <div className="error-state">
          <div className="brand-mark error" aria-hidden="true">
            <span />
          </div>
          <div className="error-title">Connection interrupted</div>
          <div className="error-message">Something went wrong while starting Fluid GPT.</div>
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
          <div className="login-ambient login-ambient-one" />
          <div className="login-ambient login-ambient-two" />
          <div className="login-card">
            <div className="login-icon">
              <span />
            </div>
            <div className="login-kicker">Excel AI Copilot</div>
            <h1 className="login-title">Fluid GPT</h1>
            <p className="login-subtitle">
              Turn workbook context into formulas, summaries, and spreadsheet actions with a focused AI task pane.
            </p>
            <div className="login-feature-row" aria-label="Capabilities">
              <span>Analyze</span>
              <span>Transform</span>
              <span>Automate</span>
            </div>
            <LoginButton />
          </div>
        </div>
      )}
    </>
  );
}

export default App;
