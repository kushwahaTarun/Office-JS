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
        <div className="action-card">
          <p className="action-text">
            Get started by signing in to your account
          </p>
          <LoginButton />
        </div>
      )}
    </>
  );
}

export default App;
