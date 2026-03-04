import { useAuth0 } from "@auth0/auth0-react";
import LoginButton from "./components/LoginButton";
import LogoutButton from "./components/LogoutButton";
import ExcelComponent from "./components/ExcelComponent";

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
        <div className="App">
          <header className="App-header">
            <h1>Fluid AI GPT - Excel Add-in</h1>
            <p>Office.js Excel Integration</p>
          </header>

          <main>
            <ExcelComponent />
          </main>

          <LogoutButton />
        </div>
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
