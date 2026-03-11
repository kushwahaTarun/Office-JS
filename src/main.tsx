import { createRoot } from "react-dom/client";
import "./index.css";
import App from "./App.tsx";
import { Auth0Provider } from "@auth0/auth0-react";

// Ensure Office.js is loaded before initializing React
if (typeof Office !== "undefined") {
  Office.onReady((info) => {
    console.log("Office.js ready. Host:", info.host, "Platform:", info.platform);

    // Wait a bit more to ensure Excel API is fully loaded
    setTimeout(() => {
      createRoot(document.getElementById("root")!).render(
        // <StrictMode>
          <Auth0Provider
            domain={import.meta.env.VITE_AUTH0_DOMAIN}
            clientId={import.meta.env.VITE_AUTH0_CLIENT_ID}
            authorizationParams={{ audience: import.meta.env.VITE_AUTH0_AUDIENCE }}
          >
            <App />
          </Auth0Provider>
        // </StrictMode>,
      );
    }, 200);
  }).catch((error) => {
    console.error("Office.onReady failed:", error);
    // Fallback: render anyway for development/testing
    createRoot(document.getElementById("root")!).render(
      <Auth0Provider
        domain={import.meta.env.VITE_AUTH0_DOMAIN}
        clientId={import.meta.env.VITE_AUTH0_CLIENT_ID}
        authorizationParams={{ audience: import.meta.env.VITE_AUTH0_AUDIENCE }}
      >
        <App />
      </Auth0Provider>
    );
  });
} else {
  // Office.js not loaded - probably development mode or browser
  console.warn("Office.js not detected. Running in standalone mode.");
  createRoot(document.getElementById("root")!).render(
    <Auth0Provider
      domain={import.meta.env.VITE_AUTH0_DOMAIN}
      clientId={import.meta.env.VITE_AUTH0_CLIENT_ID}
      authorizationParams={{ audience: import.meta.env.VITE_AUTH0_AUDIENCE }}
    >
      <App />
    </Auth0Provider>
  );
}
