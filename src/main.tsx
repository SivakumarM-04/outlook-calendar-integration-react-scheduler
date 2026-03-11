import ReactDOM from "react-dom/client";
import App from "./App";
import { MsalProvider } from "@azure/msal-react";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig } from "./authConfig";
import "./index.css";
import { enableRipple } from "@syncfusion/ej2-base";

const pca = new PublicClientApplication(msalConfig);

enableRipple(true);

const root = ReactDOM.createRoot(document.getElementById("root")!);

// Ensure MSAL is initialized and redirect handled BEFORE rendering
(async () => {
  try {
    if ("initialize" in pca) {
      // @ts-ignore - initialize exists in modern msal-browser
      await pca.initialize();
    }
    const resp = await pca.handleRedirectPromise();
    if (resp) {
      console.debug("MSAL redirect response processed", resp.account?.username || "");
    }
  } catch (e) {
    console.warn("MSAL initialization/handleRedirectPromise error", e);
  } finally {
    root.render(
      <MsalProvider instance={pca}>
        <App />
      </MsalProvider>
    );
  }
})();