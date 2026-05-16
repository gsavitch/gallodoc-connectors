import { ConnectorMode, MODE_CONFIG } from "../lib/modeConfig";
import { generateLocalGalloDoc } from "../lib/gallodocLocal";
import { HaloBridgeClient } from "../lib/halobridgeClient";
import { getConnectorSettings, saveConnectorSettings, clearConnectorSettings, ConnectorSettings } from "../lib/storage";

/* global Office, Word */

const hbClient = new HaloBridgeClient();
let currentSettings: ConnectorSettings;

let currentMode = ConnectorMode.Local;
let lastResult: any = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initializeTaskPane();
  }
});

async function initializeTaskPane() {
  const modeSelector = document.getElementById("modeSelector") as HTMLSelectElement;
  const btnAction = document.getElementById("btnAction") as HTMLButtonElement;
  const btnView = document.getElementById("btnView") as HTMLButtonElement;
  const btnCloseModal = document.getElementById("btnCloseModal") as HTMLButtonElement;
  const btnCopy = document.getElementById("btnCopy") as HTMLButtonElement;
  
  const btnConnect = document.getElementById("btnConnect") as HTMLButtonElement;
  const btnTest = document.getElementById("btnTest") as HTMLButtonElement;
  const btnDisconnect = document.getElementById("btnDisconnect") as HTMLButtonElement;
  const authModeSelect = document.getElementById("authMode") as HTMLSelectElement;

  // Load persistent settings
  currentSettings = await getConnectorSettings();
  
  // Set default if empty and in development/AI Studio
  if (!currentSettings.baseUrl && typeof window !== "undefined") {
    currentSettings.baseUrl = window.location.origin;
  }

  // Hydrate UI from settings
  hydrateUI();

  modeSelector.addEventListener("change", (e) => {
    currentMode = (e.target as HTMLSelectElement).value as ConnectorMode;
    updateUIForMode();
  });

  authModeSelect.addEventListener("change", (e) => {
    const mode = (e.target as HTMLSelectElement).value;
    document.getElementById("passwordFields")?.classList.toggle("hidden", mode !== "password");
    document.getElementById("tokenFields")?.classList.toggle("hidden", mode !== "token");
  });

  btnAction.addEventListener("click", handleAction);
  btnView.addEventListener("click", showLastResult);
  btnCloseModal.addEventListener("click", () => document.getElementById("resultOverlay")?.classList.add("hidden"));
  btnCopy.addEventListener("click", copyToClipboard);
  
  btnConnect.addEventListener("click", handleConnect);
  btnTest.addEventListener("click", handleTestConnection);
  btnDisconnect.addEventListener("click", handleDisconnect);

  updateUIForMode();
  updateUIForConnection();
}

function hydrateUI() {
  const baseUrlInput = document.getElementById("baseUrl") as HTMLInputElement;
  const authModeSelect = document.getElementById("authMode") as HTMLSelectElement;
  const apiTokenInput = document.getElementById("apiToken") as HTMLInputElement;
  const usernameInput = document.getElementById("username") as HTMLInputElement;

  baseUrlInput.value = currentSettings.baseUrl || "";
  authModeSelect.value = currentSettings.authType;
  apiTokenInput.value = currentSettings.token || "";
  usernameInput.value = currentSettings.username || "";

  document.getElementById("passwordFields")?.classList.toggle("hidden", currentSettings.authType !== "password");
  document.getElementById("tokenFields")?.classList.toggle("hidden", currentSettings.authType !== "token");

  if (currentSettings.connected && currentSettings.baseUrl) {
    hbClient.setConfiguration(currentSettings.baseUrl, currentSettings.token);
  }
}

async function handleTestConnection() {
  const baseUrl = (document.getElementById("baseUrl") as HTMLInputElement).value;
  if (!baseUrl) {
    updateStatus("ERROR", "Base URL is required for testing.");
    return;
  }
  
  updateStatus("PROCESSING", "Testing HaloBridge connectivity...");
  const result = await hbClient.testConnection(baseUrl);
  
  if (result.status === "ok") {
    updateStatus("SUCCESS", "HaloBridge Reachable");
  } else {
    updateStatus("ERROR", result.message);
  }
}

async function handleConnect() {
  const baseUrl = (document.getElementById("baseUrl") as HTMLInputElement).value;
  const authType = (document.getElementById("authMode") as HTMLSelectElement).value as "password" | "token";

  if (!baseUrl) {
    updateStatus("ERROR", "Base URL is required to connect.");
    return;
  }

  updateStatus("PROCESSING", "Connecting to HaloBridge...");

  try {
    if (authType === "password") {
      const username = (document.getElementById("username") as HTMLInputElement).value;
      const password = (document.getElementById("password") as HTMLInputElement).value;

      if (!username || !password) throw new Error("Credentials required");

      const res = await hbClient.login(baseUrl, username, password);
      currentSettings = {
        baseUrl,
        authType: "password",
        token: res.token,
        username: res.user.display_name,
        connected: true
      };
      // Clear password field immediately after use
      (document.getElementById("password") as HTMLInputElement).value = "";
    } else {
      const token = (document.getElementById("apiToken") as HTMLInputElement).value;
      if (!token) throw new Error("API Token required");

      hbClient.setConfiguration(baseUrl, token);
      currentSettings = {
        baseUrl,
        authType: "token",
        token: token,
        username: "API Client",
        connected: true
      };
    }

    await saveConnectorSettings(currentSettings);
    updateStatus("SUCCESS", `Connected to ${baseUrl}`);
    updateUIForConnection();
  } catch (error: any) {
    updateStatus("ERROR", error.message);
  }
}

function updateUIForConnection() {
  const loginForm = document.getElementById("loginForm") as HTMLDivElement;
  const connectedState = document.getElementById("connectedState") as HTMLDivElement;
  const connectedUser = document.getElementById("connectedUser") as HTMLParagraphElement;
  const connectedUrl = document.getElementById("connectedUrl") as HTMLParagraphElement;

  if (currentSettings.connected) {
    loginForm.classList.add("hidden");
    connectedState.classList.remove("hidden");
    connectedUser.innerText = currentSettings.username || "Connected";
    connectedUrl.innerText = currentSettings.baseUrl;
    updateUIForMode(); 
  } else {
    loginForm.classList.remove("hidden");
    connectedState.classList.add("hidden");
    updateUIForMode();
  }
}

async function handleDisconnect() {
  await clearConnectorSettings();
  const oldUrl = currentSettings.baseUrl;
  currentSettings = {
    baseUrl: oldUrl, // Preserve for convenience
    authType: "password",
    token: null,
    username: null,
    connected: false
  };
  hbClient.setConfiguration("", null);
  updateUIForConnection();
  updateStatus("IDLE", "Disconnected from HaloBridge.");
}

function updateUIForMode() {
  const config = MODE_CONFIG[currentMode];
  const btnAction = document.getElementById("btnAction") as HTMLButtonElement;
  const modeDescription = document.getElementById("modeDescription") as HTMLParagraphElement;
  
  if (btnAction) btnAction.innerText = config.buttonText;
  if (modeDescription) modeDescription.innerText = config.description;
  
  if (config.requiresLogin) {
    if (currentSettings.connected) {
      updateStatus("READY", `Syncing to ${currentSettings.baseUrl}`);
    } else {
      updateStatus("DISCONNECTED", "Configure connection to use cloud sync.");
    }
  } else {
    updateStatus("IDLE", "Local mode active. No data will be uploaded.");
  }
}

async function handleAction() {
  const config = MODE_CONFIG[currentMode];
  if (config.requiresLogin && !currentSettings.connected) {
    updateStatus("ERROR", "Connection required for cloud sync.");
    return;
  }

  updateStatus("PROCESSING", "Reading document...");
  
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const text = body.getText();
      const ooxml = body.getOoxml();
      
      await context.sync();

      const docName = Office.context.document.url ? Office.context.document.url.split('/').pop() || "Document.docx" : "Unsaved Document";
      
      if (currentMode === ConnectorMode.Local) {
        lastResult = await generateLocalGalloDoc(text.value, !!ooxml.value, docName);
        updateStatus("SUCCESS", "Local GalloDoc generated.");
        showLastResult();
      } else {
        updateStatus("UPLOADING", "Syncing to HaloBridge...");
        lastResult = await hbClient.saveWordDocument({
          mode: currentMode,
          document_name: docName,
          document_text: text.value,
          ooxml: ooxml.value,
          metadata: {
            client_platform: Office.context.diagnostics.platform,
            office_version: Office.context.diagnostics.version
          }
        });
        updateStatus("SYNCED", `Cloud Version ${lastResult.version_number} saved.`);
        const btnAction = document.getElementById("btnAction") as HTMLButtonElement;
        if (btnAction) btnAction.innerText = "Re-sync";
        document.getElementById("btnView")?.classList.remove("hidden");
      }
    });
  } catch (error: any) {
    if (error.message === "AUTH_EXPIRED") {
      updateStatus("ERROR", "Session expired.");
      handleDisconnect();
    } else {
      updateStatus("ERROR", error.message || "Operation failed.");
    }
  }
}

function updateStatus(status: string, message: string) {
  const connectionStatus = document.getElementById("connectionStatus") as HTMLSpanElement;
  const statusMessage = document.getElementById("statusMessage") as HTMLParagraphElement;
  
  if (connectionStatus) connectionStatus.innerText = status;
  if (statusMessage) statusMessage.innerText = message;

  const colors: Record<string, string> = {
    IDLE: "text-gray-600",
    PROCESSING: "text-blue-600",
    UPLOADING: "text-yellow-600",
    SUCCESS: "text-green-600",
    SYNCED: "text-green-600",
    ERROR: "text-red-600"
  };
  if (connectionStatus) connectionStatus.className = `text-[10px] font-mono ${colors[status] || "text-gray-600"}`;
}

function showLastResult() {
  if (!lastResult) return;
  const overlay = document.getElementById("resultOverlay") as HTMLDivElement;
  const output = document.getElementById("resultOutput") as HTMLPreElement;
  
  output.innerText = JSON.stringify(lastResult, null, 2);
  overlay.classList.remove("hidden");
}

function copyToClipboard() {
    const text = document.getElementById("resultOutput")?.innerText;
    if (text) {
        navigator.clipboard.writeText(text).then(() => {
            const btn = document.getElementById("btnCopy") as HTMLButtonElement;
            const originalText = btn.innerText;
            btn.innerText = "COPIED!";
            setTimeout(() => btn.innerText = originalText, 2000);
        });
    }
}
