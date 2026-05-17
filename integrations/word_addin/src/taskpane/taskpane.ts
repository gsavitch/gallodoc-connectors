import { ConnectorMode, MODE_CONFIG } from "../lib/modeConfig";
import { generateLocalGalloDoc } from "../lib/gallodocLocal";
import { HaloBridgeClient } from "../lib/halobridgeClient";
import { getConnectorSettings, saveConnectorSettings, clearConnectorSettings, ConnectorSettings } from "../lib/storage";
import { GalloDocManifest, readGalloDocManifest, writeGalloDocManifest, buildManifestFromSaveResponse } from "../lib/wordManifest";
import SHA256 from "crypto-js/sha256";
import "./taskpane.css";

/* global SHA256 */
declare const OfficeRuntime: any;
declare const Office: any;

const hbClient = new HaloBridgeClient();
let currentSettings: ConnectorSettings;
let currentManifest: GalloDocManifest | null = null;

let currentMode = ConnectorMode.Local;
let lastResult: any = null;

Office.onReady((info: any) => {
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

  btnView.addEventListener("click", showLastResult);
  btnCloseModal.addEventListener("click", () => document.getElementById("resultOverlay")?.classList.add("hidden"));
  btnCopy.addEventListener("click", copyToClipboard);
  
  btnConnect.addEventListener("click", handleConnect);
  btnTest.addEventListener("click", handleTestConnection);
  btnDisconnect.addEventListener("click", handleDisconnect);

  const btnSaveAs = document.getElementById("btnSaveAs") as HTMLButtonElement;
  btnSaveAs.addEventListener("click", () => handleAction("save_as"));
  btnAction.addEventListener("click", () => handleAction("save"));

  updateUIForMode();
  updateUIForConnection();
  await refreshManifest();
}

async function refreshManifest() {
  currentManifest = await readGalloDocManifest();
  updateUIForManifest();
}

function updateUIForManifest() {
  const docStatusPanel = document.getElementById("docStatusPanel") as HTMLElement;
  const stDocId = document.getElementById("stDocId") as HTMLElement;
  const stVersion = document.getElementById("stVersion") as HTMLElement;
  const stReview = document.getElementById("stReview") as HTMLElement;
  const stLastSync = document.getElementById("stLastSync") as HTMLElement;
  const btnSaveAs = document.getElementById("btnSaveAs") as HTMLButtonElement;

  if (currentManifest) {
    docStatusPanel.classList.remove("hidden");
    stDocId.innerText = currentManifest.mvp_document_id;
    stVersion.innerText = `#${currentManifest.latest_version_number}`;
    stReview.innerText = currentManifest.review_status;
    stLastSync.innerText = new Date(currentManifest.last_synced_at).toLocaleString();
    btnSaveAs.classList.remove("hidden");
  } else {
    docStatusPanel.classList.add("hidden");
    btnSaveAs.classList.add("hidden");
  }
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
    hbClient.setConfiguration(currentSettings.baseUrl, currentSettings.token, currentSettings.tokenType || "Token");
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
        tokenType: res.tokenType || "Token",
        username: res.user.display_name || res.user.username || username,
        connected: true
      };
      
      console.log(`[Auth] Connected via Password to ${baseUrl}. Token present: ${!!res.token}, Type: ${res.tokenType}`);

      // Clear password field immediately after use
      (document.getElementById("password") as HTMLInputElement).value = "";
    } else {
      const token = (document.getElementById("apiToken") as HTMLInputElement).value;
      if (!token) throw new Error("API Token required");

      const defaultType = "Token";
      hbClient.setConfiguration(baseUrl, token, defaultType);
      currentSettings = {
        baseUrl,
        authType: "token",
        token: token,
        tokenType: defaultType,
        username: "api-token",
        connected: true
      };

      console.log(`[Auth] Connected via API Token to ${baseUrl}. Type: ${defaultType}`);
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
  const oldUrl = currentSettings.baseUrl;
  currentSettings = {
    baseUrl: oldUrl, // Preserve for convenience
    authType: currentSettings.authType,
    token: null,
    tokenType: "Token",
    username: null,
    connected: false
  };
  await saveConnectorSettings(currentSettings);
  hbClient.setConfiguration("", null, "Token");
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

async function handleAction(action: "save" | "save_as" = "save") {
  const config = MODE_CONFIG[currentMode];
  
  // Re-verify connection state from storage
  const settings = await getConnectorSettings();
  currentSettings = settings; // Update local state for consistency

  if (config.requiresLogin) {
    if (!settings.baseUrl) {
      updateStatus("ERROR", "Connection required: missing HaloBridge Base URL.");
      return;
    }
    if (!settings.token) {
      updateStatus("ERROR", "Connection required: no auth token found.");
      return;
    }
    if (!settings.connected) {
      updateStatus("ERROR", "Connection required: click Connect first.");
      return;
    }
    
    // Ensure client is synchronized with storage
    hbClient.setConfiguration(settings.baseUrl, settings.token, settings.tokenType || "Token");
  }

  if (action === "save_as" && !currentManifest) {
    updateStatus("ERROR", "Unlinked document. Save to HaloBridge first.");
    return;
  }

  updateStatus("PROCESSING", "Reading document...");
  
  try {
    await Word.run(async (context) => {
      const body = context.document.body;

      body.load("text");
      const ooxmlResult = body.getOoxml();

      await context.sync();

      const documentText = body.text || "";
      const ooxmlContent = ooxmlResult.value || "";

      const docName = Office.context.document.url ? Office.context.document.url.split('/').pop() || "Document.docx" : "Unsaved Document";
      const sourceHash = SHA256(ooxmlContent || documentText).toString();

      if (currentMode === ConnectorMode.Local) {
        lastResult = await generateLocalGalloDoc(documentText, !!ooxmlContent, docName);
        updateStatus("SUCCESS", "Local GalloDoc generated.");
        showLastResult();
      } else {
        updateStatus("UPLOADING", action === "save_as" ? "Creating new branch..." : "Syncing to HaloBridge...");
        
        lastResult = await hbClient.saveWordDocument({
          mode: currentMode,
          document_name: docName,
          document_text: documentText,
          ooxml: ooxmlContent,
          save_action: action,
          document_id: (action === "save") ? currentManifest?.mvp_document_id : undefined,
          source_document_id: (action === "save_as") ? currentManifest?.mvp_document_id : undefined,
          metadata: {
            office_version: Office.context.diagnostics.version,
            connector_manifest: currentManifest
          }
        });

        // Update local manifest
        const newManifest = buildManifestFromSaveResponse(lastResult, sourceHash, currentManifest);
        const success = await writeGalloDocManifest(newManifest);
        
        if (success) {
          updateStatus("SYNCED", `Version ${lastResult.version_number} saved.`);
          await refreshManifest();
        } else {
          updateStatus("WARNING", "Cloud saved, but local manifest failed.");
        }

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
