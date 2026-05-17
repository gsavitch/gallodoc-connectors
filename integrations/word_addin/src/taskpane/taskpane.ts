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
let autoSyncTimer: number | null = null;

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
    updateUIForAutoSync();
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
  const btnToggleAutoSyncSection = document.getElementById("btnToggleAutoSyncSection") as HTMLButtonElement;

  btnSaveAs.addEventListener("click", () => handleAction("save_as"));
  btnAction.addEventListener("click", () => handleAction("save"));

  btnToggleAutoSyncSection.addEventListener("click", () => {
    const content = document.getElementById("autoSyncContent");
    content?.classList.toggle("hidden");
  });

  const toggleAutoSync = document.getElementById("toggleAutoSync") as HTMLButtonElement;
  const autoSyncInterval = document.getElementById("autoSyncInterval") as HTMLSelectElement;

  toggleAutoSync.addEventListener("click", handleToggleAutoSync);
  autoSyncInterval.addEventListener("change", handleAutoSyncIntervalChange);

  updateUIForMode();
  updateUIForConnection();
  updateUIForAutoSync();
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
  const badgeSync = document.getElementById("badgeSync") as HTMLElement;
  const btnOpenHB = document.getElementById("btnOpenHB") as HTMLButtonElement;
  const docTitleInput = document.getElementById("docTitle") as HTMLInputElement;

  if (currentManifest) {
    docStatusPanel.classList.remove("hidden");
    stDocId.innerText = currentManifest.mvp_document_id;
    stVersion.innerText = `#${currentManifest.latest_version_number}`;
    stReview.innerText = currentManifest.review_status;
    stLastSync.innerText = `Last sync: ${new Date(currentManifest.last_synced_at).toLocaleString()}`;
    btnSaveAs.classList.remove("hidden");
    badgeSync.classList.remove("hidden");

    if (currentManifest.document_name && docTitleInput && !docTitleInput.value) {
      docTitleInput.value = currentManifest.document_name;
    }

    // Handle Open in HaloBridge button
    if (currentManifest.word_control_url) {
      btnOpenHB.classList.remove("hidden");
      btnOpenHB.onclick = (e) => {
        e.preventDefault();
        window.open(currentManifest!.word_control_url!, "_blank");
      };
    } else {
      btnOpenHB.classList.add("hidden");
    }
  } else {
    docStatusPanel.classList.add("hidden");
    btnSaveAs.classList.add("hidden");
    badgeSync.classList.add("hidden");
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

  const autoSyncInterval = document.getElementById("autoSyncInterval") as HTMLSelectElement;
  if (autoSyncInterval) {
    autoSyncInterval.value = String(currentSettings.autoSyncIntervalMinutes ?? 5);
  }

  document.getElementById("passwordFields")?.classList.toggle("hidden", currentSettings.authType !== "password");
  document.getElementById("tokenFields")?.classList.toggle("hidden", currentSettings.authType !== "token");

  if (currentSettings.connected && currentSettings.baseUrl) {
    hbClient.setConfiguration(currentSettings.baseUrl, currentSettings.token, currentSettings.tokenType || "Token");
  }

  // Set initial mode from selector
  const modeSelector = document.getElementById("modeSelector") as HTMLSelectElement;
  currentMode = modeSelector.value as ConnectorMode;
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
  const passwordInput = document.getElementById("password") as HTMLInputElement;

  if (!baseUrl) {
    updateStatus("ERROR", "Base URL required.");
    return;
  }

  updateStatus("PROCESSING", "Connecting...");

  try {
    if (authType === "password") {
      const username = (document.getElementById("username") as HTMLInputElement).value;
      const password = passwordInput.value;

      if (!username || !password) throw new Error("Credentials required");

      const res = await hbClient.login(baseUrl, username, password);
      
      // Update settings with actual values from login
      currentSettings = {
        baseUrl,
        authType: "password",
        token: res.token,
        tokenType: res.tokenType || "Token",
        username: res.user.display_name || res.user.username || username,
        connected: !!res.token,
        autoSyncEnabled: currentSettings?.autoSyncEnabled ?? false,
        autoSyncIntervalMinutes: currentSettings?.autoSyncIntervalMinutes ?? 5
      };
      
      // Clear password field immediately
      passwordInput.value = "";
      
      console.log(`[Auth] Connected via Password to ${baseUrl}. Token: ${!!res.token}`);
    } else {
      const token = (document.getElementById("apiToken") as HTMLInputElement).value;
      if (!token) throw new Error("API Token required");

      const defaultType = "Token";
      hbClient.setConfiguration(baseUrl, token, defaultType);
      
      // Verification test for token
      const test = await hbClient.testConnection(baseUrl);
      if (test.status === "unreachable") throw new Error("Server unreachable with this token.");

      currentSettings = {
        baseUrl,
        authType: "token",
        token: token,
        tokenType: defaultType,
        username: "API Token User",
        connected: true,
        autoSyncEnabled: currentSettings?.autoSyncEnabled ?? false,
        autoSyncIntervalMinutes: currentSettings?.autoSyncIntervalMinutes ?? 5
      };

      console.log(`[Auth] Connected via API Token to ${baseUrl}`);
    }

    await saveConnectorSettings(currentSettings);
    
    // Synchronize client
    if (currentSettings.token) {
        hbClient.setConfiguration(currentSettings.baseUrl, currentSettings.token, currentSettings.tokenType || "Token");
    }

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
  const badgeConn = document.getElementById("badgeConn") as HTMLElement;

  const isActuallyConnected = currentSettings.connected && !!currentSettings.token && !!currentSettings.baseUrl;

  if (isActuallyConnected) {
    loginForm.classList.add("hidden");
    connectedState.classList.remove("hidden");
    connectedUser.innerText = currentSettings.username || "Authenticated";
    connectedUrl.innerText = currentSettings.baseUrl;
    
    if (badgeConn) {
      badgeConn.innerText = "Connected";
      badgeConn.className = "px-2 py-0.5 text-[9px] font-bold uppercase bg-green-100 text-green-700 rounded-full";
    }
  } else {
    loginForm.classList.remove("hidden");
    connectedState.classList.add("hidden");
    
    if (badgeConn) {
      badgeConn.innerText = "Not Connected";
      badgeConn.className = "px-2 py-0.5 text-[9px] font-bold uppercase bg-gray-100 text-gray-500 rounded-full";
    }
  }
  
  updateUIForMode();
}

async function handleDisconnect() {
  const oldUrl = currentSettings.baseUrl;
  currentSettings = {
    baseUrl: oldUrl, // Preserve for convenience
    authType: currentSettings.authType,
    token: null,
    tokenType: "Token",
    username: null,
    connected: false,
    autoSyncEnabled: currentSettings?.autoSyncEnabled ?? false,
    autoSyncIntervalMinutes: currentSettings?.autoSyncIntervalMinutes ?? 5
  };
  await saveConnectorSettings(currentSettings);
  hbClient.setConfiguration("", null, "Token");
  stopAutoSyncTimer();
  updateUIForConnection();
  updateUIForAutoSync();
  updateStatus("IDLE", "Disconnected from HaloBridge.");
}

function updateUIForAutoSync() {
  const toggleAutoSync = document.getElementById("toggleAutoSync") as HTMLButtonElement;
  const toggleKnob = document.getElementById("toggleKnob") as HTMLDivElement;
  const badgeAutoSync = document.getElementById("badgeAutoSync") as HTMLElement;
  const autoSyncStatus = document.getElementById("autoSyncStatus") as HTMLElement;

  if (currentSettings.autoSyncEnabled) {
    toggleAutoSync.className = "w-10 h-5 bg-green-500 rounded-full relative transition-colors focus:outline-none";
    toggleKnob.className = "absolute right-1 top-1 w-3 h-3 bg-white rounded-full transition-transform";
    badgeAutoSync.innerText = "ACTIVE";
    badgeAutoSync.className = "px-2 py-0.5 text-[9px] font-bold uppercase bg-green-100 text-green-700 rounded-full";
    
    if (!currentManifest) {
      autoSyncStatus.innerText = "Save once to HaloBridge to enable.";
    } else {
      autoSyncStatus.innerText = "Syncing periodically if changes detected.";
    }
  } else {
    toggleAutoSync.className = "w-10 h-5 bg-gray-300 rounded-full relative transition-colors focus:outline-none";
    toggleKnob.className = "absolute left-1 top-1 w-3 h-3 bg-white rounded-full transition-transform";
    badgeAutoSync.innerText = "OFF";
    badgeAutoSync.className = "px-2 py-0.5 text-[9px] font-bold uppercase bg-gray-100 text-gray-500 rounded-full";
    autoSyncStatus.innerText = "Auto-sync disabled.";
  }

  // Manage timer state
  if (currentSettings.autoSyncEnabled && currentSettings.connected && currentMode !== ConnectorMode.Local && currentManifest) {
    startAutoSyncTimer();
  } else {
    stopAutoSyncTimer();
  }
}

async function handleToggleAutoSync() {
  currentSettings.autoSyncEnabled = !currentSettings.autoSyncEnabled;
  await saveConnectorSettings(currentSettings);
  updateUIForAutoSync();
}

async function handleAutoSyncIntervalChange(e: Event) {
  const val = parseInt((e.target as HTMLSelectElement).value);
  currentSettings.autoSyncIntervalMinutes = val;
  await saveConnectorSettings(currentSettings);
  
  if (currentSettings.autoSyncEnabled) {
    startAutoSyncTimer(); // Restart with new interval
  }
}

function startAutoSyncTimer() {
  stopAutoSyncTimer();
  
  const ms = currentSettings.autoSyncIntervalMinutes * 60 * 1000;
  console.log(`[AutoSync] Starting timer for ${currentSettings.autoSyncIntervalMinutes} minutes.`);
  
  autoSyncTimer = window.setInterval(performAutoSync, ms);
}

function stopAutoSyncTimer() {
  if (autoSyncTimer !== null) {
    console.log(`[AutoSync] Stopping timer.`);
    window.clearInterval(autoSyncTimer);
    autoSyncTimer = null;
  }
}

async function performAutoSync() {
  if (!currentSettings.autoSyncEnabled || !currentSettings.connected || currentMode === ConnectorMode.Local || !currentManifest) {
    stopAutoSyncTimer();
    return;
  }

  const autoSyncStatus = document.getElementById("autoSyncStatus") as HTMLElement;
  autoSyncStatus.innerText = "Checking for changes...";
  console.log("[AutoSync] Triggered periodic check.");

  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      const ooxmlResult = body.getOoxml();
      await context.sync();

      const documentText = body.text || "";
      const ooxmlContent = ooxmlResult.value || "";
      const sourceHash = SHA256(ooxmlContent || documentText).toString();

      if (sourceHash === currentManifest?.last_source_hash) {
        console.log("[AutoSync] No changes detected. Skipping.");
        autoSyncStatus.innerText = `Last check: ${new Date().toLocaleTimeString()} (No changes)`;
        return;
      }

      console.log("[AutoSync] Changes detected. Syncing version...");
      autoSyncStatus.innerText = "Syncing new version...";
      
      // Perform the save
      await handleAction("save");
      
      autoSyncStatus.innerText = `Last auto-sync: ${new Date().toLocaleTimeString()}`;
    });
  } catch (err) {
    console.error("[AutoSync] Error during auto-sync:", err);
    autoSyncStatus.innerText = "Auto-sync failed. See console.";
  }
}

function updateUIForMode() {
  const config = MODE_CONFIG[currentMode];
  const btnAction = document.getElementById("btnAction") as HTMLButtonElement;
  const modeDescription = document.getElementById("modeDescription") as HTMLParagraphElement;
  const badgeMode = document.getElementById("badgeMode") as HTMLElement;
  
  if (badgeMode) {
    badgeMode.innerText = currentMode === ConnectorMode.Local ? "Local" : "Cloud";
    badgeMode.classList.add("hidden"); // We are removing badgeMode from HTML in redesign
  }

  if (btnAction) {
    if (currentMode === ConnectorMode.Local) {
        btnAction.innerText = "Create Local GalloDoc";
    } else {
        btnAction.innerText = currentManifest ? "Save Version" : "Save to HaloBridge";
    }
  }

  if (modeDescription) modeDescription.innerText = config.description;
  
  if (config.requiresLogin) {
    const isConn = currentSettings.connected && !!currentSettings.token;
    if (isConn) {
      updateStatus("READY", `Target: ${currentSettings.baseUrl}`);
    } else {
      updateStatus("AUTH", "Connect to sync to cloud.");
    }
  } else {
    updateStatus("IDLE", "Local only — no upload.");
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
    // If not linked yet, Save As behaves like a normal initial Save
    console.log("[Diagnostic] Save As requested on unlinked document. Defaulting to initial save path.");
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

      // Safe Document Naming Logic
      const getSafeDocumentName = () => {
        const docTitleInput = document.getElementById("docTitle") as HTMLInputElement;
        if (docTitleInput && docTitleInput.value.trim()) {
          return docTitleInput.value.trim();
        }

        const rawUrl = Office.context.document.url;
        if (!rawUrl) return `Word GalloDoc - ${new Date().toLocaleDateString()}`;

        // Strip path and get filename
        const fileName = rawUrl.split(/[/\\]/).pop() || "Document.docx";
        const isTemp = fileName.toLowerCase().includes("word add-in") || 
                       rawUrl.toLowerCase().includes("appdata") || 
                       rawUrl.toLowerCase().includes("temp");

        if (isTemp || fileName === "Document.docx") {
          return `Word GalloDoc - ${new Date().toLocaleDateString()}`;
        }

        return fileName;
      };

      const docName = getSafeDocumentName();
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
            connector_manifest: currentManifest,
            original_path: Office.context.document.url // Preserve original path in metadata
          }
        });

        // Update local manifest
        const newManifest = buildManifestFromSaveResponse(lastResult, sourceHash, currentManifest, docName);
        const success = await writeGalloDocManifest(newManifest);
        
        if (success) {
          updateStatus("SYNCED", `Version ${lastResult.version_number} saved.`);
          await refreshManifest();
          updateUIForAutoSync(); // Refresh auto-sync state in case it was waiting for first save
        } else {
          updateStatus("WARNING", "Cloud saved, but local manifest failed.");
        }

        const btnAction = document.getElementById("btnAction") as HTMLButtonElement;
        if (btnAction) btnAction.innerText = "Re-sync";
        document.getElementById("btnView")?.classList.remove("hidden");
      }
    });
  } catch (error: any) {
    console.error("[Diagnostic] handleAction caught error:", error);
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
    IDLE: "bg-gray-100 text-gray-600",
    READY: "bg-green-100 text-green-700",
    PROCESSING: "bg-blue-100 text-blue-700",
    UPLOADING: "bg-yellow-100 text-yellow-700",
    SUCCESS: "bg-green-100 text-green-700",
    SYNCED: "bg-green-100 text-green-700",
    ERROR: "bg-red-100 text-red-700",
    AUTH: "bg-orange-100 text-orange-700"
  };
  if (connectionStatus) {
    connectionStatus.className = `text-[9px] font-mono px-1 ${colors[status] || "bg-gray-200 text-gray-700"}`;
  }
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
