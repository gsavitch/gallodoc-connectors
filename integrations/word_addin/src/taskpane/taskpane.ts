import { ConnectorMode, MODE_CONFIG } from "../lib/modeConfig";
import { generateLocalGalloDoc } from "../lib/gallodocLocal";
import { HaloBridgeClient } from "../lib/halobridgeClient";

/* global Office, Word */

// Configuration - In production, this would be injected via environment or settings
const hbClient = new HaloBridgeClient();

let currentMode = ConnectorMode.Local;
let lastResult: any = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initializeTaskPane();
  }
});

function initializeTaskPane() {
  const modeSelector = document.getElementById("modeSelector") as HTMLSelectElement;
  const btnAction = document.getElementById("btnAction") as HTMLButtonElement;
  const btnView = document.getElementById("btnView") as HTMLButtonElement;
  const btnCloseModal = document.getElementById("btnCloseModal") as HTMLButtonElement;
  const btnCopy = document.getElementById("btnCopy") as HTMLButtonElement;
  const modeDescription = document.getElementById("modeDescription") as HTMLParagraphElement;
  
  const btnConnect = document.getElementById("btnConnect") as HTMLButtonElement;
  const btnDisconnect = document.getElementById("btnDisconnect") as HTMLButtonElement;
  const baseUrlInput = document.getElementById("baseUrl") as HTMLInputElement;

  // Set default base URL for ease of use in AI Studio
  baseUrlInput.value = window.location.origin;

  modeSelector.addEventListener("change", (e) => {
    currentMode = (e.target as HTMLSelectElement).value as ConnectorMode;
    updateUIForMode();
  });

  btnAction.addEventListener("click", handleAction);
  btnView.addEventListener("click", showLastResult);
  btnCloseModal.addEventListener("click", () => document.getElementById("resultOverlay")?.classList.add("hidden"));
  btnCopy.addEventListener("click", copyToClipboard);
  
  btnConnect.addEventListener("click", handleConnect);
  btnDisconnect.addEventListener("click", () => {
    hbClient.logout();
    updateUIForConnection();
  });

  updateUIForMode();
  updateUIForConnection();
}

async function handleConnect() {
  const baseUrl = (document.getElementById("baseUrl") as HTMLInputElement).value;
  const username = (document.getElementById("username") as HTMLInputElement).value;
  const password = (document.getElementById("password") as HTMLInputElement).value;

  if (!baseUrl || !username || !password) {
    updateStatus("ERROR", "Please fill in all connection fields.");
    return;
  }

  updateStatus("CONNECTING", "Establishing connection...");

  try {
    const conn = await hbClient.login(baseUrl, username, password);
    updateStatus("CONNECTED", `Logged in as ${conn.user.display_name}`);
    updateUIForConnection();
  } catch (error: any) {
    updateStatus("ERROR", error.message);
  }
}

function updateUIForConnection() {
  const conn = hbClient.getConnection();
  const loginForm = document.getElementById("loginForm") as HTMLDivElement;
  const connectedState = document.getElementById("connectedState") as HTMLDivElement;
  const connectedUser = document.getElementById("connectedUser") as HTMLParagraphElement;
  const connectedTenant = document.getElementById("connectedTenant") as HTMLParagraphElement;

  if (conn) {
    loginForm.classList.add("hidden");
    connectedState.classList.remove("hidden");
    connectedUser.innerText = conn.user.display_name;
    connectedTenant.innerText = conn.tenant.name;
    updateUIForMode(); // Refresh mode warnings
  } else {
    loginForm.classList.remove("hidden");
    connectedState.classList.add("hidden");
    updateUIForMode();
  }
}

function updateUIForMode() {
  const config = MODE_CONFIG[currentMode];
  const btnAction = document.getElementById("btnAction") as HTMLButtonElement;
  const modeDescription = document.getElementById("modeDescription") as HTMLParagraphElement;
  
  btnAction.innerText = config.buttonText;
  modeDescription.innerText = config.description;
  
  if (config.requiresLogin) {
    if (hbClient.isLoggedIn()) {
      updateStatus("READY", `Connected to ${hbClient.getConnection()?.tenant.name}`);
    } else {
      updateStatus("DISCONNECTED", "Connect to HaloBridge before saving in this mode.");
    }
  } else {
    updateStatus("IDLE", "Local mode active. No data will be uploaded.");
  }
}

async function handleAction() {
  const config = MODE_CONFIG[currentMode];
  if (config.requiresLogin && !hbClient.isLoggedIn()) {
    updateStatus("ERROR", "Authorization required for this mode.");
    return;
  }

  updateStatus("PROCESSING", "Reading document content...");
  
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const text = body.getText();
      const ooxml = body.getOoxml();
      
      await context.sync();

      const docName = "Word_Document.docx"; // Fallback if title not found
      
      if (currentMode === ConnectorMode.Local) {
        lastResult = await generateLocalGalloDoc(text.value, !!ooxml.value, docName);
        updateStatus("SUCCESS", "Local GalloDoc JSON generated.");
        showLastResult();
      } else {
        updateStatus("UPLOADING", "Syncing with HaloBridge Cloud...");
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
        updateStatus("SYNCED", `Version ${lastResult.version_number} saved to HaloBridge.`);
        const btnAction = document.getElementById("btnAction") as HTMLButtonElement;
        btnAction.innerText = "Re-sync to HaloBridge";
        document.getElementById("btnView")?.classList.remove("hidden");
      }
    });
  } catch (error: any) {
    console.error(error);
    updateStatus("ERROR", error.message || "Failed to process document.");
  }
}

function updateStatus(status: string, message: string) {
  const connectionStatus = document.getElementById("connectionStatus") as HTMLSpanElement;
  const statusMessage = document.getElementById("statusMessage") as HTMLParagraphElement;
  
  connectionStatus.innerText = status;
  statusMessage.innerText = message;

  // Change status color
  const colors: Record<string, string> = {
    IDLE: "text-gray-600",
    PROCESSING: "text-blue-600",
    UPLOADING: "text-yellow-600",
    SUCCESS: "text-green-600",
    SYNCED: "text-green-600",
    ERROR: "text-red-600"
  };
  connectionStatus.className = `text-[10px] font-mono ${colors[status] || "text-gray-600"}`;
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
