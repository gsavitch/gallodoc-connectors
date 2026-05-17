declare const OfficeRuntime: any;
declare const Office: any;

export interface ConnectorSettings {
  baseUrl: string;
  authType: "password" | "token";
  token: string | null;
  tokenType: string | null;
  username: string | null;
  connected: boolean;
  autoSyncEnabled: boolean;
  autoSyncIntervalMinutes: number;
}

const DEFAULT_SETTINGS: ConnectorSettings = {
  baseUrl: "",
  authType: "password",
  token: null,
  tokenType: "Token",
  username: null,
  connected: false,
  autoSyncEnabled: false,
  autoSyncIntervalMinutes: 5
};

const STORAGE_KEY = "halobridge_connector_settings";

/**
 * Persists settings using OfficeRuntime.storage (shared runtime) 
 * or partitioned localStorage (fallback).
 */
export async function saveConnectorSettings(settings: ConnectorSettings): Promise<void> {
  const data = JSON.stringify(settings);
  
  if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
    await OfficeRuntime.storage.setItem(STORAGE_KEY, data);
  } else {
    // Fallback to partitioned localStorage
    const key = Office.context.partitionKey ? `${STORAGE_KEY}_${Office.context.partitionKey}` : STORAGE_KEY;
    localStorage.setItem(key, data);
  }
}

export async function getConnectorSettings(): Promise<ConnectorSettings> {
  let data: string | null = null;

  if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
    data = await OfficeRuntime.storage.getItem(STORAGE_KEY);
  } else {
    const key = Office.context.partitionKey ? `${STORAGE_KEY}_${Office.context.partitionKey}` : STORAGE_KEY;
    data = localStorage.getItem(key);
  }

  if (!data) return { ...DEFAULT_SETTINGS };
  try {
    return JSON.parse(data);
  } catch (e) {
    return { ...DEFAULT_SETTINGS };
  }
}

export async function clearConnectorSettings(): Promise<void> {
  if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
    await OfficeRuntime.storage.removeItem(STORAGE_KEY);
  } else {
    const key = Office.context.partitionKey ? `${STORAGE_KEY}_${Office.context.partitionKey}` : STORAGE_KEY;
    localStorage.removeItem(key);
  }
}
