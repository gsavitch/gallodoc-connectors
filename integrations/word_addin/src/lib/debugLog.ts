
/* global localStorage */

export const DEBUG_KEY = "halobridge_debug_api";

export interface ApiEventDetails {
  eventName: string;
  timestamp: string;
  baseUrl?: string;
  endpoint?: string;
  method?: string;
  authType?: string;
  tokenPresent: boolean;
  tokenType?: string;
  payloadKeys?: string[];
  payloadSizeApprox?: number;
  saveAction?: string;
  mvpDocumentIdPresent: boolean;
  previousDocumentIdPresent: boolean;
  status?: number;
  errorCode?: string;
  errorMessage?: string;
  responseContentType?: string;
}

let lastEvent: ApiEventDetails | null = null;
let lastError: any = null;

export function isDebugEnabled(): boolean {
  try {
    return localStorage.getItem(DEBUG_KEY) === "true" || process.env.NODE_ENV === "development";
  } catch {
    return false;
  }
}

export function headerToString(value: unknown): string | undefined {
  if (value == null) return undefined;
  if (Array.isArray(value)) return value.join(", ");
  return String(value);
}

export function redactToken(token: string | null): { tokenPresent: boolean; tokenPrefix?: string } {
  if (!token) return { tokenPresent: false };
  return {
    tokenPresent: true,
    tokenPrefix: `${token.substring(0, 6)}...`
  };
}

export function debugApiEvent(details: Omit<ApiEventDetails, 'timestamp'>) {
  const event: ApiEventDetails = {
    ...details,
    timestamp: new Date().toISOString()
  };
  
  lastEvent = event;
  
  if (isDebugEnabled()) {
    console.group(`[HaloBridge API] ${event.eventName}`);
    console.log("Details:", event);
    if (event.errorMessage) console.error("Error Message:", event.errorMessage);
    console.groupEnd();
  }
}

export function setLastError(error: any) {
  lastError = error;
}

export function getDebugSummary(currentSettings: any, hbClient: any) {
  const tokenInfo = redactToken(currentSettings.token);
  
  return {
    baseUrl: currentSettings.baseUrl,
    healthEndpoint: "/api/health/",
    loginEndpoint: "/api/word/auth/login/",
    saveEndpoint: "/api/word/documents/save/",
    connected: currentSettings.connected,
    tokenPresent: tokenInfo.tokenPresent,
    tokenPrefix: tokenInfo.tokenPrefix,
    tokenType: currentSettings.tokenType,
    lastApiEvent: lastEvent,
    lastApiError: lastError ? {
      name: lastError.name,
      message: lastError.message,
      code: lastError.code,
      status: lastError.response?.status,
      contentType: headerToString(lastError.response?.headers?.['content-type'])
    } : null
  };
}
