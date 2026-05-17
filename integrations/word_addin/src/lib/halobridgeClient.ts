import axios from 'axios';
import { ConnectorMode } from './modeConfig';

export interface ConnectionInfo {
  baseUrl: string;
  user?: {
    username: string;
    display_name: string;
  };
  tenant?: {
    name: string;
  };
}

export class HaloBridgeClient {
  private baseUrl: string | null = null;
  private token: string | null = null;
  private tokenType: string = "Token";

  constructor() {}

  setConfiguration(baseUrl: string, token: string | null, tokenType: string = "Token") {
    this.baseUrl = baseUrl.replace(/\/$/, "");
    this.token = token;
    this.tokenType = tokenType || "Token";
  }

  async testConnection(baseUrl: string): Promise<{ status: "ok" | "unauthorized" | "unreachable", message: string }> {
    const cleanUrl = baseUrl.replace(/\/$/, "");
    const endpoint = `${cleanUrl}/api/health/`;
    
    console.debug(`[Diagnostic] Testing connection to: ${endpoint}`);
    
    // 1. Try with Axios first (GET, no custom headers to avoid CORS preflight issues)
    try {
      const response = await axios.get(endpoint, { 
        timeout: 8000,
        headers: {
          'Accept': 'application/json'
        },
        // Avoid sending any authorization header for simple health test
        transformRequest: [(data, headers) => {
          delete headers['Authorization'];
          return data;
        }]
      });

      if (response.status === 200) {
        return { status: "ok", message: "Server reachable" };
      }
      return { status: "unreachable", message: `Server status: ${response.status}` };
    } catch (error: any) {
      console.debug(`[Diagnostic] Axios test failed for ${endpoint}. Code: ${error.code}, Message: ${error.message}`);
      
      // 2. Fallback to simple fetch (CORS mode) if axios fails with generic Network Error
      try {
        console.debug(`[Diagnostic] Attempting fetch fallback for ${endpoint}`);
        const fetchResponse = await fetch(endpoint, { 
          method: "GET", 
          mode: "cors",
          headers: { 'Accept': 'application/json' }
        });
        
        if (fetchResponse.ok) {
          return { status: "ok", message: "Server reachable (via fetch)" };
        }
        
        const status = fetchResponse.status;
        if (status === 401 || status === 403) return { status: "unauthorized", message: "Unauthorized/Forbidden" };
        return { status: "unreachable", message: `Server unreachable (Status ${status})` };
      } catch (fetchError: any) {
        console.debug(`[Diagnostic] Fetch fallback also failed.`, fetchError);
        
        // Final combined error message
        const isNetworkErr = !error.response;
        const msg = isNetworkErr ? "Network/CORS error: request blocked" : (error.response?.data?.message || error.message);
        return { 
          status: "unreachable", 
          message: `${msg}${error.code ? ` (${error.code})` : ""}` 
        };
      }
    }
  }

  async login(baseUrl: string, username: string, password: string) {
    const cleanUrl = baseUrl.replace(/\/$/, "");
    const endpoint = `${cleanUrl}/api/word/auth/login/`;
    
    console.debug(`[Diagnostic] Login attempt to: ${endpoint}`);
    
    try {
      const response = await axios.post(endpoint, {
        username,
        password
      }, {
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        // Explicitly don't send auth header for login
        transformRequest: [(data, headers) => {
          delete headers['Authorization'];
          return JSON.stringify(data);
        }]
      });

      const { access_token, token_type, user, tenant } = response.data;
      
      this.token = access_token;
      this.tokenType = token_type || "Token";
      this.baseUrl = cleanUrl;

      return {
        token: access_token,
        tokenType: this.tokenType,
        user: { 
          username: user.email || user.username, 
          display_name: user.display_name || user.username 
        },
        tenant: tenant ? { name: tenant.name } : undefined
      };
    } catch (error: any) {
      this.handleDetailedError('Login', error);
      throw error;
    }
  }

  async saveWordDocument(payload: {
    mode: ConnectorMode;
    document_name: string;
    document_text: string;
    ooxml?: string;
    save_action?: "save" | "save_as";
    document_id?: string;
    source_document_id?: string;
    metadata?: any;
  }) {
    if (!this.baseUrl || !this.token) {
      throw new Error("Add-in is not connected to a HaloBridge instance.");
    }

    const endpoint = `${this.baseUrl}/api/word/gallodoc/save/`;
    console.debug(`[Diagnostic] Save request to: ${endpoint}`);

    try {
      const response = await axios.post(endpoint, payload, {
        headers: {
          'Authorization': `${this.tokenType} ${this.token}`,
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        timeout: 45000 
      });
      return response.data;
    } catch (error: any) {
      this.handleDetailedError('Save', error);
      throw error;
    }
  }

  private handleDetailedError(context: string, error: any) {
    const status = error.response?.status;
    const data = error.response?.data;
    const msg = error.message;
    const code = error.code;
    
    console.error(`[Diagnostic] ${context} Error:`, {
      code,
      message: msg,
      status,
      responseData: data,
      requestMade: !!error.request,
      responseReceived: !!error.response
    });

    if (!error.response) {
      if (error.code === 'ECONNABORTED') {
        error.message = "Request timed out. Server is taking too long.";
      } else {
        error.message = `Network/CORS error: Request blocked before response. Status code: ${code || 'Unknown'}`;
      }
    } else {
      const serverMsg = data?.message || data?.error || data?.detail;
      error.message = `${context} failed${status ? ` (${status})` : ""}: ${serverMsg || msg}`;
    }
    
    if (status === 401) {
      // Specialized message for expired tokens in Save mode
      if (context === 'Save') {
        error.message = "AUTH_EXPIRED";
      }
    }
  }
}
