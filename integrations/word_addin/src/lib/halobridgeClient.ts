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
    try {
      console.log(`[Diagnostic] Testing connection to: ${cleanUrl}/api/health/`);
      const response = await axios.get(`${cleanUrl}/api/health/`, { timeout: 10000 });
      if (response.status === 200) return { status: "ok", message: "Server Reachable" };
      return { status: "unreachable", message: `Server returned ${response.status}` };
    } catch (error: any) {
      if (error.response?.status === 401) return { status: "unauthorized", message: "Unauthorized" };
      const detail = error.response?.data?.message || error.response?.data?.error || error.message;
      console.error('[Diagnostic] Connection Test Failed:', error);
      return { status: "unreachable", message: `Connection failed: ${detail || "Network error"}` };
    }
  }

  async login(baseUrl: string, username: string, password: string) {
    const cleanUrl = baseUrl.replace(/\/$/, "");
    try {
      const response = await axios.post(`${cleanUrl}/api/word/auth/login/`, {
        username,
        password
      }, {
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        }
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
      const status = error.response?.status;
      const detail = error.response?.data?.detail || error.response?.data?.error || error.message;
      console.error('Login Error:', error.response?.data || error.message);
      throw new Error(`Login failed${status ? ` (${status})` : ""}: ${detail}`);
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
    console.log(`[Diagnostic] Syncing to: ${endpoint}`);
    console.log(`[Diagnostic] Auth Header: ${this.tokenType} [REDACTED]`);

    try {
      const response = await axios.post(endpoint, payload, {
        headers: {
          'Authorization': `${this.tokenType} ${this.token}`,
          'Content-Type': 'application/json'
        },
        timeout: 30000 // Extended timeout for document processing
      });
      return response.data;
    } catch (error: any) {
      if (error.response?.status === 401) {
        throw new Error("AUTH_EXPIRED");
      }
      const status = error.response?.status;
      const detail = error.response?.data?.message || error.response?.data?.error || error.message;
      console.error('[Diagnostic] HaloBridge Sync Error:', {
        status,
        data: error.response?.data,
        message: error.message
      });
      
      if (!status) {
        throw new Error(`Network failure: ${error.message}. Check if ${this.baseUrl} is reachable and allows CORS from Office.`);
      }

      throw new Error(`Sync failed (${status}): ${detail}`);
    }
  }
}
