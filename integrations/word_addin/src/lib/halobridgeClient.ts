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

  constructor() {}

  setConfiguration(baseUrl: string, token: string | null) {
    this.baseUrl = baseUrl.replace(/\/$/, "");
    this.token = token;
  }

  async testConnection(baseUrl: string): Promise<{ status: "ok" | "unauthorized" | "unreachable", message: string }> {
    const cleanUrl = baseUrl.replace(/\/$/, "");
    try {
      const response = await axios.get(`${cleanUrl}/api/health/`, { timeout: 5000 });
      if (response.status === 200) return { status: "ok", message: "Server Reachable" };
      return { status: "unreachable", message: `Server returned ${response.status}` };
    } catch (error: any) {
      if (error.response?.status === 401) return { status: "unauthorized", message: "Unauthorized" };
      return { status: "unreachable", message: "Network error or invalid URL" };
    }
  }

  async login(baseUrl: string, username: string, password: string) {
    const cleanUrl = baseUrl.replace(/\/$/, "");
    try {
      const response = await axios.post(`${cleanUrl}/api/word/auth/login/`, {
        username,
        password
      });

      const { access_token, user, tenant } = response.data;
      
      this.token = access_token;
      this.baseUrl = cleanUrl;

      return {
        token: access_token,
        user: { username: user.username, display_name: user.display_name },
        tenant: { name: tenant.name }
      };
    } catch (error: any) {
      console.error('Login Error:', error.response?.data || error.message);
      throw new Error(error.response?.data?.message || error.response?.data?.error || "Login failed");
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

    try {
      const response = await axios.post(`${this.baseUrl}/api/word/gallodoc/save/`, payload, {
        headers: {
          'Authorization': `Bearer ${this.token}`
        }
      });
      return response.data;
    } catch (error: any) {
      if (error.response?.status === 401) {
        throw new Error("AUTH_EXPIRED");
      }
      console.error('HaloBridge Sync Error:', error.response?.data || error.message);
      throw new Error(error.response?.data?.message || error.response?.data?.error || "Sync failed");
    }
  }
}
