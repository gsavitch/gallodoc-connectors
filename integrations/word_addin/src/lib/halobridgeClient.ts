import axios from 'axios';
import { ConnectorMode } from './modeConfig';

export interface ConnectionInfo {
  baseUrl: string;
  user: {
    username: string;
    display_name: string;
  };
  tenant: {
    name: string;
  };
}

export class HaloBridgeClient {
  private apiUrl: string | null = null;
  private token: string | null = null;
  private connection: ConnectionInfo | null = null;

  constructor() {
    this.restoreSession();
  }

  private restoreSession() {
    const savedToken = localStorage.getItem('hb_token');
    const savedUrl = localStorage.getItem('hb_url');
    const savedConn = localStorage.getItem('hb_connection');

    if (savedToken && savedUrl && savedConn) {
      this.token = savedToken;
      this.apiUrl = savedUrl;
      this.connection = JSON.parse(savedConn);
    }
  }

  async login(baseUrl: string, username: string, password: string): Promise<ConnectionInfo> {
    try {
      const response = await axios.post(`${baseUrl}/api/word/auth/login/`, {
        username,
        password
      });

      const { access_token, user, tenant } = response.data;
      
      this.token = access_token;
      this.apiUrl = baseUrl;
      this.connection = {
        baseUrl,
        user: { username: user.username, display_name: user.display_name },
        tenant: { name: tenant.name }
      };

      localStorage.setItem('hb_token', this.token!);
      localStorage.setItem('hb_url', this.apiUrl!);
      localStorage.setItem('hb_connection', JSON.stringify(this.connection));

      return this.connection;
    } catch (error: any) {
      console.error('Login Error:', error.response?.data || error.message);
      throw new Error(error.response?.data?.error || "Connection failed");
    }
  }

  logout() {
    this.token = null;
    this.apiUrl = null;
    this.connection = null;
    localStorage.removeItem('hb_token');
    localStorage.removeItem('hb_url');
    localStorage.removeItem('hb_connection');
  }

  getConnection(): ConnectionInfo | null {
    return this.connection;
  }

  isLoggedIn(): boolean {
    return !!this.token;
  }

  async saveWordDocument(payload: {
    mode: ConnectorMode;
    document_name: string;
    document_text: string;
    ooxml?: string;
    metadata?: any;
  }) {
    if (!this.apiUrl || !this.token) {
      throw new Error("Connect to HaloBridge before saving in this mode.");
    }

    try {
      const response = await axios.post(`${this.apiUrl}/api/word/gallodoc/save/`, payload, {
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${this.token}`
        }
      });
      return response.data;
    } catch (error: any) {
      if (error.response?.status === 401) {
        this.logout();
        throw new Error("Session expired. Please reconnect.");
      }
      console.error('HaloBridge Sync Error:', error.response?.data || error.message);
      throw new Error(error.response?.data?.message || error.response?.data?.error || "Sync failed");
    }
  }
}
