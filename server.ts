import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { GoogleGenerativeAI } from "@google/generative-ai";
import axios from "axios";
import cookieParser from "cookie-parser";
import cors from "cors";
import dotenv from "dotenv";
import { z } from "zod";
import { v4 as uuidv4 } from "uuid";
import crypto from "crypto";
import fs from "fs";

dotenv.config();

const app = express();
const PORT = 3000;

app.use(express.json({ limit: '10mb' }));
app.use(cookieParser());

// CORS Configuration
const allowedOrigins = process.env.HALOBRIDGE_WORD_CONNECTOR_ALLOWED_ORIGINS 
  ? process.env.HALOBRIDGE_WORD_CONNECTOR_ALLOWED_ORIGINS.split(',').map(o => o.trim())
  : ["https://localhost:3000", "https://127.0.0.1:3000"];

// Add the app's own URL if configured
if (process.env.APP_URL) {
  const appUrl = process.env.APP_URL.replace(/\/$/, "");
  if (!allowedOrigins.includes(appUrl)) {
    allowedOrigins.push(appUrl);
  }
}

app.use(cors({
  origin: function (origin, callback) {
    // Allow requests with no origin (like mobile apps, curl, or same-origin)
    if (!origin) return callback(null, true);
    
    // Check against allowedOrigins list
    if (allowedOrigins.indexOf(origin) !== -1) {
      return callback(null, true);
    }
    
    // Heuristic: Allow origins from the same platform (AI Studio / Cloud Run) for development
    if (origin.endsWith('.run.app') || origin.startsWith('http://localhost:') || origin.startsWith('https://localhost:')) {
      return callback(null, true);
    }

    console.warn(`[CORS] Blocked request from unauthorized origin: ${origin}`);
    callback(new Error('Not allowed by CORS'));
  },
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'Accept'],
  credentials: true
}));

// Mock In-Memory Database for MVP
// In a real app, use Firebase or SQL
const documents: any[] = [];
const versions: any[] = [];
const relationships: any[] = [];
const lifecycleStages: any[] = [];
const tokens: Record<string, any> = {}; // token -> { user, tenant, entitlements }

import { WordConnectorIdentityService } from "./src/services/wordConnectorIdentityService";
import { WordConnectorSyncPayload } from "./src/types/wordConnector";

const wordIdentityService = new WordConnectorIdentityService(documents, versions, relationships, lifecycleStages);

// Mock User for Dev
const MOCK_USER = {
  id: "user_123",
  username: "gsavitch",
  email: "gsavitch@example.com",
  display_name: "G. Savitch"
};

const MOCK_TENANT = {
  id: "tenant_999",
  slug: "lex-corp",
  name: "Lex Corp Legal"
};

// Gemini Initialization
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY || "");

// Recursive function to get all files for sync
function getProjectFiles(dir: string, allFiles: string[] = []) {
  const files = fs.readdirSync(dir);
  for (const file of files) {
    const name = path.join(dir, file);
    if (fs.statSync(name).isDirectory()) {
      // Skip heavy or sensitive directories
      if (['node_modules', '.git', 'dist', '.next', '.vscode'].includes(file)) continue;
      getProjectFiles(name, allFiles);
    } else {
      allFiles.push(name);
    }
  }
  return allFiles;
}

// --- API ROUTES ---

// Health Check for connectivity testing
app.get('/api/health/', (req, res) => {
  res.json({ status: "ok", version: "1.0.0", instance: "HaloBridge-Mock-Dev" });
});

// GitHub Project Sync Endpoint
app.post('/api/github/push-project', async (req, res) => {
  const token = req.cookies.github_token;
  if (!token) return res.status(401).json({ error: 'GitHub authentication required' });

  try {
    const owner = 'gsavitch';
    const repo = 'gallodoc-connectors';
    const files = getProjectFiles(process.cwd());
    const pushResults = [];

    for (const filePath of files) {
      const relativePath = path.relative(process.cwd(), filePath);
      // Skip self (don't push .env or other potential secrets in a real app, though here we're selective)
      if (relativePath.includes('.env')) continue;

      const content = fs.readFileSync(filePath, 'base64');
      
      let sha;
      try {
        const existing = await axios.get(`https://api.github.com/repos/${owner}/${repo}/contents/${relativePath}`, {
          headers: { Authorization: `token ${token}` }
        });
        sha = existing.data.sha;
      } catch (e) {}

      await axios.put(`https://api.github.com/repos/${owner}/${repo}/contents/${relativePath}`, {
        message: `Sync project file: ${relativePath}`,
        content: content,
        sha: sha
      }, {
        headers: { Authorization: `token ${token}` }
      });
      
      pushResults.push(relativePath);
    }

    res.json({ success: true, filesSynced: pushResults.length });
  } catch (error: any) {
    console.error('GitHub Push Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to push files to GitHub', details: error.response?.data });
  }
});

// Word Connector Login
app.post('/api/word/auth/login/', (req, res) => {
  const { username, password } = req.body;

  // Simple mock validation
  if (username === "admin" && password === "gallo") {
    const accessToken = crypto.randomBytes(32).toString('hex');
    const authData = {
      user: MOCK_USER,
      tenant: MOCK_TENANT,
      entitlements: {
        word_connector: true,
        free_connected: true,
        enterprise_connected: true, // Mocking full access for admin
        verifyiq: true,
        himc_review: true
      }
    };
    
    tokens[accessToken] = authData;

    return res.json({
      access_token: accessToken,
      token_type: "Bearer",
      expires_in: 3600,
      ...authData
    });
  }

  res.status(401).json({ error: "Invalid username or password" });
});

// Word Connector Schema
const WordSaveSchema = z.object({
  mode: z.enum(['free_connected', 'enterprise_connected']),
  document_name: z.string().min(1),
  document_text: z.string().optional(),
  ooxml: z.string().optional(),
  save_action: z.enum(['save', 'save_as']).optional().default('save'),
  document_id: z.string().optional(),
  source_document_id: z.string().optional(),
  metadata: z.record(z.string(), z.any()).optional(),
  source_app: z.string().default('microsoft_word'),
  source_connector: z.string().default('halobridge_word_addin'),
  tier_scope: z.enum(['limited', 'full']).optional(),
});

const WordSyncSchema = z.object({
  tenant_id: z.string(),
  event_type: z.enum(['save', 'save_as', 'rename', 'open', 'manual_sync']),
  current_filename: z.string(),
  previous_filename: z.string().optional(),
  word_file_identity: z.string(),
  previous_word_file_identity: z.string().optional(),
  embedded_halobridge_doc_id: z.string().optional(),
  embedded_parent_doc_id: z.string().optional(),
  file_hash: z.string(),
  content_hash: z.string(),
  metadata: z.record(z.string(), z.any()).optional().default({}),
});

// Word Connector Sync Endpoint
app.post('/api/word/connector/sync/', async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader || (!authHeader.startsWith('Bearer ') && !authHeader.startsWith('Token '))) {
    return res.status(401).json({ error: "not_authenticated", message: "Bearer or Token required." });
  }

  const token = authHeader.split(' ')[1];
  const authSession = tokens[token];

  if (!authSession) {
    return res.status(401).json({ error: "not_authenticated", message: "Invalid session." });
  }

  try {
    const data = WordSyncSchema.parse(req.body);
    
    // Safety: ensure tenant_id matches session
    if (data.tenant_id !== authSession.tenant.id) {
      return res.status(403).json({ error: "unauthorized", message: "Tenant mismatch." });
    }

    const payload: WordConnectorSyncPayload = {
      ...data,
      user_id: authSession.user.id
    };

    const result = wordIdentityService.resolveWordDocumentIdentity(payload);
    res.status(200).json(result);

  } catch (error: any) {
    if (error instanceof z.ZodError) {
      return res.status(400).json({ error: "validation_error", issues: error.issues });
    }
    if (error.message.startsWith("UNAUTHORIZED_ACCESS")) {
      return res.status(403).json({ error: "unauthorized", message: error.message });
    }
    console.error('Sync Error:', error);
    res.status(500).json({ error: "internal_error", message: error.message || "Failed to sync document." });
  }
});

// Word Connector Endpoint
app.post('/api/word/gallodoc/save/', async (req, res) => {
  // Check for Bearer or Token
  const authHeader = req.headers.authorization;
  if (!authHeader || (!authHeader.startsWith('Bearer ') && !authHeader.startsWith('Token '))) {
    return res.status(401).json({ error: "not_authenticated", message: "Bearer or Token required for connected modes." });
  }

  const token = authHeader.split(' ')[1];
  const authSession = tokens[token];

  if (!authSession) {
    return res.status(401).json({ error: "not_authenticated", message: "Invalid or expired session." });
  }

  try {
    const data = WordSaveSchema.parse(req.body);
    
    // Entitlement Check
    if (data.mode === 'free_connected' && !authSession.entitlements.free_connected) {
      return res.status(403).json({ error: "missing_entitlement", message: "Free mode not enabled for this user." });
    }
    if (data.mode === 'enterprise_connected' && !authSession.entitlements.enterprise_connected) {
      return res.status(403).json({ error: "missing_entitlement", message: "Enterprise mode requires a higher subscription." });
    }

    if (!data.document_text && !data.ooxml) {
      return res.status(400).json({ error: "missing_document_content", message: "No text or OOXML provided." });
    }

    // Determine source material for hashing
    const sourceContent = data.ooxml || data.document_text || data.document_name || "";
    const sourceHash = crypto.createHash('sha256').update(sourceContent).digest('hex');

    // Find or Create Document (bound to tenant)
    let doc;
    if (data.save_action === 'save_as' && data.source_document_id) {
      // Create a NEW document linked to source
      doc = {
        id: uuidv4(),
        name: `${data.document_name} (Copy)`,
        tenant_id: authSession.tenant.id,
        source_document_id: data.source_document_id,
        created_at: new Date().toISOString(),
      };
      documents.push(doc);
    } else if (data.document_id) {
      doc = documents.find(d => d.id === data.document_id && d.tenant_id === authSession.tenant.id);
    }

    if (!doc) {
      // Fallback or Initial Save
      doc = documents.find(d => d.name === data.document_name && d.tenant_id === authSession.tenant.id);
      if (!doc) {
        doc = {
          id: uuidv4(),
          name: data.document_name,
          tenant_id: authSession.tenant.id,
          created_at: new Date().toISOString(),
        };
        documents.push(doc);
      }
    }

    // Determine tier scope based on mode if not provided
    const tierScope = data.tier_scope || (data.mode === 'enterprise_connected' ? 'full' : 'limited');

    // Create New Immutable Version
    const versionNumber = versions.filter(v => v.document_id === doc.id).length + 1;
    const version = {
      id: uuidv4(),
      document_id: doc.id,
      version_number: versionNumber,
      source_hash: sourceHash,
      source_connector: data.source_connector,
      connector_mode: data.mode,
      tier_scope: tierScope,
      user_id: authSession.user.id,
      tenant_id: authSession.tenant.id,
      original_filename: data.document_name,
      metadata: data.metadata || {},
      created_at: new Date().toISOString(),
      submitted_at: new Date().toISOString(),
      content_preview: data.document_text?.substring(0, 500),
    };

    versions.push(version);

    // Enterprise Pipeline Simulation
    let responseMessage = "Word document saved to HaloBridge as a governed GalloDoc version.";
    let status = 'saved';
    let processingStatus = 'pending_pipeline';
    let processingUrl = null;

    if (data.mode === 'enterprise_connected') {
      status = 'processing';
      processingStatus = 'active_review';
      processingUrl = `/ops/documents/${doc.id}/word-control/`;
      
      setTimeout(() => {
        console.log(`[PIPELINE] Processing enterprise version ${version.id} for user ${authSession.user.username}`);
      }, 100);
    }

    res.status(201).json({
      document_id: doc.id,
      version_id: version.id,
      version_number: versionNumber,
      status,
      processing_status: processingStatus,
      mode: data.mode,
      tier_scope: tierScope,
      message: responseMessage,
      processing_url: processingUrl
    });

  } catch (error: any) {
    if (error instanceof z.ZodError) {
      return res.status(400).json({ error: "validation_error", issues: error.issues });
    }
    console.error('Save Error:', error);
    res.status(500).json({ error: "internal_error", message: "Failed to save document version." });
  }
});

// GitHub OAuth
app.get('/api/auth/github/url', (req, res) => {
  const GITHUB_CLIENT_ID = process.env.GITHUB_CLIENT_ID;
  if (!GITHUB_CLIENT_ID) {
    return res.status(500).json({ error: "GITHUB_CLIENT_ID is not configured" });
  }

  // Ensure APP_URL doesn't have a trailing slash for consistency
  const baseUrl = (process.env.APP_URL || `http://${req.headers.host}`).replace(/\/$/, "");
  const redirectUri = `${baseUrl}/auth/github/callback`;
  
  const params = new URLSearchParams({
    client_id: GITHUB_CLIENT_ID,
    redirect_uri: redirectUri,
    scope: 'repo user',
  });

  const authUrl = `https://github.com/login/oauth/authorize?${params.toString()}`;
  res.json({ url: authUrl });
});

app.get(['/auth/github/callback', '/auth/github/callback/'], async (req, res) => {
  const { code } = req.query;
  
  try {
    const response = await axios.post('https://github.com/login/oauth/access_token', {
      client_id: process.env.GITHUB_CLIENT_ID,
      client_secret: process.env.GITHUB_CLIENT_SECRET,
      code,
    }, {
      headers: { Accept: 'application/json' }
    });

    const body = response.data;
    if (body.error) {
      throw new Error(body.error_description || body.error);
    }

    const accessToken = body.access_token;

    // Securely set session cookie
    res.cookie('github_token', accessToken, {
      httpOnly: true,
      secure: true,
      sameSite: 'none',
      maxAge: 30 * 24 * 60 * 60 * 1000 // 30 days
    });

    res.send(`
      <html>
        <body>
          <script>
            if (window.opener) {
              window.opener.postMessage({ type: 'OAUTH_AUTH_SUCCESS' }, '*');
              window.close();
            } else {
              window.location.href = '/';
            }
          </script>
          <p>Authentication successful. You can close this window.</p>
        </body>
      </html>
    `);
  } catch (error: any) {
    console.error('GitHub Auth Error:', error.response?.data || error.message);
    res.status(500).send('Authentication failed');
  }
});

app.get('/api/auth/github/me', async (req, res) => {
  const token = req.cookies.github_token;
  if (!token) return res.status(401).json({ error: 'Not authenticated' });

  try {
    const userRes = await axios.get('https://api.github.com/user', {
      headers: { Authorization: `token ${token}` }
    });
    res.json(userRes.data);
  } catch (error) {
    res.status(401).clearCookie('github_token').json({ error: 'Session expired' });
  }
});

app.post('/api/auth/github/logout', (req, res) => {
  res.clearCookie('github_token').json({ success: true });
});

// Gemini Endpoint
app.post('/api/gemini/suggest-connectors', async (req, res) => {
  const { repoUrl, userDescription } = req.body;

  try {
    const model = genAI.getGenerativeModel({ model: "gemini-3-flash-preview" });
    const response = await model.generateContent({
      contents: [{
        role: "user",
        parts: [{
          text: `You are an expert system architect. An user wants to connect their repository ${repoUrl} to a "GalloDoc Connector Hub".
      User Description: ${userDescription}
      
      Suggest 3 data connectors (e.g., Slack, Jira, Notion, Google Drive) that would be most useful for this repository.
      For each connector, provide:
      1. Name
      2. Benefit
      3. A short setup description.
      
      Format as JSON.`
        }]
      }],
      generationConfig: {
        responseMimeType: "application/json",
      }
    });

    res.json({ suggestions: JSON.parse(response.response.text()) });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
});

async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
