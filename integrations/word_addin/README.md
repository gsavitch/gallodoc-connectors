# GalloDoc Word Connector Add-in

This add-in provides a 3-tier bridge between Microsoft Word and the GalloDoc ecosystem.

## Installation & Setup

For a streamlined setup on Windows, use the provided PowerShell installer:

```powershell
./scripts/install-word-connector.ps1
```

This script handles dependency installation, dev certificate setup, and a full build.

### Manual Setup
1. **Install Dependencies**:
   ```bash
   cd integrations/word_addin
   npm install
   ```
2. **Dev Certificates**:
   ```bash
   npx office-addin-dev-certs install
   ```
3. **Start Dev Server**:
   ```bash
   npm start
   ```

## Local Development
By default, the server runs at `https://localhost:3000`. Office add-ins **require** HTTPS. 

### Sideloading in Word
1. Trust the local development certificate if prompted.
2. Open Word Desktop.
3. Run: `npx office-addin-debugging start manifest.xml desktop`

## Production Deployment

When you are ready to move from `localhost` to production:

1. **Host Assets**: 
   Upload the contents of the `dist` folder to your production web server (e.g., `https://www.halobridge.ai/word-addin/`).
2. **Update Manifest**: 
   Create a production version of `manifest.xml` where all `https://localhost:3000` URLs are replaced with your production URL.
3. **AppSource / Admin Center**: 
   Submit the production manifest to the Microsoft 365 Admin Center (for internal organization deployment) or the Microsoft AppSource store.

> [!IMPORTANT]
> Ensure strictly valid HTTPS certificates are used in production. Office will block any non-secure or invalidly signed task pane content.

## Troubleshooting
- **Cannot resolve './src'**: This typically happened before `webpack.config.js` was added. Ensure you are running commands from the `integrations/word_addin` directory.
- **HTTPS Issues**: Office add-ins require HTTPS. `npm start` uses a self-signed certificate. You may need to visit `https://localhost:3000/taskpane.html` in your browser and choose "Advanced" -> "Proceed anyway" to trust the certificate before it works inside Word.
- **OfficeRuntime.storage missing**: Ensure you are using a supported version of Office (2019+ or Microsoft 365).

### Common Build Fixes
- **Missing @types/uuid**: If you see `Could not find declaration file for module 'uuid'`, run `npm install --save-dev @types/uuid`.
- **OfficeRuntime Typing**: `OfficeRuntime` is a global available in Office Add-in shared runtimes. For TypeScript, it is declared as `any` in our source to allow for safe runtime checks and browser fallbacks.
- **Word Body Text Extraction**: Always use `body.load("text")` and `await context.sync()` to read document content. The older `getText()` property is not supported in modern Office.js `Word.run` blocks.

## Persistent Connection Settings
The Word Add-in features a persistent connection settings panel.
- **OfficeRuntime.storage**: Used when available (shared runtime).
- **LocalStorage**: Fallback using `Office.context.partitionKey` for isolation.

### Configuration
1. **HaloBridge Base URL**: Fully configurable. Supports localhost, QA, and production.
2. **Authentication Modes**:
   - **Username/Password**: Standard login flow. Clears password from memory immediately after obtaining a token.
   - **API Token**: Direct token injection for automation or specialized accounts.

### Development Testing
- **Base URL**: Set to `http://localhost:3000` (or your local backend URL).
- **Credentials**: Use `admin` / `gallo`.
- **Test Connection**: Use the "Test" button to verify endpoint status without logging in.

## Security & Persistence
- Plaintext passwords are **never** persisted.
- Tokens and URLs are persisted securely within the Office add-in partition.
- All cloud communications use dynamic base URLs.

## Embedded GalloDoc Manifest
Connected documents store a minimal metadata manifest within the Word document's Custom Properties. This allows the document to maintain its identity across sessions and users.

### Stored Fields
- `mvp_document_id`: The authoritative ID in HaloBridge.
- `latest_version_id`: ID of the last synced version.
- `latest_version_number`: Human-readable version counter.
- `last_synced_at`: ISO timestamp of the last successful sync.
- `review_status`: Current governing workflow status (e.g., "draft").

### Security Note
- **No Secrets**: API tokens, passwords, and PII are never stored in the document.
- **No Content**: The Word document does not store an audit log of past versions locally; HaloBridge remains the source of truth.

## Auto-sync (Governed)
The Word Connector supports periodic auto-sync for connected modes (**Free Connected** and **Enterprise Connected**).

- **Safety First**: Auto-sync is **OFF** by default and only activates after you have manually saved the document to HaloBridge at least once. 
- **Change Detection**: The connector extracts document content and compares it against the last synced version's hash. If no changes are detected, the sync is skipped to save resources and avoid redundant versions.
- **Governed Versions**: Unlike standard "autosave" which might stream every keystroke, HaloBridge Auto-sync creates a new immutable version in the registry at set intervals (2, 5, or 10 minutes), preserving the governance trail.
- **Local Mode Privacy**: Auto-sync is strictly disabled in **Local Mode**. No document data will ever leave your device automatically in this mode.
- **Lifecycle**: The sync timer is automatically managed and will stop if you disconnect, sign out, or change to Local Mode.

### How to use Auto-sync
1. Connect to your HaloBridge instance.
2. Select a connected mode.
3. Save the document manually once to link it to a GalloDoc registry.
4. Toggle **Enable Auto-sync** in the task pane.
5. Choose your preferred interval.

## Save vs. Save As
1. **Save to HaloBridge**:
   - Updates the existing GalloDoc document with a new immutable version.
   - Preserves the `mvp_document_id`.
   
2. **Save As New GalloDoc**:
   - Creates a **new** guided document in HaloBridge.
   - Links the new document to the original via `source_document_id`.
   - Overwrites the local Word manifest with the new document identity.
   - Useful for branching, templating, or creating divergent legal positions from a common base.

## Modes
- **Local Mode**: Process documents without an account. Zero data upload. Does not write cloud manifests.
- **Connected Modes**: Requires a token. Syncs documents as immutable versions to the cloud tenant. Writes/updates embedded manifests.


