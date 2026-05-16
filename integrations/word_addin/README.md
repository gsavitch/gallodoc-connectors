# GalloDoc Word Connector Add-in

This add-in provides a 3-tier bridge between Microsoft Word and the GalloDoc ecosystem.

## Development & Build

### Prerequisites
- Node.js and npm
- Microsoft Word (Desktop or Online)

### Installation
```bash
cd integrations/word_addin
npm install
```

### Build
To generate a production-ready build in the `dist` folder:
```bash
npm run build
```

### Local Development
To start the dev server over HTTPS (required by Office):
```bash
npm start
```
By default, the server runs at `https://localhost:3000`.

### Sideloading in Word
1. Trust the local development certificate if prompted.
2. Open Word.
3. Go to `Insert` -> `My Add-ins`.
4. Select `Upload My Add-in` and choose the `manifest.xml` file from this directory.

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


