# GalloDoc Word Connector Add-in

This add-in provides a 3-tier bridge between Microsoft Word and the GalloDoc ecosystem.

## Setup
1. `npm install`
2. `npm start` (Runs a local dev server for the task pane)
3. Sideload the `manifest.xml` in Microsoft Word.

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

## Development
The task pane is built with TypeScript and styles with Tailwind-like utility patterns.
- `src/taskpane/`: Main UI logic.
- `src/lib/`: Core logic for GalloDoc generation and API client.
