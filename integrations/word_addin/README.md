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

## Modes
- **Local Mode**: Process documents without an account. Zero data upload.
- **Connected Modes**: Requires a token. Syncs documents as immutable versions to the cloud tenant.

## Development
The task pane is built with TypeScript and styles with Tailwind-like utility patterns.
- `src/taskpane/`: Main UI logic.
- `src/lib/`: Core logic for GalloDoc generation and API client.
