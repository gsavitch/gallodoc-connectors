# GalloDoc Word Connector Add-in

This add-in provides a 3-tier bridge between Microsoft Word and the GalloDoc ecosystem.

## Setup
1. `npm install`
2. `npm start` (Runs a local dev server for the task pane)
3. Sideload the `manifest.xml` in Microsoft Word.

## Authentication
Connected modes require a HaloBridge account.
- **Base URL**: The URL of your HaloBridge/GalloDoc backend (e.g. `http://localhost:3000`).
- **Dev Credentials**: Use `admin` / `gallo` for local development testing.

## Modes
- **Local Mode**: Process documents without an account. Zero data upload.
- **Connected Modes**: Requires a token. Syncs documents as immutable versions to the cloud tenant.

## Development
The task pane is built with TypeScript and styles with Tailwind-like utility patterns.
- `src/taskpane/`: Main UI logic.
- `src/lib/`: Core logic for GalloDoc generation and API client.
