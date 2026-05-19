# HaloBridge Word Connector Installation Guide

This guide describes how to install the HaloBridge Word Connector for production use or development.

## Production Installation (Recommended)

For organization-wide deployment, administrators should use the Microsoft 365 Admin Center.

### 1. Unified Deployment (Admin)
1. Go to the **Microsoft 365 Admin Center**.
2. Navigate to **Settings** > **Integrated apps**.
3. Click **Upload custom apps**.
4. Choose **Upload manifest file (.xml) from device**.
5. Select the `manifest.prod.xml` file.
6. Configure access (e.g., "Entire organization" or specific users).
7. Complete the wizard. The "GalloDoc" tab will appear in Word for users within 24 hours.

### 2. Manual Upload (Individual User)
If your organization allows it, you can upload the manifest individually:
1. Open **Word** (Desktop or Web).
2. Go to the **Insert** tab.
3. Click **Add-ins** > **My Add-ins**.
4. Click **Upload My Add-in**.
5. Select `manifest.prod.xml`.
6. The HaloBridge/GalloDoc Connector will appear in the Home tab.

---

## Developer Sideloading

Developers working on the connector can use local source code.

### 1. Prerequisites
- Node.js installed.
- Repository cloned.
- Local certificates generated (run `npx office-addin-dev-certs install`).

### 2. Running Locally
1. Navigate to `integrations/word_addin`.
2. Run `npm install`.
3. Run `npm start`. This starts the dev server at `https://localhost:3000`.
4. Sideload the `manifest.xml` (local version) into Word using the "Manual Upload" steps above.

---

## Hosting the Add-in

The production add-in is hosted at:
`https://www.halobridge.ai/word-addin/`

To update the hosted version, build the project:
1. `cd integrations/word_addin`
2. `npm run build`
3. Upload the contents of `dist/word-addin/` to the static web server at the `/word-addin/` path.

### Configuration
The add-in automatically detects if it is running on `halobridge.ai` and defaults the connection URL to `https://www.halobridge.ai`.
