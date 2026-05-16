# HaloBridge / GalloDoc Word Connector MVP

## Purpose
A governed Word-to-GalloDoc bridge that allows lawyers to remain in Microsoft Word while HaloBridge manages the document lifecycle, audit control, and governed extraction.

## Three-Tier Strategy

### 1. Local / Open Source Mode
- **Trust & Privacy**: No data leaves the local machine.
- **Functionality**: Generates a local GalloDoc JSON from the document.
- **Usage**: Good for open-source adoption and local-first workflows.

### 2. Free Connected Mode
- **Cloud Enabled**: Requires HaloBridge login.
- **Functionality**: Saves content to HaloBridge as a cloud record with basic versioning.
- **Scope**: Limited processing, no enterprise workflows.

### 3. Paid Enterprise Connected Mode
- **Full Governance**: Full HaloBridge/Gallo pipeline.
- **Functionality**: Immutable versions, SHA-256 hashing, VerifyIQ checks, HIM-C legal review workflows.
- **Artifacts**: Provenance records, normalized PDFs, and lifecycle audit trails.

## Request / Response Contract

### Save Endpoint
`POST /api/word/gallodoc/save/`

**Payload:**
```json
{
  "mode": "free_connected" | "enterprise_connected",
  "document_name": "Agreement_v1.docx",
  "document_text": "Full text...",
  "ooxml": "Base64 OOXML string...",
  "metadata": { ... },
  "source_app": "microsoft_word",
  "source_connector": "halobridge_word_addin"
}
```

**Response:**
```json
{
  "document_id": "uuid",
  "version_id": "uuid",
  "status": "saved" | "processing",
  "message": "...",
  "processing_url": "..."
}
```

## Sideloading Instructions
1. Navigate to `integrations/word_addin`.
2. Run `npm install` and `npm start`.
3. In Word (Web or Desktop), go to Insert -> Add-ins -> Upload My Add-in.
4. Select `manifest.xml`.

## Security & Privacy
- **Local Mode**: Zero data upload.
- **Encryption**: All connected traffic is over TLS.
- **Immutability**: Once saved to the cloud, versions cannot be edited, only new versions created.
