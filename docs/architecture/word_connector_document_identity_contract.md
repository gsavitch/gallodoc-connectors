# Word Connector Document Identity Contract

This document defines how HaloBridge manages document identity and lineage for Microsoft Word documents.

## Overview
The goal is to maintain a stable document identity while allowing for explicit branching (Save As) and tracking lineage across the document lifecycle. The Word file itself stores only the **minimum metadata** required to reconnect with its identity in HaloBridge.

## Save vs. Save As Behavior

| Event | Logic | Outcome |
| :--- | :--- | :--- |
| **Initial Save** | No embedded ID detected. | Creates new `MvpDocument`. Returns new ID. |
| **Normal Save** | Embedded ID exists + same/known file identity. | Syncs content to existing `MvpDocument`. Creates new version. |
| **Save As** | Explicit `save_as` event or new file identity with existing ID. | Creates **new** `MvpDocument`. Links to parent ID. Returns new ID to be written back to the new file. |
| **Rename** | Filename changes but Word file identity is stable. | Updates `halobridge_last_synced_filename`. Does not create new doc ID. |

## Metadata Allowlist (Stored in Word)
These fields are stored as Word Custom Properties. No sensitive data or rich AI/governance payloads should ever be written back to the document.

### Structured Manifest (`gallodoc_manifest`)
A JSON string containing the full session state.

### Flat Metadata (Portable Identity)
For portability (e.g., conversion to PDF), several flat properties are written:
- `HaloBridgeDocumentId`: Maps to `halobridge_doc_id`.
- `GalloDocId`: Maps to `gallodoc_id`.
- `HaloBridgeWorkspaceUrl`: Maps to `canonical_workspace_url`.
- `GalloDocVersion`: Maps to `latest_version_number`.
- `HaloBridgeSourceHash`: Maps to `last_source_hash`.
- `HaloBridgeTenantId`: Maps to `halobridge_tenant_id`.

## Recovery Logic
On loading a document, the connector follows this priority:
1. Parse `gallodoc_manifest`.
2. If missing or corrupt, reconstruct the manifest from the flat properties listed above.
3. If reconstructed, the manifest is flagged with `_source: "recovered_from_flat_metadata"`.

## Universal Relationship Principle
Relationships between documents (Word, Outlook, etc.) use canonical `DocumentRelationship` records in HaloBridge. No Office-specific relationship tables should be created.

### Planned Relationship Types
- `EMAIL_CONTAINS_ATTACHMENT`
- `ATTACHMENT_OF_EMAIL`
- `SENT_WITH`
- `FORWARDED_WITH`
- `REPLY_TO`
- `REDLINE_OF`
- `APPROVAL_FOR`
- `APPROVED_VERSION_OF`
- `DERIVED_FROM`
- `SAVED_AS`

## Future Outlook Integration
When Outlook support is added, emails will be treated as GalloDocs. 

### Expected Outlook Metadata (`metadata.source_reference`)
```json
{
  "source_system": "outlook",
  "message_id": "...",
  "internet_message_id": "...",
  "conversation_id": "...",
  "mailbox_user": "...",
  "sent_at": "...",
  "from": "...",
  "to_hashes": [],
  "subject_hash": "...",
  "attachment_ids": [],
  "web_url": "...",
  "source_tenant_id": "..."
}
```

### Attachment Continuity
If a Word document with an existing GalloDoc ID is attached to an email, the Outlook connector will preserve that identity by reading the embedded manifest.

## AI & Copilot Provenance Signals
The connector captures advisory signals of AI assistance to help inform governance decisions.

### AI Context Payload (`ai_context`)
```json
{
  "ai_assistance_detected": true,
  "ai_signal_confidence": "medium",
  "ai_signal_sources": ["text_keywords"],
  "copilot_markers_detected": null,
  "ai_mentions_in_comments": 0,
  "ai_mentions_in_metadata": 0,
  "connector_detection_notes": ["Found AI-related keywords in document text"]
}
```

### Manifest AI Summary (Safe Fields)
- `ai_assistance_status`: "detected" | "unknown"
- `ai_signal_confidence`: "none" | "low" | "medium" | "high"
- `human_review_required`: boolean

## Forbidden Metadata (Do NOT store in Word)
- Raw prompts or AI instructions
- AI outputs or suggestions
- Evidence payloads or raw extraction data
- Trust scores or risk assessments
- Human review details or reviewer IDs
- **API Tokens, Passwords, or Secrets**
- Internal file paths or stack traces

## Tenant Isolation
The Word Connector MUST enforce tenant isolation. If a document with a valid `halobridge_doc_id` is synced by a user belonging to a different tenant, the request MUST be rejected without leaking existence or metadata of the document.

## Lifecycle Stages
The following stages are appended to the document lifecycle during sync:
- `word_document_created`
- `word_document_saved`
- `word_save_as_detected`
- `word_document_renamed`
- `new_document_id_created`
- `document_lineage_linked`
- `gallodoc_initialized`

## Lineage & Relationships
On `save_as`, the system creates a relationship:
`[Parent Doc] -> (save_as_copy) -> [New Doc]`
This allows the HaloBridge Linker to visualize the provenance of the document even if the raw Word lineage is lost.

## Examples

### Create (Initial Sync)
Request: `embedded_halobridge_doc_id: null`
Response Action: `created`
Metadata: `halobridge_doc_id` is newly generated.

### Branch (Save As)
Request: `embedded_halobridge_doc_id: "doc_123"`, `event_type: "save_as"`
Response Action: `branched`
Metadata: `halobridge_doc_id` is a new ID, `halobridge_parent_doc_id` is `"doc_123"`.
