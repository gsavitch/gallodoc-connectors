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

- `halobridge_doc_id`: Unique ID of the document in HaloBridge.
- `halobridge_tenant_id`: Owner tenant ID.
- `halobridge_source_app`: Always `word_connector`.
- `halobridge_original_filename`: Name of the document at creation.
- `halobridge_last_synced_filename`: Current name known to HaloBridge.
- `halobridge_file_fingerprint`: Hash of the document content at last sync.
- `halobridge_version_number`: Iterative version count.
- `halobridge_created_at`: Initial creation timestamp.
- `halobridge_last_synced_at`: Timestamp of last successful sync.
- `halobridge_parent_doc_id`: ID of the document this was branched from (if applicable).
- `halobridge_lineage_reason`: Why this document exists (e.g., `save_as`, `manual_sync`).

## Forbidden Metadata (Do NOT store in Word)
- Raw prompts or AI instructions
- AI outputs or suggestions
- Evidence payloads or raw extraction data
- Trust scores or risk assessments
- Human review details or reviewer IDs
- Credentials, tokens, or sessions
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
