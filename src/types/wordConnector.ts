
export type WordConnectorEventType = 'save' | 'save_as' | 'rename' | 'open' | 'manual_sync' | 'create';
export type WordConnectorLineageReason = 'normal_save' | 'save_as' | 'rename' | 'manual_sync' | 'save_as_copy_same_content';
export type WordConnectorAction = 'created' | 'updated' | 'branched' | 'metadata_updated' | 'updated_no_content_change';

export interface WordConnectorMetadata {
  halobridge_doc_id: string;
  halobridge_tenant_id: string;
  halobridge_source_app: string;
  halobridge_original_filename: string;
  halobridge_last_synced_filename: string;
  halobridge_file_fingerprint: string;
  halobridge_version_number: number;
  halobridge_created_at: string;
  halobridge_last_synced_at: string;
  halobridge_parent_doc_id?: string;
  halobridge_lineage_reason?: WordConnectorLineageReason;
  // Governance fields
  review_status?: string;
  approval_status?: string;
  verifyiq_status?: string;
  himc_status?: string;
  release_ready?: boolean;
  canonical_workspace_url?: string;
}

export interface WordReviewContext {
  tracked_changes_detected?: boolean | null;
  comments_detected?: boolean | null;
  unresolved_comments_count?: number | null;
  document_protection?: string | null;
  office_host: "Word";
  capture_method: "word_addin";
}

export interface WordAIContext {
  ai_assistance_detected?: boolean | null;
  ai_signal_confidence: "none" | "low" | "medium" | "high";
  ai_signal_sources: string[];
  copilot_markers_detected?: boolean | null;
  ai_mentions_in_comments?: number | null;
  ai_mentions_in_metadata?: number | null;
  connector_detection_notes: string[];
}

export interface WordConnectorSyncPayload {
  tenant_id: string;
  user_id: string;
  event_type: WordConnectorEventType;
  current_filename: string;
  previous_filename?: string;
  word_file_identity: string;
  previous_word_file_identity?: string;
  embedded_halobridge_doc_id?: string;
  embedded_parent_doc_id?: string;
  file_hash: string;
  content_hash: string;
  review_context?: WordReviewContext;
  ai_context?: WordAIContext;
  metadata: Partial<WordConnectorMetadata>;
}

export interface WordConnectorSyncResponse {
  action: WordConnectorAction;
  halobridge_doc_id: string;
  halobridge_parent_doc_id?: string;
  halobridge_lineage_reason?: WordConnectorLineageReason;
  halobridge_original_filename: string;
  halobridge_last_synced_filename: string;
  halobridge_file_fingerprint: string;
  halobridge_version_number: number;
  halobridge_created_at: string;
  halobridge_last_synced_at: string;
  write_back_metadata: Partial<WordConnectorMetadata>;
}
