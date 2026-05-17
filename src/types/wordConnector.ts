
export type WordConnectorEventType = 'save' | 'save_as' | 'rename' | 'open' | 'manual_sync';
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
