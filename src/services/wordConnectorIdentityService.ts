
import { v4 as uuidv4 } from "uuid";
import { 
  WordConnectorSyncPayload, 
  WordConnectorSyncResponse, 
  WordConnectorMetadata,
  WordConnectorAction,
  WordConnectorLineageReason
} from "../types/wordConnector";

export class WordConnectorIdentityService {
  private documents: any[];
  private versions: any[];
  private relationships: any[];
  private lifecycleStages: any[];

  constructor(documents: any[], versions: any[], relationships: any[] = [], lifecycleStages: any[] = []) {
    this.documents = documents;
    this.versions = versions;
    this.relationships = relationships;
    this.lifecycleStages = lifecycleStages;
  }

  public resolveWordDocumentIdentity(payload: WordConnectorSyncPayload): WordConnectorSyncResponse {
    const {
      tenant_id,
      user_id,
      event_type,
      current_filename,
      word_file_identity,
      embedded_halobridge_doc_id,
      file_hash,
      content_hash,
      metadata
    } = payload;

    let action: WordConnectorAction = 'updated';
    let docId = embedded_halobridge_doc_id;
    let parentDocId = payload.embedded_parent_doc_id;
    let lineageReason: WordConnectorLineageReason | undefined = metadata.halobridge_lineage_reason;

    // 5. Tenant Isolation Check
    if (docId) {
      const existingDoc = this.documents.find(d => d.id === docId);
      if (existingDoc && existingDoc.tenant_id !== tenant_id) {
        // Safe rejection - do not reveal if it exists or not
        throw new Error("UNAUTHORIZED_ACCESS: Document identity mismatch for tenant.");
      }
    }

    // Identify current document or create new one
    let currentDoc = docId ? this.documents.find(d => d.id === docId && d.tenant_id === tenant_id) : null;

    // 1. If create or no embedded_halobridge_doc_id: Create new
    if (event_type === 'create' || !docId || !currentDoc) {
      action = 'created';
      currentDoc = this.createNewDocument(tenant_id, current_filename, 'word_document_created', 'normal_save');
      docId = currentDoc.id;
      lineageReason = 'normal_save';
    } 
    // 3. If event_type is save_as: Branch
    else if (event_type === 'save_as') {
      action = 'branched';
      parentDocId = docId;
      lineageReason = content_hash === this.getLatestContentHash(docId) ? 'save_as_copy_same_content' : 'save_as';
      
      const newDoc = this.createNewDocument(tenant_id, current_filename, 'word_save_as_detected', lineageReason, parentDocId);
      
      // Link lineage
      this.linkRelationship(parentDocId, newDoc.id, 'save_as_copy', lineageReason);
      this.addLifecycleStage(newDoc.id, 'new_document_id_created');
      this.addLifecycleStage(newDoc.id, 'document_lineage_linked');
      this.addLifecycleStage(newDoc.id, 'gallodoc_initialized');
      
      currentDoc = newDoc;
      docId = newDoc.id;
    }
    // 4. If rename but same file identity: update metadata only
    else if (event_type === 'rename' && payload.word_file_identity === payload.previous_word_file_identity) {
      action = 'metadata_updated';
      this.addLifecycleStage(docId, 'word_document_renamed');
    }
    // 2. Normal Save
    else if (event_type === 'save') {
      const latestHash = this.getLatestContentHash(docId);
      if (content_hash === latestHash) {
        action = 'updated_no_content_change';
      } else {
        action = 'updated';
      }
      this.addLifecycleStage(docId, 'word_document_saved');
    }

    // Sync content/version
    const versionNumber = this.createVersion(docId, tenant_id, user_id, current_filename, file_hash, content_hash);

    const now = new Date().toISOString();
    const responseMetadata: WordConnectorMetadata = {
      halobridge_doc_id: docId,
      halobridge_tenant_id: tenant_id,
      halobridge_source_app: 'word_connector',
      halobridge_original_filename: currentDoc.original_filename || current_filename,
      halobridge_last_synced_filename: current_filename,
      halobridge_file_fingerprint: file_hash,
      halobridge_version_number: versionNumber,
      halobridge_created_at: currentDoc.created_at,
      halobridge_last_synced_at: now,
      halobridge_parent_doc_id: parentDocId,
      halobridge_lineage_reason: lineageReason
    };

    return {
      action,
      halobridge_doc_id: docId,
      halobridge_parent_doc_id: parentDocId,
      halobridge_lineage_reason: lineageReason,
      halobridge_original_filename: responseMetadata.halobridge_original_filename,
      halobridge_last_synced_filename: responseMetadata.halobridge_last_synced_filename,
      halobridge_file_fingerprint: responseMetadata.halobridge_file_fingerprint,
      halobridge_version_number: responseMetadata.halobridge_version_number,
      halobridge_created_at: responseMetadata.halobridge_created_at,
      halobridge_last_synced_at: responseMetadata.halobridge_last_synced_at,
      write_back_metadata: this.filterWriteBackMetadata(responseMetadata)
    };
  }

  private createNewDocument(tenant_id: string, filename: string, stage: string, reason: WordConnectorLineageReason, parentId?: string) {
    const doc = {
      id: uuidv4(),
      tenant_id,
      original_filename: filename,
      created_at: new Date().toISOString(),
      metadata: {
        halobridge_parent_doc_id: parentId,
        halobridge_lineage_reason: reason
      }
    };
    this.documents.push(doc);
    this.addLifecycleStage(doc.id, stage);
    return doc;
  }

  private createVersion(docId: string, tenantId: string, userId: string, filename: string, fileHash: string, contentHash: string) {
    const docVersions = this.versions.filter(v => v.document_id === docId);
    const versionNumber = docVersions.length + 1;
    
    const version = {
      id: uuidv4(),
      document_id: docId,
      tenant_id: tenantId,
      user_id: userId,
      version_number: versionNumber,
      file_hash: fileHash,
      content_hash: contentHash,
      filename,
      created_at: new Date().toISOString()
    };
    this.versions.push(version);
    return versionNumber;
  }

  private getLatestContentHash(docId: string): string | null {
    const docVersions = this.versions.filter(v => v.document_id === docId);
    if (docVersions.length === 0) return null;
    return docVersions.sort((a, b) => b.version_number - a.version_number)[0].content_hash;
  }

  private linkRelationship(sourceId: string, targetId: string, type: string, reason: string) {
    this.relationships.push({
      id: uuidv4(),
      source_id: sourceId,
      target_id: targetId,
      relationship_type: type,
      metadata: { reason },
      created_at: new Date().toISOString()
    });
  }

  private addLifecycleStage(docId: string, stage: string) {
    this.lifecycleStages.push({
      id: uuidv4(),
      document_id: docId,
      stage,
      timestamp: new Date().toISOString()
    });
  }

  private filterWriteBackMetadata(data: WordConnectorMetadata): Partial<WordConnectorMetadata> {
    // Only return minimum identity + lineage fields
    const allowlist: (keyof WordConnectorMetadata)[] = [
      'halobridge_doc_id',
      'halobridge_tenant_id',
      'halobridge_source_app',
      'halobridge_original_filename',
      'halobridge_last_synced_filename',
      'halobridge_file_fingerprint',
      'halobridge_version_number',
      'halobridge_created_at',
      'halobridge_last_synced_at',
      'halobridge_parent_doc_id',
      'halobridge_lineage_reason'
    ];

    const result: any = {};
    for (const key of allowlist) {
      if (data[key] !== undefined) {
        result[key] = data[key];
      }
    }
    return result as Partial<WordConnectorMetadata>;
  }
}
