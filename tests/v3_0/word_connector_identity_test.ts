
import { WordConnectorIdentityService } from "../../src/services/wordConnectorIdentityService";
import { WordConnectorSyncPayload } from "../../src/types/wordConnector";

async function runTests() {
  console.log("Starting Word Connector Identity Service Tests...");

  const documents: any[] = [];
  const versions: any[] = [];
  const relationships: any[] = [];
  const lifecycleStages: any[] = [];

  const service = new WordConnectorIdentityService(documents, versions, relationships, lifecycleStages);

  const tenantId = "tenant_1";
  const userId = "user_1";
  const fileId = "file_identity_1";

  // Test 1: Initial Sync Creates New Document
  console.log("Test 1: Initial Sync...");
  const payload1: WordConnectorSyncPayload = {
    tenant_id: tenantId,
    user_id: userId,
    event_type: 'save',
    current_filename: "test.docx",
    word_file_identity: fileId,
    file_hash: "hash1",
    content_hash: "chash1",
    metadata: {}
  };

  const res1 = service.resolveWordDocumentIdentity(payload1);
  if (res1.action !== 'created') throw new Error("Test 1 failed: expected 'created'");
  if (!res1.halobridge_doc_id) throw new Error("Test 1 failed: missing doc id");
  console.log("Test 1 PASSED");

  const docId1 = res1.halobridge_doc_id;

  // Test 2: Normal Save Updates Same Document
  console.log("Test 2: Normal Save...");
  const payload2: WordConnectorSyncPayload = {
    tenant_id: tenantId,
    user_id: userId,
    event_type: 'save',
    current_filename: "test.docx",
    word_file_identity: fileId,
    embedded_halobridge_doc_id: docId1,
    file_hash: "hash2",
    content_hash: "chash2",
    metadata: res1.write_back_metadata
  };

  const res2 = service.resolveWordDocumentIdentity(payload2);
  if (res2.action !== 'updated') throw new Error("Test 2 failed: expected 'updated'");
  if (res2.halobridge_doc_id !== docId1) throw new Error("Test 2 failed: doc id changed");
  console.log("Test 2 PASSED");

  // Test 3: Save As Creates New Document and Links to Parent
  console.log("Test 3: Save As...");
  const payload3: WordConnectorSyncPayload = {
    tenant_id: tenantId,
    user_id: userId,
    event_type: 'save_as',
    current_filename: "test_copy.docx",
    word_file_identity: "file_identity_2",
    embedded_halobridge_doc_id: docId1,
    file_hash: "hash3",
    content_hash: "chash3",
    metadata: res2.write_back_metadata
  };

  const res3 = service.resolveWordDocumentIdentity(payload3);
  if (res3.action !== 'branched') throw new Error("Test 3 failed: expected 'branched'");
  if (res3.halobridge_doc_id === docId1) throw new Error("Test 3 failed: doc id should be new");
  if (res3.halobridge_parent_doc_id !== docId1) throw new Error("Test 3 failed: missing parent link");
  
  const rel = relationships.find(r => r.source_id === docId1 && r.target_id === res3.halobridge_doc_id);
  if (!rel || rel.relationship_type !== 'save_as_copy') throw new Error("Test 3 failed: missing relationship");
  console.log("Test 3 PASSED");

  // Test 4: Rename Updates Metadata Only
  console.log("Test 4: Rename...");
  const payload4: WordConnectorSyncPayload = {
    tenant_id: tenantId,
    user_id: userId,
    event_type: 'rename',
    current_filename: "renamed_test.docx",
    word_file_identity: fileId,
    previous_word_file_identity: fileId,
    embedded_halobridge_doc_id: docId1,
    file_hash: "hash2",
    content_hash: "chash2",
    metadata: res2.write_back_metadata
  };

  const res4 = service.resolveWordDocumentIdentity(payload4);
  if (res4.action !== 'metadata_updated') throw new Error(`Test 4 failed: expected 'metadata_updated', got ${res4.action}`);
  if (res4.halobridge_doc_id !== docId1) throw new Error("Test 4 failed: doc id changed");
  console.log("Test 4 PASSED");

  // Test 5: Tenant Isolation
  console.log("Test 5: Tenant Isolation...");
  try {
    service.resolveWordDocumentIdentity({
      ...payload2,
      tenant_id: "other_tenant"
    });
    throw new Error("Test 5 failed: should have thrown unauthorized");
  } catch (e: any) {
    if (!e.message.includes("UNAUTHORIZED_ACCESS")) throw e;
    console.log("Test 5 PASSED");
  }

  // Test 6: Metadata Allowlist
  console.log("Test 6: Metadata Allowlist...");
  const forbiddenKeys = ['prompts', 'evidence', 'tokens', 'secrets'];
  const wbMetadata: any = res1.write_back_metadata;
  for (const key of forbiddenKeys) {
    if (wbMetadata[key]) throw new Error(`Test 6 failed: forbidden key ${key} found`);
  }
  console.log("Test 6 PASSED");

  console.log("All Word Connector Identity Tests PASSED!");
}

runTests().catch(err => {
  console.error("Test execution failed:", err);
  process.exit(1);
});
