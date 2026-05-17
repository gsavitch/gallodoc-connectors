
import { WordConnectorMetadata } from "../../../../src/types/wordConnector";

/* global Word */

export interface GalloDocManifest extends Partial<WordConnectorMetadata> {
  schema: "halobridge.word_manifest.v1";
  halobridge_doc_id: string | null;
  mvp_document_id: string | null;
  gallodoc_id: string | null;
  latest_version_id: string | null;
  latest_version_number: number | null;
  review_status: string | null;
  approval_status: string | null;
  verifyiq_status: string | null;
  himc_status: string | null;
  release_ready: boolean | null;
  ai_assistance_status: string | null;
  ai_signal_confidence: string | null;
  human_review_required: boolean | null;
  word_control_url: string | null;
  canonical_workspace_url: string | null;
  source_document_id: string | null;
  source_gallodoc_id: string | null;
  last_source_hash: string | null;
  last_synced_at: string | null;
  _source?: string;
}

const MANIFEST_KEY = "gallodoc_manifest";

const FLAT_METADATA_MAP = {
  "HaloBridgeDocumentId": "halobridge_doc_id",
  "GalloDocId": "gallodoc_id",
  "HaloBridgeWorkspaceUrl": "canonical_workspace_url",
  "GalloDocVersion": "latest_version_number",
  "HaloBridgeSourceHash": "last_source_hash",
  "HaloBridgeTenantId": "halobridge_tenant_id",
  "HaloBridgeReviewStatus": "review_status",
  "HaloBridgeApprovalStatus": "approval_status"
};

export async function readGalloDocManifest(): Promise<GalloDocManifest | null> {
  try {
    return await Word.run(async (context) => {
      const customProps = context.document.properties.customProperties;
      const manifestProp = customProps.getItemOrNullObject(MANIFEST_KEY);
      manifestProp.load("value");
      
      // Load flat props for recovery fallback
      const flatProps: Record<string, Word.CustomProperty> = {};
      for (const flatKey of Object.keys(FLAT_METADATA_MAP)) {
        flatProps[flatKey] = customProps.getItemOrNullObject(flatKey);
        flatProps[flatKey].load("value");
      }

      await context.sync();

      if (!manifestProp.isNullObject && manifestProp.value) {
        try {
          return JSON.parse(manifestProp.value);
        } catch (e) {
          console.warn("Invalid manifest JSON, attempting recovery from flat metadata", e);
        }
      }

      // Recovery Fallback from Flat Metadata
      let recovered: any = {
        schema: "halobridge.word_manifest.v1",
        _source: "recovered_from_flat_metadata"
      };
      let foundAny = false;

      for (const [flatKey, manifestKey] of Object.entries(FLAT_METADATA_MAP)) {
        const prop = flatProps[flatKey];
        if (!prop.isNullObject && prop.value !== undefined && prop.value !== null) {
          recovered[manifestKey] = prop.value;
          foundAny = true;
        }
      }

      if (foundAny) {
        console.log("[Identity] Manifest recovered from portable flat metadata.");
        return recovered as GalloDocManifest;
      }

      return null;
    });
  } catch (error) {
    console.error("Failed to read GalloDoc manifest:", error);
    return null;
  }
}

export async function writeGalloDocManifest(manifest: any): Promise<boolean> {
  try {
    await Word.run(async (context) => {
      const customProps = context.document.properties.customProperties;
      
      // Filter out any sensitive data just in case
      const safeManifest = { ...manifest };
      delete (safeManifest as any).token;
      delete (safeManifest as any).password;
      delete (safeManifest as any).apiKey;
      delete (safeManifest as any).clientSecret;

      // 1. Write Structured Manifest
      customProps.add(MANIFEST_KEY, JSON.stringify(safeManifest));

      // 2. Write Flat Properties for Portability
      for (const [flatKey, manifestKey] of Object.entries(FLAT_METADATA_MAP)) {
        const val = safeManifest[manifestKey];
        if (val !== undefined && val !== null) {
          customProps.add(flatKey, val.toString());
        }
      }

      await context.sync();
    });
    return true;
  } catch (error) {
    console.error("Failed to write GalloDoc manifest:", error);
    return false;
  }
}

export async function clearGalloDocManifest(): Promise<void> {
  try {
    await Word.run(async (context) => {
      const customProps = context.document.properties.customProperties;
      
      // Clear Main Manifest
      const prop = customProps.getItemOrNullObject(MANIFEST_KEY);
      prop.delete();

      // Clear Flat Properties
      for (const flatKey of Object.keys(FLAT_METADATA_MAP)) {
        const flatProp = customProps.getItemOrNullObject(flatKey);
        flatProp.delete();
      }

      await context.sync();
    });
  } catch (error) {
    // Non-critical
  }
}

export function buildManifestFromSyncResponse(
  response: any,
  currentManifest: GalloDocManifest | null = null
): GalloDocManifest {
  return {
    schema: "halobridge.word_manifest.v1",
    halobridge_doc_id: response.halobridge_doc_id || response.mvp_document_id || currentManifest?.halobridge_doc_id,
    mvp_document_id: response.mvp_document_id || response.halobridge_doc_id || currentManifest?.mvp_document_id,
    gallodoc_id: response.gallodoc_id || currentManifest?.gallodoc_id,
    latest_version_id: response.version_id || response.latest_version_id || currentManifest?.latest_version_id,
    latest_version_number: response.version_number || response.latest_version_number || currentManifest?.latest_version_number,
    review_status: response.review_status || currentManifest?.review_status || "draft",
    approval_status: response.approval_status || currentManifest?.approval_status || "pending",
    verifyiq_status: response.verifyiq_status || currentManifest?.verifyiq_status || "pending",
    himc_status: response.himc_status || currentManifest?.himc_status || "pending",
    release_ready: response.release_ready ?? currentManifest?.release_ready ?? false,
    ai_assistance_status: response.ai_assistance_status || currentManifest?.ai_assistance_status || "unknown",
    ai_signal_confidence: response.ai_signal_confidence || currentManifest?.ai_signal_confidence || "none",
    human_review_required: response.human_review_required ?? currentManifest?.human_review_required ?? true,
    word_control_url: response.word_control_url || response.canonical_workspace_url || currentManifest?.word_control_url,
    canonical_workspace_url: response.canonical_workspace_url || response.word_control_url || currentManifest?.canonical_workspace_url,
    source_document_id: response.halobridge_parent_doc_id || currentManifest?.source_document_id,
    source_gallodoc_id: currentManifest?.gallodoc_id, // lineage
    last_source_hash: response.halobridge_file_fingerprint || currentManifest?.last_source_hash,
    last_synced_at: response.halobridge_last_synced_at || new Date().toISOString()
  };
}
