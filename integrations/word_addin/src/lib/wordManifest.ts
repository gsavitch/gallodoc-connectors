/* global Word */

export interface GalloDocManifest {
  gallodoc_schema: string;
  mvp_document_id: string;
  document_name?: string;
  latest_version_id: string;
  latest_version_number: number;
  word_control_url?: string;
  last_synced_at: string;
  last_source_hash: string;
  review_status: string;
  himc_approved_by: string | null;
  himc_approved_at: string | null;
}

const MANIFEST_KEY = "gallodoc_manifest";

export async function readGalloDocManifest(): Promise<GalloDocManifest | null> {
  try {
    return await Word.run(async (context) => {
      const customProps = context.document.properties.customProperties;
      const manifestProp = customProps.getItemOrNullObject(MANIFEST_KEY);
      manifestProp.load("value");
      await context.sync();

      if (manifestProp.isNullObject || !manifestProp.value) {
        return null;
      }

      return JSON.parse(manifestProp.value);
    });
  } catch (error) {
    console.error("Failed to read GalloDoc manifest:", error);
    return null;
  }
}

export async function writeGalloDocManifest(manifest: GalloDocManifest): Promise<boolean> {
  try {
    await Word.run(async (context) => {
      const customProps = context.document.properties.customProperties;
      customProps.add(MANIFEST_KEY, JSON.stringify(manifest));
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
      const prop = customProps.getItemOrNullObject(MANIFEST_KEY);
      prop.delete();
      await context.sync();
    });
  } catch (error) {
    // Non-critical
  }
}

export function buildManifestFromSaveResponse(
  response: any, 
  sourceHash: string, 
  existingManifest?: GalloDocManifest | null,
  documentName?: string
): GalloDocManifest {
  return {
    gallodoc_schema: "gallodoc.word_manifest.v1",
    mvp_document_id: response.document_id || response.gallodoc_id || existingManifest?.mvp_document_id,
    document_name: documentName || response.document_name || existingManifest?.document_name,
    latest_version_id: response.version_id,
    latest_version_number: response.version_number,
    word_control_url: response.processing_url || response.result_url || response.word_control_url,
    last_synced_at: response.last_sync || response.created_at || new Date().toISOString(),
    last_source_hash: sourceHash,
    review_status: response.review_status || response.processing_status || existingManifest?.review_status || "draft",
    himc_approved_by: response.approved_by || existingManifest?.himc_approved_by || null,
    himc_approved_at: response.approved_at || existingManifest?.himc_approved_at || null
  };
}
