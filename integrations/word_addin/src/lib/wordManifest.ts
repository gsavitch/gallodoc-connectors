
import { WordConnectorMetadata } from "../../../../src/types/wordConnector";

/* global Word */

export type GalloDocManifest = WordConnectorMetadata & {
  // We can still keep some transient UI state here if needed, 
  // but for persistence we follow the minimum metadata rule
};

const METADATA_KEYS: (keyof WordConnectorMetadata)[] = [
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

export async function readGalloDocManifest(): Promise<Partial<WordConnectorMetadata> | null> {
  try {
    return await Word.run(async (context) => {
      const customProps = context.document.properties.customProperties;
      const result: any = {};
      let found = false;

      for (const key of METADATA_KEYS) {
        const prop = customProps.getItemOrNullObject(key);
        prop.load("value");
        await context.sync();
        if (!prop.isNullObject && prop.value !== undefined) {
          result[key] = prop.value;
          found = true;
        }
      }

      return found ? result : null;
    });
  } catch (error) {
    console.error("Failed to read GalloDoc manifest:", error);
    return null;
  }
}

export async function writeGalloDocManifest(metadata: Partial<WordConnectorMetadata>): Promise<boolean> {
  try {
    await Word.run(async (context) => {
      const customProps = context.document.properties.customProperties;
      
      for (const [key, value] of Object.entries(metadata)) {
        if (value !== undefined && value !== null) {
          customProps.add(key, value.toString());
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
      for (const key of METADATA_KEYS) {
        const prop = customProps.getItemOrNullObject(key);
        prop.delete();
      }
      await context.sync();
    });
  } catch (error) {
    // Non-critical
  }
}

export function buildManifestFromSyncResponse(
  response: any
): Partial<WordConnectorMetadata> {
  return response.write_back_metadata || {};
}
