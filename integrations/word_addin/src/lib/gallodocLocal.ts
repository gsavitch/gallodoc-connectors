import { v4 as uuidv4 } from 'uuid';

export interface GalloDocLocal {
  schema: string;
  profile: string;
  metadata: {
    generated_at: string;
    source_app: string;
    source_connector: string;
    connector_mode: string;
    document_title?: string;
  };
  content: {
    text: string;
    ooxml_available: boolean;
  };
  lifecycle: Array<{
    event: string;
    timestamp: string;
  }>;
}

export async function generateLocalGalloDoc(text: string, ooxmlAvailable: boolean, title?: string): Promise<GalloDocLocal> {
  return {
    schema: "gallodoc.document.v1",
    profile: "gallodoc/legal/general_word_document/v1",
    metadata: {
      generated_at: new Date().toISOString(),
      source_app: "microsoft_word",
      source_connector: "halobridge_word_addin",
      connector_mode: "local_open_source",
      document_title: title
    },
    content: {
      text: text,
      ooxml_available: ooxmlAvailable
    },
    lifecycle: [
      {
        event: "local_generated",
        timestamp: new Date().toISOString()
      }
    ]
  };
}
