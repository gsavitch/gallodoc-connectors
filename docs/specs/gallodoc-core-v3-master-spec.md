# GalloDoc Core v3 Master Specification

**Status:** Final Release (v3.0.0)  
**Read Time:** ~20 minutes

## 1. Introduction
GalloDoc Core is a portable standard for operational intelligence documents. It bridges the gap between raw data extraction and governed, actionable metadata.

## 2. Core Components

### 2.1 Schema (Document Envelope)
The v3 envelope introduces `rev_to_v3` strategy, allowing documents to evolve while maintaining strict lineage tracking.
- Link: [Schema Spec](gallodoc-core-v3-schema.md)

### 2.2 Connectors
Connectors are the ingestion engines of the GalloDoc ecosystem.
- Link: [Connector Guide](../positioning/connector-guide.md)

### 2.3 Semantic Profile (GalloUnits)
GalloUnits provide a domain-specific semantic layer.
- Link: [Semantic Encoder Guide](../positioning/semantic-encoder-guide.md)

### 2.4 Linker & Embeddings
The linker uses trained embedders to generate relationship candidates.
- Link: [Linker Guide](../positioning/linker-guide.md)

### 2.5 AI/BI Planner
The planner interprets complex relationship queries and generates deterministic execution receipts.

## 3. Governance & Privacy
GalloDoc enforces privacy-by-design through `trust_block` enforcement and local-first processing options.
- Link: [Privacy & Governance](../positioning/privacy-and-governance-guide.md)
