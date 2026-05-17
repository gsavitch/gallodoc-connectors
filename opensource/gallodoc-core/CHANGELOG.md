# Changelog

All notable changes to GalloDoc Core will be documented in this file.

## [3.0.0] - 2026-05-17 (Beta Release)
### Added
- **AI/BI Planner**: Native support for relationship query execution receipts.
- **Linker v3**: Embedding-augmented relationship discovery.
- **Lineage Metadata**: Standardized `connector_lineage` block for all v3 documents.
- **GalloUnits**: Enhanced semantic layer for cross-domain reasoning.

### Fixed
- **Stability**: Resolved edge cases in OOXML extraction via the Word Connector.
- **Safety Gate**: Added `release_safety_gate.py` to enforce quality invariants.

### Changed
- **Supersession**: GalloDoc v3 now officially supersedes GalloDoc v1.
- **Classifier**: Development status bumped from Alpha to Beta.
- **Deprecation**: v1 support window set to 6 months.

## [2.1.0] - 2026-02-10
### Added
- **Local Mode**: Initial implementation of privacy-safe local document processing.
- **Custom Properties**: Support for embedding manifests in Word Custom Properties.

## [2.0.0] - 2025-11-20
### Added
- **Word Connector MVP**: First version of the Microsoft Word bridge.
- **HaloBridge Integration**: Direct sync capabilities to the HaloBridge registry.
- **Free Connected Mode**: Authenticated cloud sync for independent users.

## [1.0.0] - 2025-06-15
### Added
- **Initial Release**: Basic document envelope and extraction logic.
