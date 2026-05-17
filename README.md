# GalloDoc Connector Hub

**GalloDoc is an open standard for portable operational intelligence documents.**

This repository contains the official GalloDoc Connectors and the GalloDoc Core reference implementation.

## Quick Links
- **[V3 Master Spec](docs/specs/gallodoc-core-v3-master-spec.md)**
- **[Migration Guide](docs/migration/v1-to-v3.md)**
- **[Word Connector](integrations/word_addin/README.md)**

## V3 Master Release
GalloDoc Core v3.0.0 is now in Beta. This release introduces advanced operational intelligence features, AI/BI planning, and an upgraded linker.

## Developer Quick Start
```bash
# Clone the repository
git clone https://github.com/gsavitch/gallodoc-connectors.git
cd gallodoc-connectors

# Run the release safety gate (v3.0.0 requirement)
cd opensource/gallodoc-core
make release-gate
```
