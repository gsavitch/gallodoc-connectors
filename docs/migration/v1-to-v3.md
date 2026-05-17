# GalloDoc Migration Guide: v1 to v3

## Overview
Moving from GalloDoc v1 to v3 is a major step forward in operational intelligence. This guide outlines the adoption path and compatibility guarantees.

## What's New
- **V2.0/2.1 Backfill Features**: Incremental improvements to connector stability and schema flexibility.
- **V3 Native Features**: AI/BI planning and deep semantic linking.

## Compatibility Guarantees
- **6-Month Window**: v1 schemas will be supported in the native `migrate_v1_to_v3` helper for 180 days.
- **Round-Trip Safety**: v1 documents converted to v3 can be losslessly downgraded back to v1 (minus v3-only metadata).

## Adoption Path
1. **Inventory**: Identify v1 documents.
2. **Convert**: Use `gallodoc-cli migrate`.
3. **Verify**: Run the stability suite.
4. **Augment**: Add `embedding_profile_refs` to take advantage of v3 linking.
