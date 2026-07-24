# Repository-local Union Atlas data

`union-atlas.json` is the complete public-safe directory package. Compressed
source packages under `source/usunions/` retain every real organizing and
research record required to rebuild the Atlas without the source repository.

## Included

- Canonical OLMS identity, hierarchy, sector, headquarters city/state/ZIP.
- Address or ZCTA coordinate and explicit precision.
- Full available membership history.
- Latest filing finances.
- Public website, background, source citations, and organization-level coverage.
- Complete NLRB organizing attempts and petitioner-group rollups.
- Complete research article, source, and verified-statistic datasets.

## Excluded

- Street addresses after Census geocoding.
- Names or contact details for members, staff, stewards, and officers.
- Grievance, discipline, and case-management records.
- Private contacts, session tokens, credentials, and authentication data.

Counts for excluded public-person collections are retained only as coverage
indicators. Missing source values are not filled.

## Generated files

Do not hand-edit `union-atlas.json`, `manifest.json`, the geocode output, source
packages, or files under `docs/union-atlas-*`. Re-run the import and build
commands documented in `docs/UNION_ATLAS_DATA.md`.
