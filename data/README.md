# Source-backed Union Atlas data

`union-atlas.json` is the complete repository-local directory package used by
the Common Ground Union Atlas.

## Included

- 20,699 canonical union records from `Woop91/USUnions`
- Exact source repository commit and source-manifest checksum
- OLMS identity, affiliation, level, public headquarters city/state/ZIP,
  latest reported membership, filing form, public website, and enrichment state
- 2025 U.S. Census Gazetteer ZCTA representative coordinates
- Derived affiliation, state, and completeness rollups

Every displayed organization and metric comes from those source records or is
calculated directly from them. Missing values remain missing.

## Excluded

- Member rosters and personal identifiers
- Grievance, discipline, and case-level records
- Private email addresses or phone numbers
- Street addresses, session tokens, and authentication data

## Build and verification

```text
npm run data:build
npm run data:verify
```

`data:build` creates the manifest, browser data module, and Atlas HTML.
`data:verify` checks source identity, uniqueness, completeness, geography,
privacy boundaries, package parity, and checksums.
