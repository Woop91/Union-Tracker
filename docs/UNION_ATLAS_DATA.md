# Union Atlas data provenance

The Atlas uses a repository-local public filing directory:

- Canonical source: `Woop91/USUnions`
- Source grain: one canonical OLMS union record
- Record count: 20,699
- Geography: public headquarters city/state/ZIP from OLMS filings
- Coordinates: 2025 U.S. Census Gazetteer ZCTA representative points
- Local package: `data/union-atlas.json`
- Browser package: `docs/union-atlas-data.js`

The exact source commit, source-manifest SHA-256, row coverage, coordinate
source, and package checksum are stored in `data/manifest.json`.

## Display rule

The Atlas displays only source fields or values computed directly from source
rows. Missing values are shown as not reported. It does not invent platform
groups, campaigns, relationships, activities, membership trends, or engagement
metrics.

## Privacy boundary

The package contains organization-level public filing data. It excludes member
rosters, grievances, cases, authentication records, private contacts, personal
identifiers, and street addresses.

## Importing a new source snapshot

Run the importer against an authenticated local checkout of the private
USUnions repository and the official Census ZCTA Gazetteer text file:

```text
node scripts/import-union-atlas-data.mjs --source <USUnions checkout> --zcta <Gazetteer text file>
npm run data:build
npm run data:verify
```
