# Union Atlas data provenance and operation

The Atlas is repository-local and source-backed. It contains:

- 20,699 canonical OLMS union records.
- 23,497 NLRB representation cases, including 18,372 RC organizing attempts.
- 5,714 reported petitioner-group rollups.
- 64 research articles, 97 verified statistics, and 10 research sources.
- 15,243 public headquarters matched by the U.S. Census address geocoder.
- 3,809 additional headquarters using 2025 Census ZCTA representative points.

The exact `Woop91/USUnions` commit, source hashes, package hashes, coverage
denominators, Census sources, and synthetic-record count are stored in
`data/manifest.json`.

## Runtime packages

The hosted Atlas avoids a large main-thread startup:

- `union-atlas-index.js`: small summary and filter index.
- `union-atlas-data.json.gz`: 1.61 MiB initial directory/search package.
- `union-atlas-details.json.gz`: lazy record-detail package.
- `union-atlas-evidence.json.gz`: lazy organizing/research package.
- `union-atlas-worker.js`: decompression, distance calculation, filtering,
  sorting, and pagination outside the main UI thread.
- `union-atlas-standalone.html`: true one-file offline export containing every
  runtime package.

Directory search is debounced. Nearby, directory, organizing, and research
results are paginated without a hidden result cap.

## Geography

Public OLMS headquarters street addresses are sent to the official Census
batch geocoder. The Atlas stores only coordinates, match type, and an address
fingerprint used to detect changes. Street addresses are not stored.

When Census cannot match an address, the Atlas uses the representative point
for the filing ZIP's ZCTA. Each record labels its precision as `address`,
`zcta`, or unavailable.

## Privacy boundary

Included data is organization-level public filing, public NLRB case, and
research evidence. The Atlas excludes:

- Member rosters and personal identifiers.
- Grievance, discipline, and case-management records.
- Private or person-level email addresses and phone numbers.
- Street addresses after geocoding.
- Authentication records, tokens, and credentials.

Missing values stay missing. No seeded or synthetic replacement is allowed.

## Refresh

Local refresh:

```text
npm run data:geocode -- --source <USUnions checkout>
npm run data:import -- --source <USUnions checkout> --zcta <Gazetteer text file>
npm run data:build
npm run data:verify
npm run data:freshness -- --source <USUnions checkout>
```

`.github/workflows/union-atlas-refresh.yml` runs weekly and opens a refresh PR.
Because `Woop91/USUnions` is private, the SolidBase repository must have a
fine-grained read-only Actions secret named `USUNIONS_READ_TOKEN`.
