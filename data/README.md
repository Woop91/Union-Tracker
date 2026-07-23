# Standalone Union Atlas data

`union-atlas.json` is the complete local data package used by the Common Ground
Union Atlas prototype. The repository does not need the original collection
session, a spreadsheet, Google Apps Script, or a network request to render the
prototype.

## Included

- 45 public labor organizations and connective bodies
- 14 fictional/demo platform groups
- Sector, issue, and organization-type taxonomies
- Network suggestions, mutual-aid activity, joint action, and alliance data
- Admin distributions and Atlas coverage gaps
- Complete safe USUnions/SolidBase aggregate demo snapshot
- Public local-union identity/geography import schema
- Provenance and privacy metadata

## Excluded

- Member rosters and personal identifiers
- Grievance, discipline, and case-level records
- Email addresses, phone numbers, and private contacts
- Street addresses, steward identities, session tokens, and authentication data

These exclusions are intentional. They are application records, not Union Atlas
directory data. Importing them would make the Atlas unsafe, not more complete.

## Build and verification

```text
npm run data:build
npm run data:verify
```

`data:build` creates `manifest.json` with the data-file SHA-256 and record
counts. `data:verify` checks referential integrity, source independence, privacy
boundaries, and parity with the standalone HTML.
