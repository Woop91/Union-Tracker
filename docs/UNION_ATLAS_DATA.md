# Union Atlas data bridge

The Union Atlas prototype in `docs/union-atlas.html` uses the repository-local
`data/union-atlas.json` package. That package includes every dataset used by the
prototype plus the safe USUnions aggregate and schema import.

## Provenance

- Local package: `data/union-atlas.json`
- Imported from: `Woop91/SolidBase`
- Product name: SolidBase / USUnions
- Snapshot mode: seeded repository demo data
- Production status: not production and not synchronized

The prototype has no runtime network, spreadsheet, Apps Script, or source-repo
dependency. It can be opened directly from disk.

## Imported

- Aggregate engagement rates and counts
- Six monthly aggregate membership totals
- Aggregate steward workload indicators
- A mapping contract for public local-union identity and geography fields

## Excluded

- Member or roster rows
- Grievance, discipline, or case records
- Email addresses, phone numbers, and private contacts
- Street addresses, steward names, session tokens, and authentication data

The imported membership total is deliberately shown in a separate data view.
It is not added to Atlas organization counts or the invented platform-group
counts because the populations may overlap.

## Live adapter boundary

A future authenticated adapter can call these existing SolidBase functions:

- `dataGetEngagementStats(sessionToken)`
- `dataGetWorkloadSummaryStats(sessionToken)`
- `dataGetMemberCount(sessionToken)`

Only aggregate results may enter the Atlas. `CONFIG_HEADER_MAP_` can map
organization name, abbreviation, local number, parent union, state/region,
public website, and a public office city/state. Placeholder configuration
values are not factual organization data and must never be published.
