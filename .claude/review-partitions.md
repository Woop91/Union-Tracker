# Review Partitions — SolidBase

Same stack as DDS-Dashboard (GAS V8 + Google Sheets) but with org-specific features stripped. Partitions mirror DDS — see `../../DDS-Dashboard/.claude/review-partitions.md` for the full list. SolidBase-specific notes:

- **Exclude from all partitions:** WorkloadService, POMS Reference, agency_org_chart (stripped per sync workflow)
- **Extra focus for R2 "SolidBase sync":** org-specific string hunt, DDS Script ID redaction, config-driven branding
- **Deployments:** SB prod `1V6vzrczxUSYuiobdkKE64mbsZYznZHZwcI51juAtqQojy5Tz8q5zbiTl`. No dev deploys — use prod read-only for UX testing, or set up local GAS dev.

## UX personas

Use the same 14 personas from DDS, but target the SB prod deployment (read-only). Omit any workflow that depends on stripped features (WorkloadService, POMS, agency_org_chart).
