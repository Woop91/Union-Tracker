# Solid Ground design template

Imported from `Woop91/solid-ground-design-system` at:

- Commit: `42d54f1619fab56717f4cc2f0b3f4f289aa57b3c`
- Claude design session: `05fd27e3-7aa0-4384-934c-1c26921914cf`
- Imported: 2026-07-23

`tokens.json` is the upstream source of truth. `tokens.css` is its generated
CSS. Run `npm run atlas:theme` to re-inline the exact CSS token file into the
standalone Atlas HTML.

The Atlas maps these brand tokens into semantic light and dark themes. Informative
small text uses `--sg-color-sec`, not upstream `--sg-color-muted`, because the
upstream package documents that muted token as below WCAG AA on canvas.
