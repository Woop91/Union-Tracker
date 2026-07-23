# Union Atlas visual system

## Product intent

- **Audience:** union members, organizers, local leaders, and platform administrators.
- **Primary job:** find nearby unions, search canonical records, and inspect source coverage.
- **Signature element:** the north/east/south/west bullseye remains the main geographic orientation device.

## Design source

The visual system comes from Claude session
`05fd27e3-7aa0-4384-934c-1c26921914cf` and the Solid Ground design-system
repository at exact commit `42d54f1619fab56717f4cc2f0b3f4f289aa57b3c`.

The copied source tokens live in `design/solid-ground/`. Run
`npm run atlas:theme` to embed them into the standalone Atlas HTML.

## Applied language

- Forest and pine establish the civic, grounded frame.
- Cream supplies the readable working canvas.
- Gold marks selected navigation, actions, chart emphasis, and connective labor bodies.
- Terracotta is reserved for the viewer's location and destructive or warning semantics.
- Archivo, Public Sans, and Spline Sans Mono roles are preserved through offline-safe system fallbacks.
- Rounded cards, gold top rules, compact utility labels, and restrained motion follow the source recipes.

## Accessibility decisions

The source's decorative muted green is not used for informative copy on cream.
Semantic captions use the darker secondary text token. Light and dark themes
are audited independently with `npm run contrast:audit`. Mobile controls use a
minimum 44-pixel target where they are directly interactive.

## Standalone boundary

The Atlas is an offline-capable local package. Its HTML loads one generated
repository-local data script and no remote fonts, stylesheets, scripts, images,
or design-system runtime.

## Mobile behavior

- The header becomes a compact two-column identity block with safe-area padding.
- Navigation stays horizontally scrollable and keeps the active task centered.
- Search and filter controls become touch-sized responsive grids.
- ZIP plotting, paging, theme, and navigation controls remain at least 44 pixels high.
- Wide source tables scroll inside their cards.
- The compass scales to the available card width while its results move below it.
- Summary and source-coverage metrics collapse from four columns to two, then one.
- Breakpoints are verified at 320, 375, 390, 768, and 1440 pixels.
