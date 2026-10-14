# Streamline

Streamline imports hardware program schedules from Excel or Microsoft Project, renders them as a grid and timeline, and lets program managers adjust how they are presented. The imported file is the authority: Streamline never edits a source date, never calculates its own critical path, and shows `Unavailable` wherever the file is silent.

## Quick start

Requires Node 22.

```bash
npm install
npm start          # builds the web app and serves it with the API at http://127.0.0.1:4317
```

For development, with live reload:

```bash
npm run dev        # API on 127.0.0.1:4317, web app on http://127.0.0.1:5173
```

The service binds to loopback only and runs until you stop it. Schedules are stored as JSON under `./data` (set `STREAMLINE_DATA` to change it; `STREAMLINE_PORT` changes the port).

## What it does

- **Import** `.xlsx` (column mapping confirmed on a preview) and Microsoft Project MSPDI `.xml`. Nothing is saved until you accept the preview; every skipped row or unreadable value is listed in plain language.
- **Five views** from one canonical schedule: Executive, Detailed, Dependencies, Workstreams and Critical path.
- **Presentation editing only**: display labels, display groups, ordering and emphasis. Source fields are read-only.
- **Re-import** reconciles by task ID and shows added, removed and changed tasks before you accept. Presentation settings that no longer match a task are kept.
- **Baselines**: immutable snapshots of source dates, with compare bars and finish variance.
- **Exports**: PNG, SVG and PDF of the rendered view; CSV and JSON of the canonical data.
- **API parity**: every UI action is an HTTP endpoint. The API screen in the app is rendered from `GET /api/openapi.json`.

## Data rules

- Criticality is `critical`, `not-critical` or `unavailable`, exactly as the file states it. The Critical path view shows an explanation, not a computed path, when the file has no criticality.
- Float is shown only when the file provides it.
- Dependencies keep their stated type (FS, SS, FF, SF) and lag. A link without a stated type is drawn neutrally and labelled unknown.
- Downstream highlighting follows stated dependency links. It is reachability, not schedule analysis.

## Project layout

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for the module map and the rules every change is checked against, and [docs/EXTENSION-POINTS.md](docs/EXTENSION-POINTS.md) for how deferred features plug in.

```bash
npm test           # unit and API tests
npm run typecheck
```

## Re-skinning

All colors, fonts, sizes and radii are CSS custom properties in [`src/web/theme.css`](src/web/theme.css). Exports read the same file, so a replacement theme changes the screen and every exported image together. All user-facing text is in [`src/core/strings.ts`](src/core/strings.ts).

## Browser support

Target: Microsoft Edge, current stable channel. The exact qualified version is to be confirmed and recorded here before v1 is called qualified.

Text uses Segoe UI where installed (all supported Windows versions) and falls back to the system UI font. Cascadia Mono is bundled for dates and numbers.

## To be filled from internal standards

- Repository and branching: `[INTERNAL REPO / BRANCHING]`
- Authentication: `[AUTH MODEL]`
- Deployment: `[DEPLOYMENT TARGET]`
- Telemetry: `[TELEMETRY POLICY]`
- Accessibility conformance: `[ACCESSIBILITY CONFORMANCE TARGET]`
