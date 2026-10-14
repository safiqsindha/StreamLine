# Streamline architecture

Streamline imports hardware program schedules, renders them, and lets people adjust how they are presented. Source data is read-only and authoritative.

## Layout

```
src/
  core/            Pure TypeScript shared by the API and the web app. No DOM, no Node APIs, no storage.
    model.ts       Canonical schedule, presentation and baseline types
    store.ts       Store interface — the only persistence boundary
    ids.ts         Generated IDs and the deterministic task ID fallback
    dates.ts       UTC day-number date helpers
    layout.ts      The single layout authority (x, y, range, zoom)
    views.ts       View resolution: schedule + presentation + view state → rows
    import/        xlsx and MSPDI parsers, preview, reconciliation
  server/          Node 22 HTTP API, loopback only. Owns the data (FileStore under ./data).
  web/             React + Tailwind app. Talks to the API only; never touches storage.
    theme.css      Every color, font, size and radius as CSS custom properties
    strings.ts     Every user-facing string
test/              Vitest suites and synthetic fixtures
docs/reference/    The visual reference prototype
```

## Rules

These are checked on every change.

1. **Source facts are never invented.** `null` means the file did not provide it; the UI shows `Unavailable`. No defaults, no today's date, no empty strings standing in for values.
2. **No client-side CPM.** Criticality and float come from the source only. `Criticality` is tri-state: `critical | not-critical | unavailable`. Downstream reachability over stated dependencies is allowed and labelled as such.
3. **Source dates are never edited.** No endpoint or control writes a source field. Mutations touch presentation and baselines only.
4. **Source and presentation are separate stores.** Never merge them into one task object. Presentation is keyed by schedule ID.
5. **One layout authority.** Only `core/layout.ts` computes x or y. Grid, timeline and exports consume it.
6. **One ordering comparator** drives grid rows, timeline rows, pagination and keyboard focus. Ties break on task ID.
7. **One persistence boundary.** Only `Store` implementations touch storage. Every read and write is wrapped and fails with a visible message.
8. **IDs are stable and never reused.** Baselines are immutable; saving creates a new record.
9. **No hardcoded colors in components.** Use theme tokens. No inline user-facing copy; use `strings.ts`.
10. **No network calls at runtime** beyond the local API. Fonts and assets are bundled.
11. **Dependencies:** FS, SS, FF, SF are drawn faithfully; `unknown` is drawn neutrally and labelled unknown, never as FS. Lags are preserved.
12. **Code style:** conventional, in-house. No tool-attribution comments or generated-by markers anywhere.

## Task identity

The source task ID is used whenever the file provides one (MSPDI `UID`; a spreadsheet column mapped to `id`). Otherwise `fallbackTaskId` hashes the workstream and task name path, adding an occurrence suffix for exact duplicates in file order. Row position and display label are never identity.

## Running

```
npm install
npm run dev      # API on 127.0.0.1:4317, web on 127.0.0.1:5173
npm test
npm run typecheck
```
