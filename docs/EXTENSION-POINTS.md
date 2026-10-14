# Extension points

Features deferred from v1, and where each one attaches without reworking what exists.

| Deferred feature | Where it attaches |
|---|---|
| Server-side persistence | Implement `Store` (`src/core/store.ts`) against the new backend and pass it to `createApp`. Nothing else reads or writes storage. |
| PowerPoint export | Add a consumer of `renderTimeline` output or of the resolved `Layout`, alongside SVG/PNG/PDF in the export route. It must not compute positions itself. |
| Graph write-back / two-way sync | The sidebar source card keeps a slot under **Re-import** for an approval-gated sync control. Write-back would be a new endpoint family; existing endpoints stay read-only for source fields. |
| Historical replay | Schedules, baselines and imports carry stable generated IDs and creation timestamps; baselines are immutable. Keeping each accepted import as its own record is an additive change in the accept route. |
| Undo | Presentation changes are whole-record writes keyed by schedule ID, so a history of `Presentation` records can be kept by the `Store` without changing callers. |
| Theme editor | Edits `src/web/theme.css` tokens (or an equivalent token set). Components and exports already read only tokens. |
| Auth, multi-user | Add at the HTTP layer in `src/server/main.ts`; the API is already the single entry point for the UI and agents. |
| Federated dependencies, downstream exposure, forecasting | New read-only analyses over the canonical model in `src/core`, exposed as new endpoints. They must label their output as derived, never as source fact. |
