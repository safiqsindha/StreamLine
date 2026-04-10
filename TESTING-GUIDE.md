# Streamline Testing Guide

A structured set of test scenarios that exercise every shipping feature of Streamline. Use this immediately after sideloading to validate the install, and before every demo to catch regressions.

**How to use:** Work top-to-bottom. Each section is independent — you can skip sections that aren't relevant to your demo. Tick the box next to each step as you verify it.

**Setup:** Sideload the add-in following [SIDELOAD-WINDOWS.md](SIDELOAD-WINDOWS.md) or [SIDELOAD-MAC.md](SIDELOAD-MAC.md) first.

---

## Section 0 — Smoke test (do this first, every time)

- [ ] **0.1** PowerPoint launches without errors
- [ ] **0.2** "Streamline" group is visible on the **Home** ribbon
- [ ] **0.3** All four buttons render with icons: **Streamline**, **Refresh**, **Today Line**, **Export PNG**
- [ ] **0.4** Click **Streamline** — task pane opens on the right within ~2 seconds
- [ ] **0.5** Task pane shows the Streamline header, no JavaScript errors visible
- [ ] **0.6** Open browser DevTools (right-click in task pane → Inspect Element) — Console tab shows no red errors

If any step fails, see "Smoke test failures" at the bottom of this guide.

---

## Section 1 — Excel import (the most common workflow)

This is the most heavily exercised path; if it works, ~70% of users will be happy.

### Prerequisites
A sample Excel file with columns: `Swim Lane`, `Task Name`, `Type`, `Start Date`, `End Date`, `% Complete`, `Status`, `Owner`, `Dependency`. Generate one with `npm run test:data` from the repo root — it produces `test/sample-project.xlsx`.

### Steps

- [ ] **1.1** In the task pane, click **Import from Excel** (or the equivalent button — verify the label)
- [ ] **1.2** File picker opens; select `test/sample-project.xlsx`
- [ ] **1.3** Rows preview in the task pane; verify swim lane count, task count, and date range match the file
- [ ] **1.4** Click **Generate Gantt** (or equivalent)
- [ ] **1.5** Within ~3 seconds, the active slide fills with a Gantt chart
- [ ] **1.6** Verify visually: swim lanes labeled on the left, time axis at top, bars correctly positioned, milestones rendered as diamonds, dependencies drawn as arrows
- [ ] **1.7** Click any task bar → it should be selectable as a native PowerPoint shape (not opaque)
- [ ] **1.8** Right-click a task bar → standard PowerPoint shape context menu appears

**Pass criteria:** chart renders without errors, all elements are addressable as PPT shapes.

---

## Section 2 — Other import paths

### 2A. Clipboard paste

- [ ] **2A.1** Open the sample Excel file in Excel
- [ ] **2A.2** Select all rows with headers, copy (Ctrl/Cmd+C)
- [ ] **2A.3** In Streamline task pane, click **Paste from Clipboard** (or equivalent)
- [ ] **2A.4** Rows preview as in 1.3
- [ ] **2A.5** Generate — chart renders identically to 1.5

### 2B. MS Project XML import

- [ ] **2B.1** In MS Project, save a project as `.xml`. Or use any sample MSP XML file.
- [ ] **2B.2** Click **Import from MS Project** in the task pane
- [ ] **2B.3** Select the `.xml` file
- [ ] **2B.4** Tasks, dependencies, and resource owners populate the preview
- [ ] **2B.5** Generate — chart renders with MSP's task hierarchy as swim lanes

### 2C. Manual entry

- [ ] **2C.1** Click **Add Task Manually** (or equivalent — open the data editor)
- [ ] **2C.2** Add 3 swim lanes
- [ ] **2C.3** Add 2 tasks per lane with start/end dates
- [ ] **2C.4** Add 1 milestone
- [ ] **2C.5** Add a dependency (Task B depends on Task A)
- [ ] **2C.6** Click Generate
- [ ] **2C.7** Chart renders with 6 bars, 1 milestone, 1 dependency arrow

### 2D. Natural language entry

- [ ] **2D.1** Click **Quick Add** (or natural language input)
- [ ] **2D.2** Type: `Alpha release May 3 to May 20, owner Kim`
- [ ] **2D.3** Verify it parses and adds the task with correct dates
- [ ] **2D.4** Type: `Launch milestone June 15`
- [ ] **2D.5** Verify it adds as a milestone

### 2E. .mpp binary detection (negative test)

- [ ] **2E.1** Try to import a `.mpp` binary file
- [ ] **2E.2** Streamline detects the binary format and shows a friendly error: "MS Project binary files aren't supported yet — save as XML in Project (File → Save As → XML) and import that file instead."

---

## Section 3 — Visualization elements

After running an import (use sample-project.xlsx), verify each visual element:

### 3A. Bars
- [ ] **3A.1** Each task has a horizontal bar
- [ ] **3A.2** Bar length is proportional to duration
- [ ] **3A.3** Bar fill color matches the active template
- [ ] **3A.4** Bars labeled with task name (default position: inside)

### 3B. Milestones
- [ ] **3B.1** Milestones render as diamonds (or whatever shape the active template specifies)
- [ ] **3B.2** Milestone label appears next to the diamond
- [ ] **3B.3** Multi-day "milestones" still render as a single diamond at the start date

### 3C. Swim lanes
- [ ] **3C.1** Lane label appears on the left edge
- [ ] **3C.2** Lane has visible background tint or border
- [ ] **3C.3** Tasks group inside their lane with no overlap into other lanes

### 3D. Dependency arrows
- [ ] **3D.1** Arrows draw from predecessor to successor
- [ ] **3D.2** Arrow head clearly visible at successor end
- [ ] **3D.3** Different dependency types (FS, SS, FF, SF) visually distinguishable if your test data uses them

### 3E. Today marker
- [ ] **3E.1** Click the **Today Line** ribbon button
- [ ] **3E.2** A vertical line draws at today's date with a "Today" label
- [ ] **3E.3** Click the button again — the line disappears

### 3F. Timescale bands
- [ ] **3F.1** Top of chart shows year/quarter/month tier
- [ ] **3F.2** Tier alignment is correct (Jan in Q1, etc.)
- [ ] **3F.3** Switch fiscal year mode in settings — labels update accordingly

### 3G. Critical path
- [ ] **3G.1** Open settings → toggle "Show Critical Path"
- [ ] **3G.2** Critical path tasks highlight (color, border, or stroke change)
- [ ] **3G.3** Toggle off — highlighting removed

### 3H. Weekend shading
- [ ] **3H.1** Open settings → enable weekend highlighting
- [ ] **3H.2** Saturdays and Sundays show as a tinted vertical band
- [ ] **3H.3** Switch working day preset to Sun-Thu (Middle East) — Friday and Saturday now shaded
- [ ] **3H.4** Switch back to Mon-Fri

### 3I. Percent complete
- [ ] **3I.1** Tasks with `% Complete > 0` show a partial fill
- [ ] **3I.2** Fill width matches the percentage (50% = half-filled bar)
- [ ] **3I.3** 100% complete tasks render fully filled

---

## Section 4 — Editing & interaction

### 4A. Drag to reschedule
- [ ] **4A.1** Click and drag a task bar horizontally
- [ ] **4A.2** Bar moves to the new date range
- [ ] **4A.3** Dependent tasks cascade automatically
- [ ] **4A.4** Critical path recalculates

### 4B. Drag endpoints to resize
- [ ] **4B.1** Hover near the right edge of a bar — resize cursor appears
- [ ] **4B.2** Drag to extend duration
- [ ] **4B.3** End date updates in the data preview

### 4C. Inline editing
- [ ] **4C.1** Open the data editor in the task pane
- [ ] **4C.2** Change a task name; press Enter
- [ ] **4C.3** Click Refresh in the ribbon — chart updates to reflect the new name
- [ ] **4C.4** Verify shape tags (`task_001`, etc.) preserved any manual position tweaks

### 4D. Refresh preserves manual edits
- [ ] **4D.1** Manually drag a task bar to a slightly different position on the slide (without changing dates)
- [ ] **4D.2** Click **Refresh** in the ribbon
- [ ] **4D.3** The bar should keep its manual position; only data changes propagate

### 4E. Cascading auto-schedule
- [ ] **4E.1** With dependencies set up, change a predecessor's end date
- [ ] **4E.2** Successor tasks automatically shift
- [ ] **4E.3** Cascade follows the full DAG (test with 3+ chained dependencies)

---

## Section 5 — Keyboard shortcuts

Test all 16 documented shortcuts (full list in `src/ui/keyboardShortcuts.js`). Spot-check at minimum:

- [ ] **5.1** **Ctrl/Cmd+Z** — undo last action
- [ ] **5.2** **Ctrl/Cmd+Shift+Z** — redo
- [ ] **5.3** Add task shortcut — opens the add task dialog or inserts a row
- [ ] **5.4** Add milestone shortcut — same for milestones
- [ ] **5.5** Delete shortcut — removes the selected task
- [ ] **5.6** Today shortcut — toggles the today line (mirrors ribbon button)
- [ ] **5.7** Refresh shortcut — re-renders chart
- [ ] **5.8** Export shortcut — opens the export dialog or directly exports

If any shortcut conflicts with PowerPoint's own bindings, document it as a known issue — Office.js shortcut handling has limitations on certain hosts.

---

## Section 6 — Templates & styling

### 6A. Template switching
- [ ] **6A.1** Open template picker in the task pane
- [ ] **6A.2** Verify at least 10 templates listed across 5 categories
- [ ] **6A.3** Click a corporate template — colors, fonts, and bar styles update
- [ ] **6A.4** Click a high-contrast template — verify accessibility-friendly colors
- [ ] **6A.5** Click back to "Standard" — reverts cleanly

### 6B. Per-element text styling
- [ ] **6B.1** Open text style settings
- [ ] **6B.2** Verify 11 element types listed (lane label, task name, milestone label, date label, percent complete, owner label, today label, timescale year, timescale quarter, timescale month, project title — confirm exact list against `src/core/templateManager.js`)
- [ ] **6B.3** For each element, verify bold/italic/underline toggles
- [ ] **6B.4** Apply bold to "Task Name" only — only task names go bold; other text unchanged
- [ ] **6B.5** Refresh — styles persist

### 6C. Color slots
- [ ] **6C.1** Open color settings
- [ ] **6C.2** Verify 25+ color slots
- [ ] **6C.3** Change one color (e.g., on-track bar fill) — refresh — color updates only that element

### 6D. Task label positions
- [ ] **6D.1** Open label position setting
- [ ] **6D.2** Verify 5 options: inside, above, below, left, right
- [ ] **6D.3** Switch through each — labels reposition accordingly
- [ ] **6D.4** Verify collision avoidance: dense charts shouldn't overlap labels

---

## Section 7 — Export

### 7A. PNG export
- [ ] **7A.1** Click **Export PNG** ribbon button
- [ ] **7A.2** File save dialog opens
- [ ] **7A.3** Save and open the file — image is rendered correctly
- [ ] **7A.4** Resolution should be high enough for screen display (≥1920×1080 typical)

### 7B. JPG export
- [ ] **7B.1** Use the export menu in the task pane
- [ ] **7B.2** Choose JPG
- [ ] **7B.3** Verify file opens correctly

### 7C. PDF export
- [ ] **7C.1** Choose PDF from export menu
- [ ] **7C.2** Verify multipage handling if chart spans multiple pages
- [ ] **7C.3** Verify text remains selectable (not flattened to image)

### 7D. MS Project XML export
- [ ] **7D.1** Choose MS Project XML
- [ ] **7D.2** Open the resulting `.xml` file in Microsoft Project (or any XML viewer)
- [ ] **7D.3** Verify all tasks, dependencies, and dates round-trip correctly
- [ ] **7D.4** **Bidirectional test:** import the exported XML back into Streamline — verify chart renders identically

---

## Section 8 — Microsoft 365 integration (requires backend)

Skip this section unless you've completed Part C of the sideload guide.

### 8A. Sign-in
- [ ] **8A.1** Click **Sign in to Microsoft 365** in the task pane
- [ ] **8A.2** Office SSO prompt appears (or silent if already signed in)
- [ ] **8A.3** First-time only: consent dialog lists the requested Graph scopes
- [ ] **8A.4** After consent, your displayName appears in the task pane
- [ ] **8A.5** **Sign out** button now visible

### 8B. Backend healthcheck
- [ ] **8B.1** In a terminal: `curl -k https://localhost:3001/health` returns `{"ok":true,...}`
- [ ] **8B.2** Backend logs (in its terminal) show a `[graph] OBO exchange...` line for the sign-in

### 8C. Planner import
- [ ] **8C.1** Source picker → choose **Planner**
- [ ] **8C.2** A list of your Planner plans appears
- [ ] **8C.3** Select one
- [ ] **8C.4** Click **Import**
- [ ] **8C.5** Tasks pull in; buckets become swim lanes; tasks with only a due date become milestones
- [ ] **8C.6** Generate — chart renders

### 8D. To Do import
- [ ] **8D.1** Source picker → **To Do**
- [ ] **8D.2** Pick a list
- [ ] **8D.3** Tasks import; status mapping verified (completed → Complete, waiting → At Risk)

### 8E. Outlook Calendar import
- [ ] **8E.1** Source picker → **Calendar**
- [ ] **8E.2** Date range selector visible
- [ ] **8E.3** Pick "next 30 days" → import
- [ ] **8E.4** Events render as tasks; ≤1 day events render as milestones

### 8F. SharePoint list import
- [ ] **8F.1** Source picker → **SharePoint**
- [ ] **8F.2** Select a site → list
- [ ] **8F.3** Default field mapping works for standard list templates
- [ ] **8F.4** Items import as tasks

### 8G. OneDrive file import
- [ ] **8G.1** Source picker → **OneDrive**
- [ ] **8G.2** File browser opens
- [ ] **8G.3** Navigate to an Excel or MSP XML file
- [ ] **8G.4** File downloads and imports as if it were a local file

### 8H. Sign out
- [ ] **8H.1** Click **Sign out**
- [ ] **8H.2** User name disappears
- [ ] **8H.3** Source picker disables
- [ ] **8H.4** No leftover token in memory (you can't easily verify this without DevTools — open console and check `window.streamline?.graphClient?.accessToken` is null after sign-out)

---

## Section 9 — Copilot integration (requires backend + Copilot studio setup)

Skip this section unless you've uploaded the declarative agent in `copilot-package/declarativeAgent.json` to Copilot Studio for your tenant.

### 9A. Copilot prompt → Gantt
- [ ] **9A.1** Open Copilot in PowerPoint (Home ribbon → Copilot)
- [ ] **9A.2** Type: "Use Streamline to build a 6-week launch plan for project Orion"
- [ ] **9A.3** Copilot calls `/api/copilot/createGantt` on your backend
- [ ] **9A.4** Backend returns a render-pending envelope
- [ ] **9A.5** Streamline task pane (if open) picks up the request and renders the chart
- [ ] **9A.6** If task pane isn't open, Copilot displays a summary card

### 9B. Update via Copilot
- [ ] **9B.1** With a chart on the slide, type to Copilot: "Mark Alpha release at risk"
- [ ] **9B.2** `/api/copilot/updateTasks` is called
- [ ] **9B.3** The Alpha release bar's status indicator changes (color shift to At Risk)

### 9C. Describe via Copilot
- [ ] **9C.1** Type: "Summarize the current Gantt chart"
- [ ] **9C.2** Copilot returns: project name, lane count, task count, milestone count, critical path length, at-risk tasks
- [ ] **9C.3** Numbers match what's actually on the slide

### 9D. Import via Copilot
- [ ] **9D.1** Type: "Import my Engineering Planner board into a Gantt chart"
- [ ] **9D.2** Copilot calls `/api/copilot/importFromM365` with `source=planner`
- [ ] **9D.3** Backend fetches via Graph; chart renders

---

## Section 10 — Teams message extension (optional)

Skip unless you've sideloaded the Teams package.

### 10A. Search Gantts
- [ ] **10A.1** In Teams chat, click the Streamline app in the message extension menu
- [ ] **10A.2** Type a search query
- [ ] **10A.3** Cards return matching Gantts (or a stub if no shared store yet)

### 10B. Create Gantt from message
- [ ] **10B.1** In Teams chat, find a message with task-like content (`## Sprint\nTask A - 2026-04-15 to 2026-04-30 (Alice)`)
- [ ] **10B.2** Click ⋯ on the message → **Streamline → Create Gantt from message**
- [ ] **10B.3** A preview card appears showing extracted tasks
- [ ] **10B.4** Click **Open in PowerPoint** → PowerPoint launches with the Gantt rendered

### 10C. Summarize Gantt
- [ ] **10C.1** From the message extension, choose **Summarize** with a Gantt ID
- [ ] **10C.2** A summary card returns

---

## Section 11 — Cross-platform parity (the demo killer)

Run this section on **both** Windows and Mac to prove the platform-reach claim. The whole point of the strategic positioning is that Streamline runs everywhere Office Timeline can't.

- [ ] **11.1** All of Section 0–7 passes on **Windows desktop**
- [ ] **11.2** All of Section 0–7 passes on **Mac desktop**
- [ ] **11.3** Section 0–4 passes on **PowerPoint for the web** (visit office.com → PowerPoint → New blank → sideload via the cloud-based admin centralized deployment, or via "Get Add-ins" → My Add-ins → upload manifest)
- [ ] **11.4** Section 0–3 passes on **PowerPoint for iPad** (sideload via M365 admin centralized deployment)

If any platform fails on Section 0 (smoke test), the platform claim is broken — fix before any demo.

---

## Smoke test failures (Section 0 troubleshooting)

| Symptom | Most likely cause | Fix |
|---|---|---|
| Streamline ribbon group missing | Manifest sideload didn't take | Re-sideload using the appropriate guide |
| Streamline ribbon group present but task pane empty | Webpack dev server not running | Check terminal — should see "Project is running at https://localhost:3000" |
| Task pane shows "Couldn't open this add-in" | Manifest validation error | `npm run validate` from repo root, fix any errors, re-sideload |
| Task pane shows white screen with "Not Secure" cert warning behind it | Dev cert not trusted | `npx office-addin-dev-certs install` again, or trust the cert manually in the OS keychain/cert store |
| Task pane shows JS errors in DevTools console | Build artifact stale | `npm run build` from repo root, restart `npm start` |
| Excel import button does nothing | sheetjs failed to load | Check browser console for module errors; verify `node_modules/xlsx` exists |
| Drag-to-reschedule doesn't work | Office.js interaction event handlers not wiring up | Check browser console; usually means stale build — `npm run build && npm start` |

---

## Test data generation

To regenerate the sample Excel file:

```bash
npm run test:data
```

This produces `test/sample-project.xlsx` with realistic data covering all task types, all dependency types, milestones, and at-risk tasks. Use it for every test run so you have consistent, repeatable data.

---

## Pre-demo checklist (do this in the 5 minutes before showing the product)

- [ ] **D.1** Restart PowerPoint clean (no stale state)
- [ ] **D.2** Run Section 0 smoke test — must pass 100%
- [ ] **D.3** Open `test/sample-project.xlsx` in Excel and verify it opens correctly
- [ ] **D.4** Run Section 1 (Excel import) end-to-end
- [ ] **D.5** Verify dev server and (if relevant) backend are still running with no errors
- [ ] **D.6** Have a fresh blank PowerPoint deck ready
- [ ] **D.7** Browser DevTools closed (don't accidentally show internals during a demo)
- [ ] **D.8** Backup: a pre-rendered Streamline slide saved as a .pptx file you can fall back to if live rendering fails

If all D-items pass, you are demo-ready.
