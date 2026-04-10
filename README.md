# Streamline

> A modern, AI-native Gantt chart add-in for Microsoft PowerPoint — built on Office.js, integrated with Microsoft 365, and powered by Copilot.

**Status:** Private, pre-handoff. Being prepared for first-party integration into the Microsoft 365 ecosystem.

---

## What it is

Streamline is a PowerPoint add-in that turns Excel schedules, Microsoft Project files, Planner boards, and natural-language prompts into rich, editable Gantt charts rendered as native PowerPoint shapes. It runs everywhere Office.js runs — Windows, **Mac**, **PowerPoint for the web**, and **iPad** — including platforms where the leading competitor (Office Timeline Expert) cannot ship at all.

### What makes it different

- **Cross-platform** — runs on Windows, Mac, Web, and iPad. Office Timeline is Windows-only.
- **Microsoft 365 native** — first-class import from Planner, To Do, Outlook Calendar, SharePoint, OneDrive via Microsoft Graph + Office SSO.
- **Copilot-integrated** — declarative agent, function commands, and Teams compose extension. Generate, edit, and describe Gantt charts from chat.
- **Native PowerPoint shapes** — every bar, milestone, arrow, and label is a tagged PPT shape, not an opaque container. Animatable, accessible, co-author-friendly.
- **Cascading auto-schedule** — full FS/SS/FF/SF dependency engine with lag/lead support and critical-path highlighting.
- **Multiple export formats** — PNG, JPG, PDF, and bidirectional MS Project XML round-trip.

---

## Quick start

The fastest path from zero to a working sideloaded add-in:

| Platform | Guide |
|---|---|
| **Windows** | [SIDELOAD-WINDOWS.md](SIDELOAD-WINDOWS.md) |
| **Mac** | [SIDELOAD-MAC.md](SIDELOAD-MAC.md) |
| **Testing checklist** | [TESTING-GUIDE.md](TESTING-GUIDE.md) |
| **Backend (M365 / Copilot)** | [backend/README.md](backend/README.md) |

In short:

```bash
npm install
npx office-addin-dev-certs install
npm start          # serves https://localhost:3000
npm run start:desktop   # sideloads + launches PowerPoint
```

---

## Architecture overview

```
┌─────────────────────────────────────────────────────────────┐
│ PowerPoint (Windows / Mac / Web / iPad)                     │
│ ┌─────────────────────────────────────────────────────────┐ │
│ │ Streamline task pane (Office.js)                        │ │
│ │  ├─ Data import (Excel, MPP XML, Clipboard, Manual,     │ │
│ │  │   Natural Language, M365 sources)                    │ │
│ │  ├─ Layout engine + auto-scheduler                      │ │
│ │  ├─ PowerPoint renderer (native shapes, tagged)         │ │
│ │  ├─ Templates + per-element text styles                 │ │
│ │  └─ Export (PNG / JPG / PDF / MS Project XML)           │ │
│ └─────────────────────────────────────────────────────────┘ │
│ ┌─────────────────────────────────────────────────────────┐ │
│ │ Function command runtime (hidden iframe)                │ │
│ │  └─ refreshGantt, toggleTodayMarker, exportPng, ...     │ │
│ └─────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────┘
                              │
                              │ Office SSO + OBO
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ Streamline backend (Express)                                │
│  ├─ /api/graph-token   (OBO exchange)                       │
│  └─ /api/copilot/*     (4 declarative agent endpoints)      │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ Microsoft Entra ID + Microsoft Graph                        │
│  └─ Planner, To Do, Calendar, OneDrive, SharePoint          │
└─────────────────────────────────────────────────────────────┘
```

---

## Repository layout

```
Streamline/
├── src/
│   ├── core/              # Data model, parsers, layout, renderer, scheduler
│   ├── ui/                # Task pane HTML/CSS/JS, keyboard shortcuts
│   └── copilot/           # Function commands, agent actions, message extension
├── backend/               # OBO + Copilot action endpoints (Express)
├── copilot-package/       # Declarative Copilot agent manifest + OpenAPI spec
├── teams-package/         # Teams compose extension manifest
├── test/                  # 512+ test assertions
├── assets/                # Icons, sample data
├── manifest.xml           # Office add-in manifest
├── SIDELOAD-WINDOWS.md    # Windows install guide
├── SIDELOAD-MAC.md        # Mac install guide
├── TESTING-GUIDE.md       # 11-section test checklist
├── SECURITY.md            # Vulnerability disclosure policy
└── LICENSE                # Proprietary
```

---

## Documentation index

| Document | Purpose |
|---|---|
| [SIDELOAD-WINDOWS.md](SIDELOAD-WINDOWS.md) | Clean-machine setup for Windows + PowerPoint Desktop |
| [SIDELOAD-MAC.md](SIDELOAD-MAC.md) | Clean-machine setup for macOS + PowerPoint for Mac |
| [TESTING-GUIDE.md](TESTING-GUIDE.md) | Step-by-step verification of every feature |
| [backend/README.md](backend/README.md) | Entra ID app registration + backend setup |
| [SECURITY.md](SECURITY.md) | Vulnerability disclosure |
| [copilot-package/streamline-actions.json](copilot-package/streamline-actions.json) | OpenAPI spec for the Copilot agent |

---

## Tests

```bash
npm test            # runs all 512 test assertions
npm run validate    # validates manifest.xml
npm audit --omit=dev --audit-level=high   # dependency security check
```

---

## License

This software is **proprietary**. All rights reserved. See [LICENSE](LICENSE).

No license to use, copy, modify, or distribute is granted without prior written permission of the copyright holder.

---

## Contact

For licensing inquiries, partnership discussions, or security reports, see [SECURITY.md](SECURITY.md) or open a private issue / advisory in this repository.
