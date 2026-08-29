<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="assets/banner-dark.svg">
    <img src="assets/banner-light.svg" alt="Streamline" width="100%">
  </picture>
</p>

# Streamline

**A modern, AI-native Gantt chart add-in for Microsoft PowerPoint — built on Office.js, integrated with Microsoft 365, and powered by Copilot.**

Streamline turns Excel schedules, Microsoft Project files, Planner boards, and natural-language prompts into rich, editable Gantt charts rendered as **native PowerPoint shapes**. It runs everywhere Office.js runs — Windows, Mac, PowerPoint for the web, and iPad — including the platforms the leading competitor cannot ship to at all.

- **Cross-platform where the incumbent isn't** — Windows, Mac, Web, and iPad. Office Timeline is Windows-only.
- **Native shapes, not an opaque container** — every bar, milestone, arrow, and label is a tagged PowerPoint shape: animatable, accessible, and co-author-friendly
- **Microsoft 365 native** — first-class import from Planner, To Do, Outlook Calendar, SharePoint and OneDrive through Microsoft Graph with Office SSO
- **Copilot-integrated** — a declarative agent, function commands, and a Teams compose extension, so charts can be generated, edited, and described from chat

![Platform](https://img.shields.io/badge/platform-Windows%20%C2%B7%20Mac%20%C2%B7%20Web%20%C2%B7%20iPad-2563EB?style=flat-square)
![Office.js](https://img.shields.io/badge/Office.js-PowerPoint-2563EB?style=flat-square)
![Tests](https://img.shields.io/badge/tests-512%20assertions-22c55e?style=flat-square)
![License](https://img.shields.io/badge/license-proprietary-DC2626?style=flat-square)
![Visibility](https://img.shields.io/badge/visibility-private-6B7280?style=flat-square)

**[Windows setup](SIDELOAD-WINDOWS.md)** · **[Mac setup](SIDELOAD-MAC.md)** · **[Testing guide](TESTING-GUIDE.md)** · **[Backend](backend/README.md)** · **[Security](SECURITY.md)**

```bash
npm install
npx office-addin-dev-certs install
npm run start:desktop      # sideloads + launches PowerPoint
```

> **Status: private, pre-handoff.** Being prepared for first-party integration into the Microsoft 365 ecosystem.

---

## What makes it different

| | Streamline | Office Timeline Expert |
|---|---|---|
| Windows | ✅ | ✅ |
| Mac · Web · iPad | ✅ | ❌ |
| Output | Native tagged PPT shapes | Opaque container |
| M365 import (Planner, To Do, Calendar, SharePoint) | ✅ via Graph + SSO | ❌ |
| Copilot / chat-driven authoring | ✅ | ❌ |

Beyond the comparison: a full **cascading auto-scheduler** with FS/SS/FF/SF dependencies, lag and lead support, and critical-path highlighting; and export to PNG, JPG, PDF, plus a bidirectional MS Project XML round-trip.

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
