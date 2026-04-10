# Sideload & Test Streamline on Mac

This guide walks you from a clean macOS machine to a fully sideloaded Streamline add-in running inside PowerPoint for Mac, with optional backend (M365 / Copilot integration).

**Total time:** 20–30 minutes for the core add-in. Add 30–60 minutes if you also set up the M365 backend.

> **Why Mac matters:** Office Timeline (Streamline's main competitor) does not run on Mac at all. Verifying Streamline on Mac is the single most important demo of the strategic advantage. Don't skip this.

---

## What you need

- **macOS 12 (Monterey) or newer** — Apple Silicon (M1/M2/M3/M4) or Intel both supported
- **PowerPoint for Mac** — Microsoft 365 subscription (the standalone "PowerPoint 2021 for Mac" works too, but M365 is preferred for the latest Office.js APIs)
- **Node.js 18 LTS or newer** — install via [nodejs.org](https://nodejs.org) or via Homebrew: `brew install node`
- **Git** — pre-installed on macOS, or `brew install git`
- **Terminal** — macOS Terminal.app, iTerm2, or Warp all work
- **Admin password** — required only for installing the dev HTTPS certificate (one-time)

---

## Part A — Get the code running

### A.1 Clone the repository

```bash
cd ~/Desktop
git clone <your-streamline-repo-url> Streamline
cd Streamline
```

### A.2 Install dependencies

```bash
npm install
```

If you're on Apple Silicon and see warnings about native rebuilds, ignore them — Streamline has no native deps in the runtime path.

### A.3 Generate and trust the local HTTPS certificate

PowerPoint for Mac requires HTTPS for add-ins, and won't accept untrusted certificates. The `office-addin-dev-certs` tool generates a self-signed cert and installs it into the macOS keychain as a trusted root.

```bash
npx office-addin-dev-certs install
```

You'll be prompted for your macOS admin password (this is the keychain modification prompt). Enter it. The tool prints a success message when done.

Verify the cert is trusted:

```bash
npx office-addin-dev-certs verify
# You should have trusted access to https://localhost
```

If verification fails, open **Keychain Access.app**, search for `localhost`, find the developer cert, double-click it, expand **Trust**, and set "When using this certificate" to **Always Trust**. Close the keychain window (you'll be prompted for your password to save).

### A.4 Build and start the dev server

```bash
npm start
```

You should see webpack output ending with:

```
webpack 5.x.x compiled successfully
[webpack-dev-server] Project is running at:
[webpack-dev-server] Loopback: https://localhost:3000/
```

Leave this terminal window running. Open a **new** Terminal window or tab for the next steps.

### A.5 Sanity check the dev server

In Safari or Chrome, navigate to `https://localhost:3000/taskpane.html`. The Streamline UI should render. If you see "This Connection Is Not Private," the cert isn't trusted — repeat A.3, then click **Show Details** → **visit this website** in Safari to force-trust it for your browser session.

---

## Part B — Sideload the manifest into PowerPoint for Mac

There are two ways to do this on macOS. **B.1** is the easy automated path. **B.2** is the manual fallback when you need fine control.

### B.1 — Easiest: `office-addin-debugging start` (recommended)

The Streamline `package.json` already has this script wired up. From the repo root in a new terminal tab:

```bash
npm run start:desktop
```

What this does:
1. Starts (or reuses) the webpack dev server on https://localhost:3000
2. Copies `manifest.xml` into PowerPoint's `wef` (Web Extension Framework) folder
3. Launches PowerPoint
4. Streamline appears on the **Home** ribbon

When PowerPoint opens, you'll see a **Streamline** group on the **Home** tab with **Streamline**, **Refresh**, **Today Line**, and **Export PNG** buttons. Click **Streamline** to open the task pane on the right.

When you're done:

```bash
npm run stop
```

This unregisters the manifest from PowerPoint's wef folder. Without it, the registration persists across PowerPoint launches (which is usually what you want during dev).

### B.2 — Manual sideload (fallback)

If `office-addin-debugging start` misbehaves on your machine (it sometimes does on macOS Sequoia with hardened runtime), copy the manifest by hand.

The "wef" folder is where PowerPoint for Mac looks for sideloaded add-ins. Its location depends on whether PowerPoint is sandboxed (M365 install) or not.

#### Step 1 — Find or create the wef folder

```bash
WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$WEF_DIR"
echo "Will install to: $WEF_DIR"
```

#### Step 2 — Copy the manifest

From the Streamline repo root:

```bash
cp manifest.xml "$WEF_DIR/streamline.xml"
ls -la "$WEF_DIR"
```

You should see `streamline.xml` listed.

#### Step 3 — Restart PowerPoint

If PowerPoint is open, **fully quit it** (Cmd+Q, not just close window). Re-open PowerPoint.

#### Step 4 — Find the add-in in PowerPoint's UI

In PowerPoint for Mac:

1. **Insert** menu → **Add-ins** → **My Add-ins**
2. The "My Add-ins" dialog opens. Click the **Developer Add-ins** tab at the top of the dialog (it only appears when there's at least one file in the wef folder).
3. **Streamline** should be listed. Double-click it.

The **Streamline** group should now appear on the **Home** ribbon with all four buttons. Click **Streamline** to open the task pane.

> **Note:** The "Developer Add-ins" tab in the dialog is the giveaway that PowerPoint successfully discovered your sideloaded manifest. If the tab doesn't appear, the wef folder path is wrong or the manifest is invalid.

#### Removing a manually sideloaded add-in

```bash
rm "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/streamline.xml"
```

Restart PowerPoint.

---

## Part C — (Optional) Set up the M365 / Copilot backend

Skip this section if you only want to test the core Gantt features (Excel import, rendering, export, etc.). Come back to it when you want to demo Microsoft 365 sign-in or Copilot actions.

### C.1 Register an Entra ID app

Follow [`backend/README.md`](backend/README.md) Step 1 in full. You need:
- Application (client) ID
- Directory (tenant) ID
- Client secret
- Application ID URI

This is done in the Azure portal at [portal.azure.com](https://portal.azure.com) — works fine on Mac in Safari, Chrome, or Edge.

### C.2 Update `manifest.xml` with the registered IDs

Open `manifest.xml` in your editor (VS Code: `code manifest.xml`) and replace:

| Line | Find | Replace with |
|---|---|---|
| 9 | `a1b2c3d4-e5f6-7890-abcd-ef1234567890` | A new GUID. macOS: `uuidgen` |
| 176 | `00000000-0000-0000-0000-000000000000` | Your Entra **client ID** |
| 177 | `api://localhost:3000/00000000-0000-0000-0000-000000000000` | `api://localhost:3000/<your-client-id>` |

Save. Re-sideload the manifest using one of the methods in Part B (or `npm run stop` and `npm run start:desktop` again).

### C.3 Configure and start the backend

In a **third** terminal tab:

```bash
cd backend
cp .env.example .env
nano .env    # or: code .env, vim .env, open -e .env
```

Fill in the four `AAD_*` values from your Entra app, then save.

```bash
npm install
npm start
```

You should see:

```
[server] HTTPS listening on https://localhost:3001
```

### C.4 Verify backend connectivity

In a fourth terminal tab:

```bash
curl -k https://localhost:3001/health
# {"ok":true,"ts":1712...}
```

### C.5 Patch `graphClient.js` to round-trip through the backend

See the patch snippet in [`backend/README.md`](backend/README.md) Step 5. This is one ~15-line change in `src/core/graphClient.js:acquireTokenViaOfficeSSO()`. Save the file — webpack hot-reloads it automatically, no rebuild needed.

After the patch, click **Sign in to Microsoft 365** in the Streamline task pane. The first sign-in prompts you to consent to the Graph scopes. After consent, your name appears in the task pane and the **Import from M365** dropdown becomes active.

---

## Common Mac-specific issues

| Symptom | Diagnosis | Fix |
|---|---|---|
| **"Streamline" group not on Home tab after sideload** | Manifest in wrong wef folder, or PowerPoint not restarted | Verify `ls ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` shows the file; fully quit PowerPoint (Cmd+Q) and reopen |
| **Task pane is white/blank** | Dev server not running or cert not trusted | Visit `https://localhost:3000/taskpane.html` in Safari — if you see a cert warning, repeat A.3 |
| **`Office.auth.getAccessToken` returns 13003** | Office SSO not supported in this PowerPoint build | Update PowerPoint via the Microsoft AutoUpdate app (Help → Check for Updates) — needs build 16.55+ |
| **"Cannot find Office.js"** in browser console | Network blocked or Office CDN unreachable | Check you can reach `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` in a browser; if not, check VPN/firewall |
| **Cmd+R to reload task pane doesn't work** | PowerPoint for Mac doesn't always honor it | Right-click inside the task pane → **Reload** instead. Or close and reopen the task pane. |
| **"Developer Add-ins" tab doesn't appear in My Add-ins dialog** | wef folder is empty or path is wrong | Some M365 builds use a different container path. Try: `~/Library/Group Containers/UBF8T346G9.Office/wef` or the un-sandboxed path `~/Library/Containers/Microsoft PowerPoint/Data/Documents/wef`. Whichever exists, copy manifest there. |
| **`uuidgen` returns uppercase, manifest expects lowercase** | macOS `uuidgen` is uppercase by default | Pipe through tr: `uuidgen \| tr A-F a-f` |
| **Backend: `EADDRINUSE: 3001`** | Port in use | `lsof -i :3001` to find the offender; `kill <pid>` or change `PORT` in `.env` |
| **`EACCES` on `npm install`** | Trying to install globally without sudo | Don't `sudo npm install`; install in the project directory only. If you must install globally, use `nvm` instead. |
| **"Operation not permitted" copying to wef folder** | macOS sandbox protection | System Settings → Privacy & Security → Files and Folders → grant Terminal access to the relevant container |

---

## Updating after code changes

- **Frontend changes** — webpack hot-reloads. Just refresh the task pane (right-click → **Reload** or close/reopen the task pane).
- **Manifest changes** — must re-sideload: `npm run stop` then `npm run start:desktop`. Or copy the file again to the wef folder and restart PowerPoint.
- **Backend changes** — restart `npm start` in the `backend/` directory (no hot reload).

---

## Testing the install

Once the task pane is open, follow [`TESTING-GUIDE.md`](TESTING-GUIDE.md) for a step-by-step verification of every major feature.

---

## Cleanup

When you're done with a session:

```bash
# From the repo root
npm run stop

# Stop the dev server with Ctrl+C in its terminal tab
# Stop the backend with Ctrl+C in its terminal tab
```

To completely uninstall the dev cert:

```bash
npx office-addin-dev-certs uninstall
```

To remove a manually-installed manifest:

```bash
rm "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/streamline.xml"
```

---

## Why Mac sideload matters for the handoff

When you walk into the room with the PowerPoint and Planner teams, **opening Streamline on a Mac** is the demo that closes the deal. Office Timeline literally does not exist on macOS. A working Mac demo proves:

1. The Office.js architecture choice was correct
2. Roughly 40% of M365 users (Mac + Web + iPad) gain a Gantt tool that didn't exist before
3. The product is shippable today on the platforms Office Timeline can never reach

Spend 20 minutes verifying everything works on Mac before any demo. It's the highest-leverage prep you can do.
