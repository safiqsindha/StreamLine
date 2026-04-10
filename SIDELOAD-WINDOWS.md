# Sideload & Test Streamline on Windows

This guide walks you from a clean Windows machine to a fully sideloaded Streamline add-in running inside PowerPoint Desktop, with optional backend (M365 / Copilot integration).

**Total time:** 20–30 minutes for the core add-in. Add 30–60 minutes if you also set up the M365 backend.

---

## What you need

- **Windows 10 (build 19041+) or Windows 11**
- **PowerPoint** — one of:
  - Microsoft 365 (any subscription tier with the Office desktop apps)
  - PowerPoint 2021
  - PowerPoint 2019 (limited Office.js version — some features may degrade)
- **Node.js 18 LTS or newer** — download from [nodejs.org](https://nodejs.org). After install, in a new terminal run `node --version` to confirm.
- **Git for Windows** — [git-scm.com](https://git-scm.com)
- **A modern terminal** — Windows Terminal recommended (it ships with Windows 11; on Windows 10 install from the Microsoft Store)
- **Administrator access** — required only for installing the dev HTTPS certificate (one-time)

---

## Part A — Get the code running

### A.1 Clone the repository

Open Windows Terminal (PowerShell tab):

```powershell
cd $HOME\Desktop
git clone <your-streamline-repo-url> Streamline
cd Streamline
```

### A.2 Install dependencies

```powershell
npm install
```

This pulls Office.js, webpack, sheetjs, the dev cert tool, and the rest. Expect 30–90 seconds depending on network speed.

### A.3 Generate and trust the local HTTPS certificate

PowerPoint refuses to load add-ins served over plain HTTP, and won't load HTTPS with an untrusted cert. The `office-addin-dev-certs` tool generates a self-signed cert and installs it into the Windows certificate store as a trusted root.

```powershell
npx office-addin-dev-certs install
```

You will see a UAC prompt — accept it. The tool prints a confirmation when the cert is installed. Verify with:

```powershell
npx office-addin-dev-certs verify
# You should have trusted access to https://localhost
```

If verification fails, reboot once and retry — Windows sometimes caches the previous untrusted state.

### A.4 Build and start the dev server

```powershell
npm start
```

You should see webpack output ending with something like:

```
webpack 5.x.x compiled successfully
[webpack-dev-server] Project is running at:
[webpack-dev-server] Loopback: https://localhost:3000/
```

Leave this terminal running. Open a **new** terminal tab for the next steps.

### A.5 Sanity check the dev server

In your browser, navigate to `https://localhost:3000/taskpane.html`. You should see the Streamline UI render (without the PowerPoint host APIs working, but the HTML/CSS/JS is loaded). If you see a certificate warning, the dev cert isn't trusted — repeat step A.3.

---

## Part B — Sideload the manifest into PowerPoint

There are three ways to do this on Windows. Pick **B.1** (the office-addin-debugging method) for normal dev. Use B.2 if B.1 misbehaves. Use B.3 for headless/CI scenarios.

### B.1 — Easiest: `office-addin-debugging start` (recommended)

The Streamline `package.json` already has this script wired up:

```powershell
npm run start:desktop
```

What this does:
1. Starts (or reuses) the webpack dev server on https://localhost:3000
2. Registers `manifest.xml` in the Windows registry under the developer add-ins key
3. Launches PowerPoint
4. Adds Streamline to the **Home ribbon** automatically

When PowerPoint opens, you'll see a **Streamline** group on the **Home** tab with **Streamline**, **Refresh**, **Today Line**, and **Export PNG** buttons. Click **Streamline** to open the task pane.

When you're done testing:

```powershell
npm run stop
```

This unregisters the manifest. (Without it, the registration persists and `npm run start:desktop` will reuse it next time, which is usually what you want.)

### B.2 — Network share method (fallback when B.1 doesn't work)

This is the original Office add-in sideload method. It works when `office-addin-debugging` has issues with your AV or registry permissions.

1. **Create a folder to use as a shared catalog.** It can be anywhere; for example `C:\OfficeAddins`.

   ```powershell
   mkdir C:\OfficeAddins
   ```

2. **Share the folder.** Right-click `C:\OfficeAddins` → **Properties** → **Sharing** tab → **Share…** → add your own user with **Read** permission → **Share** → **Done**. Note the network path it gives you, e.g. `\\YOUR-PC-NAME\OfficeAddins`.

3. **Copy the manifest into the share:**

   ```powershell
   copy manifest.xml C:\OfficeAddins\
   ```

4. **Tell PowerPoint to trust the catalog.** Open PowerPoint → **File** → **Options** → **Trust Center** → **Trust Center Settings…** → **Trusted Add-in Catalogs**.
   - **Catalog URL:** paste the network path from step 2 (`\\YOUR-PC-NAME\OfficeAddins`)
   - Click **Add catalog**
   - Tick **Show in Menu**
   - Click **OK** twice
   - **Close PowerPoint completely** (File → Exit, not just close-window)

5. **Re-open PowerPoint** → **Insert** → **My Add-ins** → switch to the **SHARED FOLDER** tab → you should see **Streamline** → click **Add**.

6. The Streamline group appears on the Home ribbon. Click the **Streamline** button to open the task pane.

When you update `manifest.xml` later, copy the new version into the shared folder and click **Refresh** in the My Add-ins dialog.

### B.3 — Registry sideload (headless / scripted)

For automated lab setups, you can write the manifest path directly to the registry:

```powershell
$manifestPath = (Resolve-Path .\manifest.xml).Path
$keyPath = "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer"
if (-not (Test-Path $keyPath)) { New-Item -Path $keyPath -Force | Out-Null }
New-ItemProperty -Path $keyPath -Name "Streamline" -Value $manifestPath -PropertyType String -Force
```

Then launch PowerPoint normally — the add-in appears under **Insert → My Add-ins → Developer Add-ins**.

To remove:

```powershell
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer" -Name "Streamline"
```

---

## Part C — (Optional) Set up the M365 / Copilot backend

Skip this section if you only want to test the core Gantt features (Excel import, rendering, export, etc.). Come back to it when you want to demo Microsoft 365 sign-in or Copilot actions.

### C.1 Register an Entra ID app

Follow [`backend/README.md`](backend/README.md) Step 1 in full. You need:
- Application (client) ID
- Directory (tenant) ID
- Client secret
- Application ID URI

### C.2 Update `manifest.xml` with the registered IDs

Open `manifest.xml` in VS Code (or any editor) and replace these placeholders:

| Line | Find | Replace with |
|---|---|---|
| 9 | `a1b2c3d4-e5f6-7890-abcd-ef1234567890` | A new GUID. PowerShell: `[guid]::NewGuid().ToString()` |
| 176 | `00000000-0000-0000-0000-000000000000` | Your Entra **client ID** |
| 177 | `api://localhost:3000/00000000-0000-0000-0000-000000000000` | `api://localhost:3000/<your-client-id>` |

Save. Re-sideload the manifest using one of the methods in Part B (or run `npm run stop` and `npm run start:desktop` again).

### C.3 Configure and start the backend

In a **third** terminal tab:

```powershell
cd backend
copy .env.example .env
notepad .env
```

Fill in the four `AAD_*` values from your Entra app, then save.

```powershell
npm install
npm start
```

You should see:

```
[server] HTTPS listening on https://localhost:3001
```

### C.4 Verify backend connectivity

In a fourth terminal tab:

```powershell
curl.exe -k https://localhost:3001/health
# {"ok":true,"ts":1712...}
```

If `curl.exe` isn't available, use PowerShell:

```powershell
Invoke-WebRequest -Uri https://localhost:3001/health -SkipCertificateCheck
```

### C.5 Patch `graphClient.js` to round-trip through the backend

See the patch snippet in [`backend/README.md`](backend/README.md) Step 5. This is one ~15-line change in `src/core/graphClient.js:acquireTokenViaOfficeSSO()`. Save the file — webpack hot-reloads it automatically.

After the patch, click **Sign in to Microsoft 365** in the Streamline task pane. The first time, you'll be prompted to consent to the Graph scopes. After consent, your name appears in the task pane and the **Import from M365** dropdown becomes active.

---

## Common Windows-specific issues

| Symptom | Diagnosis | Fix |
|---|---|---|
| **Task pane shows "Couldn't open this add-in"** | Manifest validation failed | Run `npm run validate` from the repo root. Fix any reported issues, re-sideload. |
| **Task pane is blank / white** | Dev server not running or HTTPS cert not trusted | Verify `npm start` is still running; visit `https://localhost:3000/taskpane.html` in Edge — if it shows a cert warning, repeat A.3 |
| **"This add-in is not from a trusted source"** dialog | The add-in domain isn't whitelisted | In the dialog click **Trust this add-in** |
| **Streamline button not visible on Home tab after sideload** | Manifest changed but not re-loaded | `npm run stop` then `npm run start:desktop` again |
| **Console error: `Office.context.requirements.isSetSupported('Presentation', '1.x') = false`** | Office build is too old | Update Office to a current Microsoft 365 build (File → Account → Update Options → Update Now) |
| **Office.auth.getAccessToken returns 13003** | SSO not supported in this build of PowerPoint | Use a current M365 PowerPoint build; SSO requires Office build 16.0.13127+ |
| **PowerShell: `npx` not recognized** | Node.js not in PATH | Re-run the Node installer and tick "Add to PATH"; restart your terminal |
| **`EPERM` errors during `npm install`** | Antivirus locking node_modules | Add the Streamline folder to your AV exclusions, or run install from a non-OneDrive path |
| **OneDrive sync conflicts on `dist/`** | Repo lives inside a OneDrive folder | Move the repo to `C:\Streamline` (outside OneDrive) |
| **Backend: `EADDRINUSE` on port 3001** | Another process is using the port | `netstat -ano \| findstr :3001` to find it; kill it or change `PORT` in `.env` |

---

## Updating after code changes

- **Frontend changes** — webpack hot-reloads. Just refresh the task pane (right-click inside it → **Reload**, or close and re-open).
- **Manifest changes** — must re-sideload: `npm run stop` then `npm run start:desktop`.
- **Backend changes** — restart `npm start` in the `backend/` directory (no hot reload).

---

## Testing the install

Once the task pane is open, follow [`TESTING-GUIDE.md`](TESTING-GUIDE.md) for a step-by-step verification of every major feature.

## Cleanup

When you're done with a session:

```powershell
# Stop the add-in registration
cd <streamline-repo>
npm run stop

# Stop the dev server (Ctrl+C in its terminal)
# Stop the backend (Ctrl+C in its terminal)
```

To completely uninstall the dev cert:

```powershell
npx office-addin-dev-certs uninstall
```
