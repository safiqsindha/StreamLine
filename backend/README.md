# Streamline Backend

A minimal Express service that powers the Microsoft 365 / Copilot integration for the Streamline PowerPoint add-in. It does two things:

1. **OBO token exchange** — accepts the bootstrap token Office SSO issues to the add-in, validates it, and exchanges it for a Microsoft Graph access token via the OAuth 2.0 On-Behalf-Of flow.
2. **Copilot action endpoints** — implements the four operations declared in [`copilot-package/streamline-actions.json`](../copilot-package/streamline-actions.json) so the declarative Copilot agent can call into Streamline.

This is a **dev/sideload** backend. It's deliberately small (~6 source files, ~400 lines) and prioritizes correctness over feature breadth. For production deployment notes see the bottom of this file.

---

## What you need before starting

- Node.js 18 or newer (`node --version`)
- An Azure subscription with permission to create app registrations (Owner or Application Administrator on the directory)
- The Streamline repo cloned and `npm install` already run from the **root** (the backend reuses `office-addin-dev-certs` from the root `node_modules`)

---

## Step 1 — Register the Entra ID (Azure AD) application

This single registration backs both the add-in's Office SSO and the backend's OBO exchange.

1. Open the [Azure portal](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **+ New registration**.
2. **Name:** `Streamline Dev`
3. **Supported account types:** `Accounts in this organizational directory only (Single tenant)` for dev. Switch to multi-tenant only after Microsoft's review.
4. **Redirect URI:** leave blank for now.
5. Click **Register**.
6. On the **Overview** blade, copy these values — you'll paste them into `.env`:
   - **Application (client) ID** → `AAD_CLIENT_ID`
   - **Directory (tenant) ID** → `AAD_TENANT_ID`

### 1a — Expose an API

1. Left nav → **Expose an API** → **Add** next to "Application ID URI".
2. Accept the default `api://<client-id>` or change it to `api://localhost:3000/<client-id>` to match the dev hostname. Click **Save**.
3. Copy the resulting URI — this is your `AAD_API_AUDIENCE`.
4. Click **+ Add a scope**:
   - **Scope name:** `access_as_user`
   - **Who can consent:** Admins and users
   - **Admin consent display name:** `Access Streamline as the signed-in user`
   - **Admin consent description:** `Allows Streamline to call Microsoft Graph on behalf of the signed-in user.`
   - **State:** Enabled
   - Click **Add scope**.
5. Under **Authorized client applications**, click **+ Add a client application** and add these well-known Office client IDs (these tell Entra to skip the consent dialog when called from inside Office):
   - `d3590ed6-52b3-4102-aeff-aad2292ab01c` — Microsoft Office
   - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` — Office on the web
   - `57fb890c-0dab-4253-a5e0-7188c88b2bb4` — Office on iOS
   - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` — Outlook desktop
   - Tick the `access_as_user` scope checkbox for each.

### 1b — Add API permissions for Microsoft Graph

1. Left nav → **API permissions** → **+ Add a permission** → **Microsoft Graph** → **Delegated permissions**.
2. Add (one at a time, or check them all then click "Add permissions"):
   - `User.Read`
   - `Tasks.Read`
   - `Files.Read`
   - `Calendars.Read`
3. (Optional, only if you'll demo SharePoint import or write-back) `Sites.Read.All`, `Tasks.ReadWrite` — these require admin consent.
4. Click **Grant admin consent for <tenant>** at the top of the permissions list. You must be a tenant admin for this button to work.

### 1c — Create a client secret

1. Left nav → **Certificates & secrets** → **+ New client secret**.
2. **Description:** `streamline-backend-dev`
3. **Expires:** 6 months (rotate before expiry — for production use a certificate, not a secret)
4. Click **Add**, then immediately copy the **Value** column. **You will not see it again.**
5. Paste it into `.env` as `AAD_CLIENT_SECRET`.

---

## Step 2 — Configure the backend

```bash
cd backend
cp .env.example .env
```

Open `.env` and fill in the four values you collected above:

```
AAD_TENANT_ID=<your tenant id>
AAD_CLIENT_ID=<your client id>
AAD_CLIENT_SECRET=<your client secret>
AAD_API_AUDIENCE=api://localhost:3000/<your client id>
```

Leave the rest at their defaults for local sideload.

---

## Step 3 — Update the add-in manifest to point at the same app

Open `manifest.xml` (in the repo root) and replace these placeholders with the values from your registration:

| Line | Field | Replace with |
|---|---|---|
| `manifest.xml:9` | `<Id>` | A new GUID you generate (this is the **add-in's** ID — separate from the Entra app ID). On macOS/Linux: `uuidgen`. On Windows PowerShell: `[guid]::NewGuid()`. |
| `manifest.xml:176` | `<WebApplicationInfo><Id>` | Your Entra **client ID** |
| `manifest.xml:177` | `<Resource>` | Your `api://localhost:3000/<client-id>` URI |

Save. The add-in must be re-sideloaded after this change (close PowerPoint and re-add the manifest).

---

## Step 4 — Install backend dependencies and start the server

```bash
cd backend
npm install
npm start
```

You should see:

```
[server] HTTPS listening on https://localhost:3001
[server] CORS origin: https://localhost:3000
[server] Tenant: <your tenant>
```

Test the healthcheck (use `-k` because the dev cert is locally signed):

```bash
curl -k https://localhost:3001/health
# {"ok":true,"ts":1712...}
```

---

## Step 5 — Wire the add-in's GraphClient to call the backend

Currently `src/core/graphClient.js` calls Graph directly with the SSO bootstrap token. That fails in production because the bootstrap token has the wrong audience for Graph.

**Minimal change required** — modify `acquireTokenViaOfficeSSO()` to round-trip through the backend:

```javascript
async acquireTokenViaOfficeSSO() {
  // ... existing Office.auth.getAccessToken() call producing `bootstrap` ...
  const bootstrap = await Office.auth.getAccessToken({...});

  // NEW: exchange the bootstrap token for a Graph token via the backend.
  const r = await fetch("https://localhost:3001/api/graph-token", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${bootstrap}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ scopes: ["User.Read", "Tasks.Read", "Files.Read", "Calendars.Read"] }),
  });
  if (!r.ok) {
    const { error } = await r.json().catch(() => ({}));
    throw new Error(`Backend OBO exchange failed: ${error || r.status}`);
  }
  const { accessToken } = await r.json();
  this.accessToken = accessToken;
  return accessToken;
}
```

This is the only client-side change needed. The rest of `graphClient.js` already uses `this.accessToken` correctly.

---

## How it all fits together

```
┌──────────────────────────────────────────────────────────────────┐
│ PowerPoint (Mac / Windows / Web)                                 │
│ ┌────────────────────────────────────────────────────────────┐  │
│ │ Streamline task pane (https://localhost:3000)              │  │
│ │                                                            │  │
│ │ 1. Office.auth.getAccessToken() → bootstrap JWT            │  │
│ │ 2. POST /api/graph-token  ──────┐                          │  │
│ └────────────────────────────────  │  ─────────────────────────┘  │
│                                    │                              │
│                                    ▼                              │
│ ┌────────────────────────────────────────────────────────────┐   │
│ │ Streamline backend (https://localhost:3001)                │   │
│ │                                                            │   │
│ │ 3. Validate JWT against Entra JWKS                         │   │
│ │ 4. msal.acquireTokenOnBehalfOf() ─────────┐                │   │
│ │ 5. Return Graph access token              │                │   │
│ └────────────────────────────────────────── │ ───────────────┘   │
│                                             │                    │
│                                             ▼                    │
│ ┌────────────────────────────────────────────────────────────┐   │
│ │ Microsoft Entra ID + Microsoft Graph                       │   │
│ └────────────────────────────────────────────────────────────┘   │
└──────────────────────────────────────────────────────────────────┘
```

---

## Endpoints

| Method | Path | Auth | Purpose |
|---|---|---|---|
| `GET` | `/health` | none | Liveness probe |
| `POST` | `/api/graph-token` | bootstrap JWT | Returns a Graph access token via OBO |
| `POST` | `/api/copilot/createGantt` | bootstrap JWT | Validates request, returns render envelope |
| `POST` | `/api/copilot/importFromM365` | bootstrap JWT | Calls Graph for Planner/To Do/Calendar/OneDrive/SharePoint |
| `POST` | `/api/copilot/updateTasks` | bootstrap JWT | Returns update envelope |
| `POST` | `/api/copilot/describeGantt` | bootstrap JWT | Returns describe envelope (real summary is client-side) |

All POST endpoints expect `application/json`. All return `application/json`.

---

## Common errors and fixes

| Symptom | Cause | Fix |
|---|---|---|
| `[server] failed to start: Missing required environment variables` | `.env` not filled in | Copy `.env.example` and fill the four AAD values |
| `office-addin-dev-certs not installed` | Certs not generated | From repo root: `npx office-addin-dev-certs install` |
| `401 invalid_token` from `/api/graph-token` | Wrong audience or expired token | Verify `AAD_API_AUDIENCE` exactly matches the Application ID URI in Entra; re-sign in to refresh the bootstrap token |
| `403 consent_required` | User hasn't consented to Graph scopes | Sign out and back in via the add-in; the second sign-in prompts for consent |
| `403 interaction_required` | Conditional Access policy requires MFA / device compliance | Sign in interactively (a non-SSO sign-in path) to satisfy the policy |
| `500 obo_failed` with `AADSTS50013` | Bootstrap token not actually meant for OBO | Verify the add-in's manifest `<WebApplicationInfo>` `Resource` URI matches `AAD_API_AUDIENCE` |
| `CORS error` in browser console | Backend CORS doesn't include the add-in origin | Set `ALLOWED_ORIGIN` in `.env` to the exact origin (including https:// and port) |

---

## What this backend does NOT do (production checklist)

This is a dev backend. Before production deployment, add:

- [ ] Per-user rate limiting (Redis + token bucket)
- [ ] Tenant isolation enforcement on every cache key
- [ ] Structured logging with correlation IDs (no PII)
- [ ] Application Insights / OpenTelemetry instrumentation
- [ ] Certificate-based auth to Entra (replace `AAD_CLIENT_SECRET` with a cert)
- [ ] Managed Identity for any downstream Azure resources
- [ ] WAF in front (Azure Front Door + OWASP ruleset)
- [ ] CSP / HSTS / `frame-ancestors` headers
- [ ] Health probe distinct from authenticated endpoints
- [ ] Graceful shutdown handler
- [ ] Request size limits per route, not just global
- [ ] Sensitivity label propagation via MIP SDK
- [ ] M365 unified audit log emission for every action
- [ ] Microsoft Purview DLP integration

See `streamline_security_backlog.md` (in your conversation memory) for the complete prioritized list.
