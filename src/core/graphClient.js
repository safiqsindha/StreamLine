/**
 * Streamline Graph Client
 *
 * Thin wrapper over Microsoft Graph REST endpoints. Handles auth via two paths:
 *   1. Office SSO (`Office.auth.getAccessToken`) - native add-in SSO, preferred
 *      when running inside PowerPoint because it's silent and uses the signed-in
 *      M365 identity. Requires `<WebApplicationInfo>` in the manifest.
 *   2. Bearer token injection - used by the declarative Copilot agent and Teams
 *      message extension, where the caller already holds an OBO (on-behalf-of)
 *      access token issued via Azure AD.
 *
 * This module intentionally does NOT pull in @azure/msal-browser at runtime -
 * that dependency is only needed in hosts that don't support Office SSO. The
 * add-in task pane uses Office SSO; server-side callers (Copilot plugin, Teams
 * bot) supply their own tokens via `setAccessToken()`.
 *
 * All endpoint methods return already-parsed JSON and throw on non-2xx.
 */

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

// Scopes needed by Streamline. Kept here so manifest and code stay in sync.
const REQUIRED_SCOPES = [
  "User.Read",
  "Tasks.Read",
  "Tasks.ReadWrite",
  "Files.Read",
  "Files.Read.All",
  "Sites.Read.All",
  "Calendars.Read",
];

class GraphClient {
  constructor(options = {}) {
    this.accessToken = options.accessToken || null;
    this.fetchImpl = options.fetch || (typeof fetch !== "undefined" ? fetch : null);
    this.lastError = null;
  }

  /**
   * Manually set an access token (used by server-side callers - Copilot agent,
   * Teams bot - that obtained a token through the on-behalf-of flow).
   */
  setAccessToken(token) {
    this.accessToken = token;
  }

  hasAccessToken() {
    return !!this.accessToken;
  }

  /**
   * Acquire a token via Office SSO. Only works inside an Office Add-in host
   * (task pane, dialog). Returns the token string; also caches it internally.
   *
   * Requires `<WebApplicationInfo>` in the manifest. If SSO fails (user not
   * consented, host doesn't support it, etc.) the error is re-thrown so the
   * caller can fall back to a popup sign-in.
   */
  async acquireTokenViaOfficeSSO() {
    if (typeof Office === "undefined" || !Office.auth || !Office.auth.getAccessToken) {
      throw new Error("Office.auth.getAccessToken is not available in this host.");
    }

    try {
      const token = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true,
      });
      this.accessToken = token;
      return token;
    } catch (err) {
      this.lastError = err;
      const code = err && err.code;
      // Known SSO error codes - wrap with friendlier messages
      const msg =
        code === 13001 ? "User is not signed in to Office." :
        code === 13002 ? "User declined consent for Streamline to access Microsoft 365." :
        code === 13003 ? "Office SSO not supported in this host." :
        code === 13005 ? "Anonymous user cannot use SSO." :
        code === 13006 || code === 13007 ? "Transient SSO error. Try again." :
        code === 13008 ? "SSO is already in progress." :
        `SSO failed (${code || "unknown"}): ${err.message || err}`;
      const wrapped = new Error(msg);
      wrapped.originalCode = code;
      wrapped.isSSO = true;
      throw wrapped;
    }
  }

  async _request(method, path, body = null, extraHeaders = {}) {
    if (!this.fetchImpl) {
      throw new Error("No fetch implementation available.");
    }
    if (!this.accessToken) {
      throw new Error("GraphClient: no access token set. Call acquireTokenViaOfficeSSO() or setAccessToken() first.");
    }

    const url = path.startsWith("http") ? path : `${GRAPH_BASE}${path}`;
    const headers = {
      Authorization: `Bearer ${this.accessToken}`,
      Accept: "application/json",
      ...extraHeaders,
    };
    if (body !== null) {
      headers["Content-Type"] = "application/json";
    }

    const res = await this.fetchImpl(url, {
      method,
      headers,
      body: body !== null ? JSON.stringify(body) : undefined,
    });

    if (!res.ok) {
      let detail = "";
      try {
        const errBody = await res.json();
        detail = (errBody && errBody.error && errBody.error.message) || JSON.stringify(errBody);
      } catch (_) {
        detail = await res.text().catch(() => "");
      }
      const err = new Error(`Graph ${method} ${path} failed: ${res.status} ${res.statusText}. ${detail}`);
      err.status = res.status;
      throw err;
    }

    // 204 No Content
    if (res.status === 204) return null;
    return res.json();
  }

  // ── User ──────────────────────────────────────────────
  async getMe() {
    return this._request("GET", "/me");
  }

  // ── Planner ───────────────────────────────────────────
  async getMyPlans() {
    const res = await this._request("GET", "/me/planner/plans");
    return res.value || [];
  }

  async getPlanTasks(planId) {
    if (!planId) throw new Error("planId required");
    const res = await this._request("GET", `/planner/plans/${encodeURIComponent(planId)}/tasks`);
    return res.value || [];
  }

  async getPlanBuckets(planId) {
    if (!planId) throw new Error("planId required");
    const res = await this._request("GET", `/planner/plans/${encodeURIComponent(planId)}/buckets`);
    return res.value || [];
  }

  // ── To Do ─────────────────────────────────────────────
  async getTodoLists() {
    const res = await this._request("GET", "/me/todo/lists");
    return res.value || [];
  }

  async getTodoTasks(listId) {
    if (!listId) throw new Error("listId required");
    const res = await this._request("GET", `/me/todo/lists/${encodeURIComponent(listId)}/tasks`);
    return res.value || [];
  }

  // ── Files / OneDrive / SharePoint ─────────────────────
  async getOneDriveChildren(folderPath = "") {
    const p = folderPath
      ? `/me/drive/root:/${encodeURIComponent(folderPath)}:/children`
      : "/me/drive/root/children";
    const res = await this._request("GET", p);
    return res.value || [];
  }

  async downloadDriveItem(itemId) {
    // /drive/items/{id}/content returns a 302 to a pre-signed download URL
    // which fetch follows transparently. Returns an ArrayBuffer.
    if (!itemId) throw new Error("itemId required");
    if (!this.fetchImpl) throw new Error("No fetch implementation available.");
    const url = `${GRAPH_BASE}/me/drive/items/${encodeURIComponent(itemId)}/content`;
    const res = await this.fetchImpl(url, {
      method: "GET",
      headers: { Authorization: `Bearer ${this.accessToken}` },
    });
    if (!res.ok) {
      const err = new Error(`Graph download failed: ${res.status}`);
      err.status = res.status;
      throw err;
    }
    return res.arrayBuffer();
  }

  async searchDrive(query) {
    if (!query) throw new Error("query required");
    const res = await this._request("GET", `/me/drive/root/search(q='${encodeURIComponent(query)}')`);
    return res.value || [];
  }

  // ── SharePoint ────────────────────────────────────────
  async getSiteLists(siteId) {
    if (!siteId) throw new Error("siteId required");
    const res = await this._request("GET", `/sites/${encodeURIComponent(siteId)}/lists`);
    return res.value || [];
  }

  async getListItems(siteId, listId) {
    if (!siteId || !listId) throw new Error("siteId and listId required");
    const res = await this._request(
      "GET",
      `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items?expand=fields`
    );
    return res.value || [];
  }

  // ── Calendar ──────────────────────────────────────────
  async getCalendarEvents(fromIso, toIso) {
    const q = fromIso && toIso
      ? `?startDateTime=${encodeURIComponent(fromIso)}&endDateTime=${encodeURIComponent(toIso)}`
      : "";
    const res = await this._request("GET", `/me/calendar/calendarView${q}`);
    return res.value || [];
  }
}

module.exports = { GraphClient, GRAPH_BASE, REQUIRED_SCOPES };
