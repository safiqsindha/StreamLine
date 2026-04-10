/**
 * Streamline Backend — Graph token exchange endpoint
 *
 * The Office add-in calls POST /api/graph-token with its bootstrap token in
 * the Authorization header. This endpoint runs the OBO exchange and returns
 * a Microsoft Graph access token the add-in can attach to its own Graph
 * REST calls (via graphClient.setAccessToken()).
 *
 * Body (optional): { scopes: ["User.Read", ...] }
 * Response: { accessToken: "...", expiresOn: 1234567890 }
 *
 * Error codes (HTTP 401/403/500):
 *   401 invalid_token         — bootstrap token failed validation
 *   403 consent_required      — user must consent to scopes (AADSTS65001)
 *   403 interaction_required  — interactive sign-in required (CA, MFA)
 *   500 obo_failed            — other OBO failure
 */

const express = require("express");
const { validateBootstrapToken, exchangeForGraphToken, DEFAULT_GRAPH_SCOPES } = require("../auth");

const router = express.Router();

router.post("/graph-token", validateBootstrapToken, async (req, res) => {
  // Allow caller to request a narrower scope set than the default.
  let scopes = DEFAULT_GRAPH_SCOPES;
  if (req.body && Array.isArray(req.body.scopes) && req.body.scopes.length > 0) {
    // Normalize: prepend the Graph resource URL if the caller passed bare scope names.
    scopes = req.body.scopes.map((s) =>
      s.startsWith("https://") ? s : `https://graph.microsoft.com/${s}`
    );
  }

  try {
    const accessToken = await exchangeForGraphToken(req.user.rawJwt, scopes);
    // Decode expiry from the token (without verifying — we just got it).
    let expiresOn = null;
    try {
      const payloadB64 = accessToken.split(".")[1];
      const payload = JSON.parse(Buffer.from(payloadB64, "base64").toString("utf8"));
      expiresOn = payload.exp;
    } catch (_) {
      /* non-fatal */
    }

    res.json({ accessToken, expiresOn });
  } catch (err) {
    const code = err.errorCode || "";
    if (code === "invalid_grant" || (err.subError && err.subError === "consent_required")) {
      return res.status(403).json({ error: "consent_required" });
    }
    if (err.subError === "basic_action" || err.subError === "additional_action") {
      return res.status(403).json({ error: "interaction_required" });
    }
    console.error("[graph] OBO exchange failed:", err.message);
    res.status(500).json({ error: "obo_failed" });
  }
});

module.exports = router;
