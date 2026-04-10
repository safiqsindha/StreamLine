/**
 * Streamline Backend — Authentication
 *
 * Two responsibilities:
 *   1. validateBootstrapToken(): Express middleware that validates the JWT
 *      issued to the Office add-in by Office SSO. Verifies signature against
 *      Microsoft's JWKS, checks audience, issuer, and expiration. Attaches
 *      the parsed token + raw JWT to req.user.
 *
 *   2. exchangeForGraphToken(): performs the OAuth 2.0 On-Behalf-Of flow,
 *      swapping the bootstrap token for a Microsoft Graph access token using
 *      the backend's confidential client credentials.
 *
 * Both rely on the Entra ID app registration described in backend/README.md.
 */

const jwt = require("jsonwebtoken");
const jwksClient = require("jwks-rsa");
const msal = require("@azure/msal-node");
const config = require("./config");

// ── JWKS client (cached, rate-limited) ────────────────────────────────────
const jwks = jwksClient({
  jwksUri: config.jwksUri,
  cache: true,
  cacheMaxAge: 24 * 60 * 60 * 1000, // 24h
  rateLimit: true,
  jwksRequestsPerMinute: 10,
});

function getSigningKey(header, callback) {
  jwks.getSigningKey(header.kid, (err, key) => {
    if (err) return callback(err);
    callback(null, key.getPublicKey());
  });
}

// ── MSAL confidential client (singleton) ──────────────────────────────────
const msalApp = new msal.ConfidentialClientApplication({
  auth: {
    clientId: config.clientId,
    authority: config.authority,
    clientSecret: config.clientSecret,
  },
});

/**
 * Express middleware: validate the bearer token in Authorization header.
 * On success, attaches:
 *   req.user.token   - parsed JWT claims
 *   req.user.rawJwt  - the raw bearer string (needed for OBO exchange)
 *
 * Returns 401 on any failure. Error responses contain ONLY a code, never
 * stack traces or token contents (avoid leaking through telemetry).
 */
function validateBootstrapToken(req, res, next) {
  const header = req.headers.authorization || "";
  if (!header.startsWith("Bearer ")) {
    return res.status(401).json({ error: "missing_bearer_token" });
  }
  const rawJwt = header.substring("Bearer ".length).trim();

  jwt.verify(
    rawJwt,
    getSigningKey,
    {
      audience: config.apiAudience,
      issuer: config.issuer,
      algorithms: ["RS256"],
    },
    (err, decoded) => {
      if (err) {
        // Don't leak the failure reason to the client; log server-side.
        console.warn("[auth] JWT validation failed:", err.message);
        return res.status(401).json({ error: "invalid_token" });
      }
      req.user = { token: decoded, rawJwt };
      next();
    }
  );
}

/**
 * Exchange the user's bootstrap token for a Microsoft Graph access token
 * via the OAuth 2.0 On-Behalf-Of flow.
 *
 * @param {string} bootstrapToken  Raw JWT from req.user.rawJwt
 * @param {string[]} scopes        Graph scopes to request, e.g.
 *                                  ["https://graph.microsoft.com/User.Read"]
 * @returns {Promise<string>}      Graph access token
 */
async function exchangeForGraphToken(bootstrapToken, scopes) {
  if (!bootstrapToken) throw new Error("exchangeForGraphToken: bootstrapToken required");
  if (!Array.isArray(scopes) || scopes.length === 0) {
    throw new Error("exchangeForGraphToken: scopes array required");
  }

  try {
    const result = await msalApp.acquireTokenOnBehalfOf({
      oboAssertion: bootstrapToken,
      scopes,
    });
    if (!result || !result.accessToken) {
      throw new Error("OBO exchange returned no access token");
    }
    return result.accessToken;
  } catch (err) {
    // Surface the OBO error code so callers can map to "needs consent" UX.
    const wrapped = new Error(`OBO exchange failed: ${err.errorCode || err.name || "unknown"}`);
    wrapped.errorCode = err.errorCode;
    wrapped.subError = err.subError;
    throw wrapped;
  }
}

/**
 * Default Graph scopes used when the caller doesn't specify. Mirrors the
 * scopes declared in manifest.xml's <WebApplicationInfo> block.
 */
const DEFAULT_GRAPH_SCOPES = [
  "https://graph.microsoft.com/User.Read",
  "https://graph.microsoft.com/Tasks.Read",
  "https://graph.microsoft.com/Files.Read",
  "https://graph.microsoft.com/Calendars.Read",
];

module.exports = {
  validateBootstrapToken,
  exchangeForGraphToken,
  DEFAULT_GRAPH_SCOPES,
};
