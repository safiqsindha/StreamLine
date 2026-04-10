/**
 * Streamline Backend — Configuration loader
 *
 * Loads environment variables from .env, validates required keys, and
 * exposes a frozen config object. Fails fast on startup if anything is
 * missing so misconfiguration can't silently produce broken auth flows.
 */

require("dotenv").config();

const REQUIRED_KEYS = [
  "AAD_TENANT_ID",
  "AAD_CLIENT_ID",
  "AAD_CLIENT_SECRET",
  "AAD_API_AUDIENCE",
];

function validate() {
  const missing = REQUIRED_KEYS.filter((k) => !process.env[k]);
  if (missing.length > 0) {
    throw new Error(
      `Missing required environment variables: ${missing.join(", ")}. ` +
        `Copy backend/.env.example to backend/.env and fill in values from your Entra ID app registration.`
    );
  }
}

validate();

const config = Object.freeze({
  tenantId: process.env.AAD_TENANT_ID,
  clientId: process.env.AAD_CLIENT_ID,
  clientSecret: process.env.AAD_CLIENT_SECRET,
  apiAudience: process.env.AAD_API_AUDIENCE,

  port: parseInt(process.env.PORT || "3001", 10),
  allowedOrigin: process.env.ALLOWED_ORIGIN || "https://localhost:3000",
  useHttps: (process.env.USE_HTTPS || "true").toLowerCase() === "true",
  logLevel: process.env.LOG_LEVEL || "info",

  // Derived URLs
  authority: `https://login.microsoftonline.com/${process.env.AAD_TENANT_ID}`,
  jwksUri: `https://login.microsoftonline.com/${process.env.AAD_TENANT_ID}/discovery/v2.0/keys`,
  issuer: `https://login.microsoftonline.com/${process.env.AAD_TENANT_ID}/v2.0`,
});

module.exports = config;
