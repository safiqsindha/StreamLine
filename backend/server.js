/**
 * Streamline Backend — Express entry point
 *
 * Hosts the OBO token exchange and the four Copilot action endpoints.
 * Serves HTTPS using office-addin-dev-certs (the same certs the webpack
 * dev server uses) so the browser trusts cross-origin calls from the
 * add-in's task pane.
 *
 * Run:
 *   cd backend
 *   cp .env.example .env       # then fill in real values
 *   npm install
 *   npm start
 *
 * Healthcheck: GET https://localhost:3001/health → {"ok": true}
 */

const express = require("express");
const cors = require("cors");
const https = require("https");
const http = require("http");

const config = require("./src/config");
const graphRouter = require("./src/routes/graph");
const copilotRouter = require("./src/routes/copilot");

const app = express();

// Body parsing — keep payloads small; Copilot requests should never exceed
// a few hundred KB. Larger limits invite DoS.
app.use(express.json({ limit: "256kb" }));

// CORS — allow only the add-in's webpack dev server origin. Credentials
// stay false because we use bearer tokens, not cookies.
app.use(
  cors({
    origin: config.allowedOrigin,
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
    credentials: false,
    maxAge: 600,
  })
);

// Security headers — minimal but covers the obvious gaps. Production should
// add CSP, HSTS, and frame-ancestors via a reverse proxy (Front Door / nginx).
app.use((_req, res, next) => {
  res.setHeader("X-Content-Type-Options", "nosniff");
  res.setHeader("X-Frame-Options", "DENY");
  res.setHeader("Referrer-Policy", "no-referrer");
  next();
});

// Healthcheck — unauthenticated, returns 200 + a tiny payload.
app.get("/health", (_req, res) => res.json({ ok: true, ts: Date.now() }));

// Routes
app.use("/api", graphRouter);
app.use("/api/copilot", copilotRouter);

// 404 fallback
app.use((req, res) => res.status(404).json({ error: "not_found", path: req.path }));

// Error handler — never leak stack traces to the client.
app.use((err, _req, res, _next) => {
  console.error("[server] unhandled:", err);
  res.status(500).json({ error: "internal_error" });
});

// ── Server start ──────────────────────────────────────────────────────────
async function start() {
  if (config.useHttps) {
    let httpsOptions;
    try {
      const devCerts = require("office-addin-dev-certs");
      httpsOptions = await devCerts.getHttpsServerOptions();
    } catch (e) {
      console.error(
        "[server] office-addin-dev-certs not installed or certs not generated.\n" +
          "  Run from the repo root:  npx office-addin-dev-certs install\n" +
          "  Then restart the backend, or set USE_HTTPS=false in .env to run plain HTTP."
      );
      process.exit(1);
    }
    https.createServer(httpsOptions, app).listen(config.port, () => {
      console.log(`[server] HTTPS listening on https://localhost:${config.port}`);
      console.log(`[server] CORS origin: ${config.allowedOrigin}`);
      console.log(`[server] Tenant: ${config.tenantId}`);
    });
  } else {
    http.createServer(app).listen(config.port, () => {
      console.log(`[server] HTTP listening on http://localhost:${config.port}`);
    });
  }
}

start().catch((err) => {
  console.error("[server] failed to start:", err);
  process.exit(1);
});
