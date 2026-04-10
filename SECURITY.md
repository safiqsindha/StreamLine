# Security Policy

## Reporting a Vulnerability

If you discover a security vulnerability in Streamline, **please do not open a public issue**. Instead, report it privately so it can be triaged and patched before disclosure.

### How to report

Send a private report via one of the following:

- **GitHub Security Advisories** — preferred. Open a draft advisory at [Security → Advisories → Report a vulnerability](../../security/advisories/new) on this repository.
- **Email** — `<security-contact-email>` (replace before publishing)

Please include:

1. A description of the vulnerability
2. Steps to reproduce (or a proof-of-concept)
3. The version / commit SHA where you observed it
4. The potential impact (data exposure, privilege escalation, etc.)
5. Any suggested mitigation

### Response commitment

| Severity | Acknowledgement | Patch target |
|---|---|---|
| Critical | Within 24 hours | Within 7 days |
| High | Within 48 hours | Within 14 days |
| Medium | Within 7 days | Within 30 days |
| Low | Within 14 days | Next release |

We will keep you informed of progress and credit you in the release notes (with your permission) once a patch is available.

## Scope

In scope:
- The Streamline PowerPoint add-in (`src/`, `manifest.xml`)
- The Streamline backend service (`backend/`)
- The Copilot agent and Teams message extension packages (`copilot-package/`, `teams-package/`)
- Build configuration and dependency tree

Out of scope:
- Vulnerabilities in upstream dependencies (please report to the upstream project, then notify us so we can pin / patch)
- Issues requiring a compromised host machine (e.g., physical access to a developer's laptop)
- Self-XSS via developer tools console
- Missing security headers on the localhost dev server (production hardening is tracked separately)

## Supported versions

This project is in pre-release / pre-handoff status. Only the `main` branch is supported. Older commits and branches do not receive security updates.

## Security practices

This codebase is being prepared for handoff to Microsoft's internal PowerPoint and Planner teams. As part of that work, the following security controls are in active development:

- Microsoft Threat Modeling Tool review (STRIDE)
- CodeQL static analysis on every PR
- Dependabot version updates and security alerts
- Secret scanning with push protection
- Third-party penetration test (planned pre-handoff)
- AI red-team review of the Copilot agent (planned pre-handoff)

A full prioritized security backlog is maintained internally and shared with reviewers under NDA.
