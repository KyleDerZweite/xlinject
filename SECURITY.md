# Security Policy

## Project maturity

`xlinject` is in **alpha testing/release**.
The library is usable, but security hardening and API refinements are still ongoing.

## Supported versions

At this stage, only the latest `main` branch state is considered supported for security fixes.

## Reporting a vulnerability

Please report security issues privately via GitHub Security Advisories (preferred).
If that is not possible, contact maintainers directly and avoid public issue disclosure.

When reporting, include:

- Affected version/commit
- Reproduction steps and sample input (sanitized)
- Impact assessment
- Suggested fix (if available)

## Response targets

- Initial acknowledgement: within 7 days
- Triage decision: within 14 days
- Fix timeline: depends on severity and maintainers' availability

## Coordinated disclosure

- Please allow maintainers reasonable time to triage and patch before public disclosure.
- We will publish remediation notes in release notes when fixes are available.

## Scope notes

- `xlinject` performs targeted XLSX XML mutations and does not execute workbook formulas.
- Do not process untrusted files in privileged environments without sandboxing.
- Always validate outputs in your own operational context before production use.
