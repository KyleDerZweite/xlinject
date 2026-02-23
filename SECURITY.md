# Security Policy

## Project maturity

`xlinject` is in **early development (alpha)** and is **not fully released**.
Security hardening is ongoing and interfaces may change.

## Supported versions

At this stage, only the latest `main` branch state is considered supported for security fixes.

## Reporting a vulnerability

Please report security issues privately by opening a GitHub Security Advisory (preferred) or by contacting the maintainers directly.

When reporting, include:

- Affected version/commit
- Reproduction steps and sample input (sanitized)
- Impact assessment
- Suggested fix (if available)

## Response targets

- Initial acknowledgement: within 7 days
- Triage decision: within 14 days
- Fix timeline: depends on severity and maintainers' availability

## Scope notes

- `xlinject` performs targeted XLSX XML mutations and does not execute workbook formulas.
- Do not process untrusted files in privileged environments without sandboxing.
- Always validate outputs in your own operational context before production use.
