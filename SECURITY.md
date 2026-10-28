# Security Policy

## Project maturity

`xlinject` is in **beta prerelease**.
The library is usable, but security hardening and API refinements are still ongoing.

## Supported versions

At this stage, only the latest `main` branch state is supported for security fixes.
The latest published package is `0.2.0b2`, a beta prerelease. Security fixes on
`main` are not available to package consumers until a new release is published.
There are no maintained backport branches.

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
