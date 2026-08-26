# Security Policy

## Supported Versions

LoanLogic is developed on the `main` branch, and fixes land there. Only the
latest commit on `main` receives security updates.

| Version | Supported |
| ------- | --------- |
| `main` (latest) | ✅ |
| Older tags / forks | ❌ |

## Reporting a Vulnerability

Please **do not open a public issue** for a security problem.

Report it privately to **yevtushenkobohdanyevbank@gmail.com** with:

* what the issue is and where in the code it lives,
* how to reproduce it,
* what an attacker could achieve with it.

You can expect an acknowledgement within a few days. If the report is valid,
the fix and a note in the commit history follow; credit is given unless you
ask otherwise.

## Scope

LoanLogic is a self-hosted Streamlit application. It has no backend of its own
and stores no user accounts — the operator runs it and controls the
environment it runs in. Reports are most useful when they concern:

* **Injection through user input.** The app renders parts of its UI with
  `unsafe_allow_html=True`, so anything a user can type that reaches HTML is
  in scope. Currency symbols are the obvious path and are sanitized in
  `_sanitize_currency_symbol()`; a bypass of that is a valid report.
* **Theme data.** Custom themes round-trip through browser `localStorage` and
  a `?theme_b64=` query parameter, and their colors are interpolated into a
  `<style>` block. `build_css()` accepts only well-formed hex values for this
  reason; a way to smuggle other CSS through it is in scope.
* **Email headers.** `send_report_email()` strips CR/LF from header values to
  prevent header injection. A bypass is in scope.
* **Secrets handling.** SMTP credentials come from `.streamlit/secrets.toml`,
  which is gitignored and must never be committed. Anything that causes them
  to be logged, rendered, or written into an exported report is in scope.
* **Export files.** Crafted input that makes an `.xlsx`, `.pdf`, `.docx` or
  `.csv` export dangerous to open — CSV formula injection in particular.

## Out of Scope

* Findings that require the operator to deliberately misconfigure the app, or
  that only apply when it is deployed publicly without authentication in front
  of it. LoanLogic ships no authentication and does not claim to.
* Vulnerabilities in Streamlit, pandas, ReportLab or other dependencies —
  please report those upstream. If a dependency version pinned in
  `requirements.txt` is known-vulnerable, that *is* worth telling us about.
* Denial of service through deliberately extreme inputs (a 100-trillion
  principal over 1200 periods). The calculation core validates and bounds its
  inputs, but the app is not hardened against an operator attacking their own
  instance.

## A Note on the Numbers

LoanLogic performs financial calculations for informational purposes only. A
wrong number is a **bug**, not a vulnerability — please report incorrect math
as a normal issue, with the inputs you used and the figure you expected.
