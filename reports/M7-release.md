# M7 release validation

## 2026-08-23 — IN PROGRESS

- Added failure-first release checks. Initial collection failed because `scripts.verify_release` did not exist;
  the release boundary is now represented by an executable standard-library verifier and regression tests.
- README release metadata now matches version `1.0.0`, lists all three systems, removes placeholder contact links,
  documents CPython 3.12, 1 GB free space, Linux EGL prerequisites, repair/conflict handling, diagnostic locations,
  supported format boundaries and `PENDING_USER_VALIDATION` for unknown instrument variants.
- `pyproject.toml` now publishes the real repository and issue tracker URLs.
- Local `uv 0.12.5` regenerated `requirements.lock` byte-for-byte from `pyproject.toml` with Python 3.12,
  universal resolution and hashes. The same operation is now a dedicated CI release job.
- The native matrix now measures each Essentials environment, installs the same-version full PySide6 only inside the
  disposable runner, and records logical bytes and savings. Results remain pending until the workflow runs.
- No real instrument file was available. Automated fixtures remain accepted, while real TXT/DAT/VSM/XLS/XLSX
  validation and a human desktop-session install remain pending rather than being inferred from CI.
