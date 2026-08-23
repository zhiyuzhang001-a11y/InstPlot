# M7 release validation

## 2026-08-23 — AUTOMATION COMPLETE / PENDING_USER_VALIDATION

- Added failure-first release checks. Initial collection failed because `scripts.verify_release` did not exist;
  the release boundary is now represented by an executable standard-library verifier and regression tests.
- README release metadata now matches version `1.0.0`, lists all three systems, removes placeholder contact links,
  documents CPython 3.12, 1 GB free space, Linux EGL prerequisites, repair/conflict handling, diagnostic locations,
  supported format boundaries and `PENDING_USER_VALIDATION` for unknown instrument variants.
- `pyproject.toml` now publishes the real repository and issue tracker URLs.
- Local `uv 0.12.5` regenerated `requirements.lock` byte-for-byte from `pyproject.toml` with Python 3.12,
  universal resolution and hashes. The same operation is now a dedicated CI release job.
- The footprint probe initially installed and removed full PySide6 in the test environment. Although its measurement
  passed, uninstalling shared Qt files damaged later imports on all three systems. The final probe copies the environment
  to a temporary directory, measures and mutates only that copy, then proves the original environment remains healthy.
- GitHub Actions run `32648822899` passed on commit `61759a0`: Ubuntu and macOS each ran `258 passed`; Windows ran
  `257 passed, 1 skipped` for its intentionally POSIX-only executable-bit check. Each system also passed first install,
  healthy repeat, repair gating, explicit repair, `pip check`, installed-runtime smoke checks, release documentation and
  byte-reproducible lock validation.
- Same-run logical footprint results with PySide6 `6.11.2`:

  - Linux: Essentials `635,202,444 bytes` (`605.78 MiB`), full `1,068,298,618 bytes` (`1,018.81 MiB`), saving
    `433,096,174 bytes` (`40.54%`).
  - macOS arm64: Essentials `638,141,562 bytes` (`608.58 MiB`), full `1,518,455,076 bytes` (`1,448.11 MiB`), saving
    `880,313,514 bytes` (`57.97%`).
  - Windows x64: Essentials `569,322,737 bytes` (`542.95 MiB`), full `1,021,148,587 bytes` (`973.84 MiB`), saving
    `451,825,850 bytes` (`44.25%`).
- CI evidence: `https://github.com/zhiyuzhang001-a11y/InstPlot/actions/runs/32648822899`.
- No real instrument file was available. Automated fixtures remain accepted, while real TXT/DAT/VSM/XLS/XLSX
  validation and a human desktop-session install remain pending rather than being inferred from CI.
