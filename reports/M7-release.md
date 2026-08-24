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
  validation remains pending rather than being inferred from CI.

## 2026-08-24 — native desktop follow-up

- Ran `INSTPLOT_INSTALL_ONLY=1 ./install_macos.command` against the existing healthy environment. It atomically created
  the missing `run_instplot.command`; the next installer dry-run reported `healthy`, an identical launcher and no action.
- Launched through `run_instplot.command` with the native Cocoa backend. The project `.venv` process remained alive and
  `~/Library/Logs/InstPlot/instplot.log` recorded `diagnostics_started` and `application_start`, closing the macOS desktop
  startup item without relying on the offscreen CI backend.
- The first native run exposed deprecated Qt 6 high-DPI attributes and repeated missing-`SimHei` font messages. Added two
  failure-first release tests, removed the Qt 5-era attributes (high DPI is always enabled in Qt 6), and filtered named
  font fallbacks against Matplotlib's installed font list while retaining a generic fallback. The repeated native launch
  produced no terminal warning output; local regression is `260 passed`.
- Repository history contains no instrument sample beyond the fixed parser `.xls` fixture, and the GitHub repository has
  no issue carrying an external sample. Real instrument validation therefore remains the only M7 user-supplied gate.
