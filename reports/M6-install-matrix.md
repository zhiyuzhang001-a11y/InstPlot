# M6 install matrix

## M6.1 core and macOS workflow — 2026-08-23 — IMPLEMENTED

- Added a standard-library-only, dry-run-first installer with fixed states `missing`, `healthy`, `repair-needed`, and
  `conflict`. It validates project inputs, rejects a symlink/non-directory `.venv` and modified launcher, accepts only
  CPython 3.12, executes argument arrays without a shell, installs the hash lock and project into project-local
  `.venv`, runs `pip check`, and invokes the installed-runtime smoke test in Python isolated mode.
- Apply writes timestamped UTF-8 logs under `.install-logs`. Healthy reruns do not reinstall. An unhealthy environment
  exits nonzero unless the caller explicitly supplies `--repair`; failures retain the command output and log path.
  Launchers are exclusively created and never overwrite existing differing content.
- Added `install_windows.bat`, executable `install_macos.command`, and executable `install_linux.sh`. They use their
  own location rather than the current directory, forward `--repair`, never invoke sudo/PowerShell/curl, and launch
  the generated system-specific runner after successful installation. Missing Python/uv produces instructions and a
  nonzero exit rather than a silent system mutation.
- Failure-first installer tests initially failed because `scripts.install` and `scripts.verify_install` did not exist.
  Directed matrix now passes 13 tests, including Unicode/space paths, zero-write dry-run, Python version rejection,
  command-array construction, environment/launcher conflicts, dependency-health detection, readable failure logs,
  entrypoint safety and installed TXT/XLSX/CSV/PNG/SVG workflows.

## Real macOS installation matrix

- Source was copied to `/tmp/plotapp-m6-install.sEacgM/中文 安装 project`. The current working project remained free
  of `.venv`, generated launchers and `.install-logs` after dry-run.
- First apply selected uv-managed CPython 3.12, created `.venv`, installed `requirements.lock` with hashes, installed
  the project, passed `pip check`, created `run_instplot.command`, and passed isolated installed-runtime verification:
  TXT/XLSX import, CSV/XLSX export, plotting, PNG output and 67 SVG resources.
- A healthy second apply returned no actions. The first repeated check exposed two counterexamples which were fixed:
  standard macOS venv Python is legitimately a symlink, and a probe run from the project directory can import source
  instead of the installed wheel. Health/verification now uses `-I`; the confirmed module path is inside the venv
  `site-packages`.
- Removing chardet from the temporary environment changed dry-run to `repair-needed`. Apply without `--repair`
  returned exit 1 and a timestamped log. Explicit repair reinstalled the lock/project, reran verification and returned
  healthy; final `pip check` passed.
- Full local warnings-as-error suite passed `251 passed in 4.14s`; shell syntax, compilation and `git diff --check`
  passed. Windows and Linux wrappers have static/pure-core coverage but no native runtime execution yet, so M6 and the
  cross-platform part of M5 remain pending rather than being reported complete.

## Native CI matrix preparation — 2026-08-23 — READY / NOT YET RUN

- Added `.github/workflows/native-install.yml` for `ubuntu-latest`, `macos-latest`, and `windows-latest`, with read-only
  repository permission and explicit CPython 3.12. It uses the current official major examples
  `actions/checkout@v6` and `actions/setup-python@v6`.
- Each job runs the actual platform entrypoint with `INSTPLOT_INSTALL_ONLY=1`, repeats the healthy installation,
  removes chardet, proves an unapproved repair fails, repairs through the same entrypoint, then runs `pip check`, the
  isolated installed-runtime smoke and the full pytest suite. The environment flag only suppresses launching the GUI
  after installation; interactive user behavior is unchanged.
- Two failure-first tests froze the three runners, all entrypoints, repair, smoke, pytest, permissions and nonlaunching
  wrapper mode. Directed installer/CI suite passed 15 items; final local warnings-as-error suite passed
  `253 passed in 4.11s`, shell syntax, YAML parsing, compilation and `git diff --check` passed.
- The workflow is local only. The working tree contains the accumulated uncommitted project implementation, so no
  commit or push was made without explicit user authorization. Consequently this section does not claim Windows or
  Linux results yet.
