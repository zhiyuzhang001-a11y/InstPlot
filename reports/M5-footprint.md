# M5 footprint record

## macOS dependency reduction — 2026-08-23 — IMPLEMENTED

- Baseline clean CPython 3.12.14 environment with full PySide6: `1,487,064 KiB`. Distribution-owned file totals:
  PySide6 meta/tools `4,888,256`, Addons `885,118,755`, Essentials `342,851,134`, and shiboken6 `1,314,088`
  bytes. QtAwesome depends on QtPy and does not require the full PySide6 meta package.
- A failure-before-metadata-change probe installed only PySide6-Essentials and all other locked runtime libraries.
  The current wheel correctly failed `pip check` only because its old metadata still named PySide6, while the full
  external suite passed `238 passed in 49.87s`. This proved runtime coverage before changing dependency policy.
- Changed project metadata to `PySide6-Essentials>=6.6,<7` and regenerated the universal Python 3.12 hash lock.
  The final locked environment contains no PySide6 meta package or Addons. Final wheel install passed `pip check`,
  `238 passed in 4.12s`, QtCore/Gui/Widgets/Svg imports, QtAwesome, and offscreen `PlotApp` startup.
- Post-import environment size is `607,412 KiB`, a `59.15%` reduction from the comparable full environment. Three-
  process median startup max RSS changed from `351,797,248` to `347,979,776 bytes` (`1.09%`), confirming the disk
  saving does not imply an equivalent runtime-memory saving.
- Direct dependency audit: chardet handles text encoding; openpyxl/xlrd cover xlsx/xls; NumPy/Pandas/SciPy implement
  data and fitting; Matplotlib provides plots; QtAwesome provides toolbar icons. Pillow remains transitive through
  Matplotlib. The 154,833-byte wheel contains nine modules and 67 SVG resources, not README image assets.
- Platform gate: macOS implementation is accepted. Docker CLI was present but its daemon was unavailable; no Linux
  runtime evidence was claimed. Windows and Linux Essentials validation remains coupled to M6 installers/CI, so M5
  is not yet marked fully complete.
