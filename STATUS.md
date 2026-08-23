# Current project status

- Date: `2026-08-23`
- Project state: `M3-M4 COMPLETE / M5 SIZE METRICS PENDING / M6 COMPLETE / M7 PLANNED`
- Active plan: `M7 release acceptance planning`
- Model Handoff Protocol: `paused`；文件保留，`MODEL_HANDOFF.md` 当前不参与普通工作路由。
- Last verified baseline: local `254 passed`；GitHub 原生矩阵 Ubuntu/macOS 各 `254 passed`，Windows
  `253 passed, 1 skipped`。三系统 clean CPython 3.12 Essentials 首次/重复/repair、`pip check`、
  TXT/XLSX/CSV/PNG/67 SVG 和 offscreen 绘图均通过。

## Overall goal

不制作独立 App 安装包；逐步完成数据正确性、差量撤销、GUI 流畅性、依赖体积优化，以及 Windows、
macOS、Linux 对应的一键隔离环境安装和启动。

## Completed

- M0 审计、M1 回归/依赖基线已完成。
- M2 数据 I/O 核心和自动化替代验收已接受；真实仪器文件验证延期到 M7。
- M3.1 行为和旧历史基线已接受：八处全量快照，10 步载荷 `640,005,280 bytes`。
- M3.2 已创建五类无 Qt 处理核心和 GUI 薄包装；参数、NaN/Inf、短数组和有限极值第一轮修复完成。
- M3.2 shared-conversion 全矩阵已闭环；本机/clean Python 3.12 wheel 均为 181 passed，`pip check`、
  offscreen 启动和八处 history owner 门禁通过。
- M3.3 差量历史纯核心完成：15 项新核心测试，本机及 clean installed-wheel 完整套件 196 passed；
  列、行、文件和组合命令支持精确 undo/redo，八处旧 GUI owner 尚未迁移。
- M3.4 GUI 历史迁移和 M3.5 验收完成：八处快照归零，redo/会话 reset/原子发布接入；正式基准
  `160,010,560 / 640,005,280 bytes = 25.00%`，十步往返通过。

## Current unresolved issues

1. GUI：大型 `PlotApp` 仍影响维护性，但 M4 的异步导入/拟合、交互重绘和诊断边界已闭环。
2. Footprint：macOS Essentials 已将 clean 环境从 1,487,064 降至 607,412 KiB（-59.15%）；
   Windows/Linux 功能兼容已通过，但各自完整 PySide6 对照体积尚未采集。
3. Installation：共享安装器和三系统入口已通过原生 CI 的首次、重复、repair 和安装后功能验证；
   Ubuntu 运行 PySide6 前需要宿主系统提供 `libEGL.so.1`（CI 使用 `libegl1`）；日志文件使用时间戳前缀
   和原子唯一分配，避免 Windows 低时钟分辨率下的名称碰撞。
4. Release：真实仪器文件、普通用户实机复核和最终用户文档尚未闭环。
5. Fonts：macOS 缺少 `SimHei` 时正式基准出现重复 Matplotlib 字体回退日志；改字体可能影响外观，
   留到 M7 配置化处理。
6. Packaging：许可证元数据已迁移到 SPDX `MIT` 和 `license-files`，warnings-as-error 构建通过。

## Stage order

`A M3.2 conversion → B M3.3 history core → C M3.4 GUI history → D M3.5 benchmark → E M4 GUI/performance → F M5 footprint → G M6 installers → H M7 release`

## Recording rule

每阶段结束同步更新 `docs/PROJECT_ROADMAP.md`、本文件和对应 `reports/M*-*.md`，记录修改范围、命令、
退出码、测试数量、耗时/内存/体积、残余风险、清理状态和唯一下一动作。后续阶段不自动获得实施授权。

## Next action

`制定并执行 M7.1 发布说明、锁文件可重建性和普通用户安装复核；真实仪器文件验收继续标记为待用户样本。`
