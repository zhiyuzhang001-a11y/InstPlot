# InstPlot - 实验数据可视化工具

<div align="center">

![InstPlot Logo](logo.ico)

**简单、便捷的科研数据绘图软件**

[![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](#)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)](#)

</div>

---

## 📖 简介

**InstPlot** 主要用于实验数据的即时观察和预处理，具有操作简单和响应快速的特点。

### ✨ 核心特性

- 📊 **多数据源导入** - 支持 TXT、CSV、DAT（含自动识别的 VSM 数据）、XLS 和 XLSX
- 🎯 **智能数据处理** - 内置对称处理、归一化、背景去除等功能
- 🖱️ **交互式操作** - 缩放、平移、单点选择和矩形选择
- 💾 **图片导出** - 支持 PNG、JPG、SVG、PDF 等多种图片格式
- 📊 **数据导出** - 支持 CSV、TXT、Excel 等格式
---

## 📚 主要功能

### 1️⃣ 数据导入

**导入方式**：
- 点击工具栏 **"打开文件"** 按钮
- 直接**拖拽文件**到软件窗口

### 2️⃣ 数据可视化

**轴选择**：
- 通过下拉菜单选择 X 轴和 Y 轴数据列，然后点击绘制曲线
- 支持中文列名和 LaTeX 格式

**图表类型**：
- 点线图
- 自动图例生成
- 网格线显示

### 3️⃣ 数据处理功能

#### 🔄 对称处理
将数据沿 Y 轴镜像对称。当我们同时处理多条曲线，其电压或电阻值差别较大，在同一个图形中无法看到它们的形状，”对称处理“可以一键解决

<table>
<tr>
<td width="50%">
<b>处理前</b><br>
<img src="ReadMe图片/bf_对称处理.png" alt="对称处理前" width="100%">
</td>
<td width="50%">
<b>处理后</b><br>
<img src="ReadMe图片/af_对称处理.png" alt="对称处理后" width="100%">
</td>
</tr>
</table>

#### 📏 归一化
将数据缩放到 [0, 1] 范围，便于比较不同曲线横坐标上的差别，例如矫顽场等

<table>
<tr>
<td width="50%">
<b>处理前</b><br>
<img src="ReadMe图片/bf_归一化.png" alt="归一化前" width="100%">
</td>
<td width="50%">
<b>处理后</b><br>
<img src="ReadMe图片/af_归一化.png" alt="归一化后" width="100%">
</td>
</tr>
</table>

#### 🧹 去背底
选择合适的背底区间，对于VSM, 可以取这个范围，注意只需要红色方框中x轴对应的最小值和最大值，图中对应大概为6000，8500，可以移动鼠标观察左下角坐标位置。

<table>
<tr>
<td width="50%">
<b>处理前</b><br>
<img src="ReadMe图片/bf_去背底.png" alt="去背底前" width="100%">
</td>
<td width="50%">
<b>处理后</b><br>
<img src="ReadMe图片/af_去背底.png" alt="去背底后" width="100%">
</td>
</tr>
</table>
对于SMR，也需要找到合适的线性背底，可以找到曲线的两个最低点，大约是96，273。第一次去背底有时候并不完美，原因是最低点没有找准。可以进行第二次去背底，换个点，例如最高点也可以，图中对应4.2, 362.
<table>
<tr>
<td width="33%">
<b>处理前</b><br>
<img src="ReadMe图片/bf_去背底smr.png" alt="去背底前" width="100%">
</td>
<td width="33%">
<b>第一次处理</b><br>
<img src="ReadMe图片/af_去背底smr1.png" alt="去背底后" width="100%">
</td>
<td width="33%">
<b>第二次处理</b><br>
<img src="ReadMe图片/af_去背底smr2.png" alt="去背底后" width="100%">
</td>
</tr>
</table>

#### 🗑️ Remove 功能
快速删除多余的点或实验记录中的跳点，可以鼠标左键单击需要删除的点也可以画矩形框同时删去多个点。

<table>
<tr>
<td width="50%">
<b>删除前</b><br>
<img src="ReadMe图片/bf_remove.png" alt="删除前" width="100%">
</td>
<td width="50%">
<b>删除后</b><br>
<img src="ReadMe图片/af_remove.png" alt="删除后" width="100%">
</td>
</tr>
</table>

### 4️⃣ 交互式操作

**鼠标操作**：
- **左键单击** 高亮显示坐标或选择单个要删除的点
- **左键拖拽** - 拉出矩形选择一个或多个要删除的点
- **滚轮** - 缩放图表
- **右键拖拽** - 移动视图

---

## 🖥️ 系统要求

- **Python**：仅支持 CPython 3.12；3.11 和 3.13 尚未进入发布矩阵。
- **操作系统**：Windows、macOS 或带桌面环境的 Linux。自动化发布矩阵使用 GitHub 当前的
  `windows-latest`、`macos-latest` 和 Ubuntu `ubuntu-latest`；其他版本应以 Qt for Python 当前支持范围为准。
- **内存**：至少 2 GB；处理大型文件时建议 4 GB 或更多。
- **磁盘**：项目所在磁盘至少 1 GB 可用空间；依赖安装在项目内 `.venv`。
- **Linux 图形库**：系统必须提供 `libEGL.so.1`。Ubuntu/Debian 可在安装 InstPlot 前由管理员运行
  `sudo apt-get install libegl1`；其他发行版请安装提供同名库的系统包。InstPlot 安装器不会自行提权。

---

## ⚡ 一键安装与启动

下载或解压项目后，根据系统运行对应入口。安装器只创建项目内 `.venv`，不会修改系统 Python，也不会
自动下载 Python、uv 或系统软件。Linux 的图形库前置条件需由用户或管理员提前完成。

- Windows：双击 `install_windows.bat`
- macOS：双击 `install_macos.command`；若下载后没有执行权限，先运行
  `chmod +x install_macos.command`
- Linux：运行 `./install_linux.sh`；必要时先运行 `chmod +x install_linux.sh`

安装成功后会自动启动。以后可直接运行生成的 `run_instplot.bat`、`run_instplot.command` 或
`run_instplot.sh`。修复已有环境时，在终端运行对应安装入口并增加 `--repair`。

安装器需要 Python 3.12 或现有的 uv。如果两者都不存在，它会停止并显示安装指引，不会自动执行远程
脚本。安装日志位于项目的 `.install-logs`。仅检查计划而不写入时可运行：

```bash
python scripts/install.py --json
```

## 手动安装

### 1. 安装依赖

```bash
# 克隆或下载本项目
git clone https://github.com/zhiyuzhang001-a11y/InstPlot.git
cd InstPlot

# macOS / Linux：创建隔离环境（避免依赖当前系统 Python）
python3.12 -m venv .venv

source .venv/bin/activate

# Windows PowerShell：创建并激活隔离环境
# py -3.12 -m venv .venv
# .venv\Scripts\Activate.ps1

# 按 `requirements.lock` 的固定版本安装运行依赖，再安装本项目
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install --no-deps .
```

`requirements.lock` 是以 Python 3.12 生成、包含哈希的通用锁定文件。使用经过发布验证的 uv 版本
重新生成时运行：

```bash
uv pip compile pyproject.toml --python-version 3.12 --universal --generate-hashes --output-file requirements.lock
```

提交锁文件前必须运行 `python scripts/verify_release.py --check-lock`，并在 Windows、macOS、Linux
安装矩阵中验证。

### 2. 运行程序

```bash
# 直接运行主程序
python -m InstPlot
```



## 📥 获取更新

### 从 GitHub 拉取最新版本

```bash
# 进入项目目录
cd InstPlot

# 获取最新代码
git pull origin main

# 更新锁定依赖与程序
python -m pip install -r requirements.txt
python -m pip install --no-deps .
```

## 诊断日志

程序会保存轮转诊断日志，单个文件最多 1MB，并保留 3 个备份。未捕获错误会显示“复制错误详情”按钮，
便于远程排查；日志不会主动写入导入表格的行内容。

- macOS：`~/Library/Logs/InstPlot/instplot.log`
- Windows：`%LOCALAPPDATA%\InstPlot\Logs\instplot.log`
- Linux：`$XDG_STATE_HOME/instplot/instplot.log`，未设置时为 `~/.local/state/instplot/instplot.log`
- 自定义位置：启动前设置环境变量 `INSTPLOT_LOG_DIR`

## 安装故障排查

- 安装器报告 `repair-needed`：重新运行系统对应入口并增加 `--repair`。安装器不会在未授权时重装环境。
- 安装器报告 `conflict`：不要删除或覆盖提示的 `.venv`/启动器；先备份用户修改，再人工处理冲突。
- Linux 报告找不到 `libEGL.so.1`：先安装发行版提供的 EGL 运行库；Ubuntu/Debian 包名通常为 `libegl1`。
- 找不到 Python：安装 CPython 3.12 或 uv 后重试；不要使用项目未验证的 Python 3.11/3.13。
- 仍然失败：保留 `.install-logs` 中最新安装日志和上方“诊断日志”对应的平台日志，在 GitHub Issue
  中附上系统版本、复现步骤和错误文本；不要上传包含敏感数据的原始实验文件。

## 当前限制

- 自动化样本已经覆盖 TXT、CSV、DAT、XLS 和 XLSX 的主要解析边界，但尚未获得用户的真实仪器
  TXT/DAT/VSM/XLS/XLSX 文件做最终对照，因此该项状态为 `PENDING_USER_VALIDATION`。
- VSM 当前指可从内容识别的 VSM 风格 `.dat` 文件，不代表支持任意 `.vsm` 扩展名或厂商私有变体。
- Linux 不生成 `.desktop` 文件；Windows/macOS/Linux 都通过项目目录中的启动脚本运行。
- 本项目不提供 `.app`、DMG、MSI、独立 EXE 或 AppImage。

---

## 🙏 致谢

感谢以下开源项目的支持：
- [Python](https://www.python.org/)
- [PySide6](https://www.qt.io/qt-for-python)
- [Matplotlib](https://matplotlib.org/)
- [NumPy](https://numpy.org/)
- [Pandas](https://pandas.pydata.org/)
- [SciPy](https://scipy.org/)

## 🤝 贡献指南

欢迎提交 Issue 和 Pull Request！

### 报告 Bug
- 请在 GitHub Issues 中详细描述问题
- 包含重现步骤和错误信息
- 附加截图或数据文件（如适用）

### 提交功能建议
- 在 Issues 中开启讨论
- 描述功能的用途和预期表现
- 提供原型或示例

### 代码贡献
1. Fork 本仓库
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

---

## 📞 联系方式

- 🐛 Bug 反馈与功能建议：[GitHub Issues](https://github.com/zhiyuzhang001-a11y/InstPlot/issues)

---

<div align="center">

**⭐ 如果觉得软件好用，欢迎推荐给您的同事和朋友！⭐**

</div>
