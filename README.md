# InstPlot Python Legacy

> [!IMPORTANT]
> 这是 InstPlot 的原 Python 版源码与算法参考仓库，处于维护归档状态，
> **不再提供安装包，也不推荐普通用户安装**。需要直接下载安装的软件请使用
> [InstPlot Lite](https://github.com/zhiyuzhang001-a11y/InstPlot)。

[![Legacy version](https://img.shields.io/badge/version-1.0.0-blue.svg)](pyproject.toml)
[![Platforms](https://img.shields.io/badge/Windows%20%7C%20macOS%20%7C%20Linux-source--only-555555.svg)](.github/workflows/python-tests.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-555555.svg)](LICENSE)

## 保留这个仓库的原因

InstPlot 最初使用 Python、PySide6、Matplotlib、NumPy、Pandas 和 SciPy
开发，包含完整绘图界面、数据处理、曲线拟合和历史操作。Rust 原生的
InstPlot Lite 已接替普通用户版本；本仓库继续保留：

- 原始 Python 功能和算法行为；
- 数据导入、拟合、处理与撤销历史测试；
- M1–M7 开发计划、性能记录和安装阶段报告；
- 可用于解析器回归验证的真实仪器样本。

完整提交历史经过路径过滤保留，可以追溯旧版从初始提交开始的演进。

## 当前状态

- 维护方式：源码归档和必要缺陷修复。
- 发布方式：不构建 DMG、EXE、DEB 或其他预编译安装包。
- Python：支持 CPython 3.10–3.14。
- 数据样本：`tests/fixtures/real_samples/` 中的脱敏实验文件用于验证真实格式。
- 现行产品：[InstPlot Lite](https://github.com/zhiyuzhang001-a11y/InstPlot)。

## 开发者运行

本节只面向需要研究旧实现的开发者，需要自行准备兼容的 Python 环境。

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --require-hashes -r requirements.lock
python -m pip install --no-deps .
python -m InstPlot
```

Windows 激活环境时可使用 `.venv\Scripts\Activate.ps1`。Linux 图形环境仍需
系统提供 `libEGL.so.1`，Ubuntu / Debian 对应软件包为 `libegl1`。

运行测试：

```bash
python -m pytest -q
```

重新生成锁文件的规范命令：

```bash
uv pip compile pyproject.toml --python-version 3.10 --universal --generate-hashes --output-file requirements.lock
```

## 文档

- [历史实施计划](docs/IMPLEMENTATION_PLAN.md)
- [历史项目路线图](docs/PROJECT_ROADMAP.md)
- [M2 数据 I/O 合同](docs/M2_DATA_IO_CONTRACT.md)
- [M3–M7 计划与合同](docs/)
- [历史阶段报告](reports/)

问题请提交到
[InstPlot Python Legacy Issues](https://github.com/zhiyuzhang001-a11y/InstPlot-Python-Legacy/issues)。

本项目使用 [MIT License](LICENSE)。
