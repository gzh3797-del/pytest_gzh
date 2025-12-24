# 🚀 快速入手 My Project 

这是一个基于 **Python + Pytest + Selenium** 的混合自动化测试项目，涵盖了 **IOM项目AIAO精度测试**、**Web UI 业务流程**以及 **GUI 上位机交互**等的自动化测试。

项目采用 **PO (Page Object)** 设计模式与 **数据驱动 (Data-Driven)** 思想，旨在实现高效的回归测试与设备校准。

## 🛠 技术栈 (Tech Stack)

* **核心框架**: `Python 3.12+`, `Pytest`
* **通讯协议**: `Modbus RTU/TCP` (针对 IOM 模块)
* **UI 自动化**: `Selenium` / `Playwright` (Web端), `AutoHotkey` (GUI工具)
* **测试报告**: `Allure` (集成中)
* **依赖管理**: `pip`

---

## 📂 项目结构 (Project Structure)

```text
.
├── api/                    # 🔌 接口封装与基础页面
│   ├── modbus_connet.py    # Modbus 底层通讯连接封装
│   └── ui_base_page.py     # UI 自动化 BasePage 基类
├── common/                 # 🔨 公共方法与第三方资源
│   └── Source/             # 硬件/外部通讯协议资源 (如 CL3021)
├── config/                 # ⚙️ 环境配置读取脚本
│   └── modbus_config.py    # Modbus 参数配置加载
├── datas/                  # 📄 测试数据 (YAML/JSON) - 数据驱动层
├── operation/              # 🕹️ 业务操作层 (Operation Layer)
│   ├── IOM/                # 硬件模块操作逻辑
│   └── WEB2/               # Web 页面对象 (PO模式实现)
├── test_case/              # ✅ 测试用例层 (Test Cases)
│   ├── IOM/                # 硬件精度与功能测试
│   └── WEB2/               # Web 业务流程测试
├── tools/                  # 🧰 辅助工具箱
│   ├── Gui自动化升级/        # AHK 脚本：处理非标准 GUI 交互
│   └── modbus报文生成/       # 报文生成工具
├── config.json             # 🌍 全局环境配置文件
├── conftest.py             # ⚡ Pytest Fixture (前后置/钩子函数)
├── pytest.ini              # ⚙️ Pytest 核心配置文件
├── requirements.txt        # 📦 项目依赖清单
└── README.md               # 📘 项目说明文档

```

---

## 💻 环境搭建 (Environment Setup)

### 1. 前置要求

* 操作系统：Windows (推荐，因涉及 GUI 自动化) / Linux / macOS
* Python 版本：**3.12+** (必需)

### 2. 安装依赖

建议在虚拟环境中运行以下命令：

```bash
pip install -r requirements.txt
```

### 3. 项目配置

运行测试前，**必须**检查并配置根目录下的 `config.json` 文件：

* **硬件测试**：配置串口号 (COM Port)、波特率、设备地址。
* **Web 测试**：配置目标 URL、测试账号、数据库连接信息。

---

## ▶️ 运行指南 (Usage)

本项目支持通过 `pytest` 命令行结合 `marker` (标签) 运行不同模块的测试。

### 1. IOM 模块测试 (硬件精度)

| 测试场景 | 运行命令 | 备注 |
| --- | --- | --- |
| **AI 电压精度** | `pytest -m ai_v` | Analog Input Voltage |
| **AI 电流精度** | `pytest -m ai_c` | Analog Input Current |
| **AO 电压精度** | `pytest -m ao_v` | Analog Output Voltage |
| **AO 电流精度** | `pytest -m ao_c` | Analog Output Current |

### 2. Web2 业务测试 (UI 自动化)

| 测试场景 | 运行命令 | 备注           |
| --- | --- |--------------|
| **批量添加用户** | `pytest -m w_add` | 默认执行 20 个用户添加 |
| **清理用户数据** | `pytest -m w_delete` | 删除所有非admin用户 |
| **版本升级压测** | `pytest -m w_update` | 执行 100 次升级流程 |

### 3. 特殊工具运行

硬件校准脚本需单独作为 Python 脚本运行：

```bash
# 进入目录后选择对应方法执行
python operation/IOM/Calibration.py
```
---

## 📊 测试报告 (Reporting)

> 🚧 **当前状态**：Allure 报告集成中，暂未完全实装。

未来版本查看报告的标准方式：

1. **执行测试并生成数据**：
```bash
pytest --alluredir=./reports
```

2. **启动报告服务**：
```bash
allure serve ./reports
```



---

## ⚠️ 注意事项与常见问题 (FAQ)

1. **浏览器驱动 (Web Testing)**
* 请确保本地安装了与 Chrome 浏览器版本匹配的 `chromedriver`，并已配置到环境变量中。


2. **Web2 元素遮挡问题**
* **现象**：运行 `w_add` 添加用户时，偶发点击失败。
* **原因**：部分弹窗或 Toast 消息遮挡了提交按钮。
* **解决方案**：目前的临时方案是重新运行用例；后续将优化 `ui_base_page.py` 中的显式等待逻辑。


3. **数据清理**
* **Web 测试**：建议在每轮回归测试前手动运行 `pytest -m w_delete` 以保证环境纯净。
* **IOM 测试**：测试结束后请断开 Modbus 连接，以免占用串口。



---

**Maintainer:** [Zihan Gao/Team Test]