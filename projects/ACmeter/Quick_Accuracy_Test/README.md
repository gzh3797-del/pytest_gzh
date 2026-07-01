# 快速精度测试工具

AcuRev 系列电表快速精度测试独立 GUI 工具，不依赖 autotest 工程，开箱即用。

---

## 目录结构

```
快速精度测试/
├── precision_tool.py          # 主程序入口（双击或 python 运行）
├── requirements.txt           # Python 依赖
├── core/                      # 核心模块
│   ├── source_comm.py         # CL3021 交流源控制（UDP）
│   ├── meter_comm.py          # 电表 Modbus RTU 通信 + 精度计算
│   ├── test_engine.py         # 测试流程引擎
│   ├── addr_loader.py         # Modbus 地址表加载（含内置默认地址）
│   ├── testpoint_reader.py    # 测试点 Excel 读取
│   └── power_calc.py          # 功率计算（P/Q/S）
├── modbus_addr/               # Modbus 地址表 Excel 存放目录（可多个项目）
├── test_point/                # 测试点文件目录
│   └── acuvimseries_test_case.xlsx
├── results/                   # 测试结果自动输出目录
└── CL3021通信协议/             # 交流源通信协议文档（参考）
```

---

## 环境要求

- Python 3.10+
- 依赖安装：

```bash
pip install -r requirements.txt
```

---

## 硬件连接

| 设备 | 连接方式 | 默认参数 |
|------|---------|---------|
| AcuRev 电表 | RS-485 串口（RTU） | COM3，19200 baud，N81，Slave ID 1 |
| CL3021 交流源 | 网线（UDP） | IP 192.168.0.50，Port 10003 |
| 本机网口 | — | IP 192.168.0.51，本地 Port 10005 |

> 串口通信仅用于 CL3021 **维护口（改 IP）**，实时控源只走 UDP/网口。

---

## 启动方式

```bash
python precision_tool.py
```

加 `--debug` 可在控制台输出详细日志（包含每条 UDP 发送字节）：

```bash
python precision_tool.py --debug
```

---

## 使用流程

### 1. 连接电表
填写串口号和波特率，点击 **连接电表**，状态变绿即可。

### 2. 连接源
选择 **TCP/UDP** 模式，确认 IP / Port / 本地Port 后点击 **连接源**。

### 3. 加载 Modbus 地址表（可选）
- 工具启动时自动扫描 `modbus_addr/` 目录：目录中只有一个 `.xlsx` 时自动填入，多个时需手动选择。
- 若不加载，则使用内置默认地址（AcuRev1320）。
- 更换电表型号时，将对应地址表 Excel 放入 `modbus_addr/`，点 **浏览** 选中后点 **加载**。

### 4. 测试配置

| 配置项 | 说明 |
|--------|------|
| CT 类型 | 100mA / 333mV / RCT |
| 测点文件 | 从 `test_point/` 选取，列格式与 `acuvimseries_test_case.xlsx` 一致 |
| 接线方式 | 多选，勾选本次需要测试的接线方式 |
| 源稳定等待 | set_ac 命令发出后等待源稳定的秒数（默认 5s） |
| 采样间隔 | 每次读 Modbus 的间隔（20–500 ms） |
| 采样次数 | 每个测试点采样次数（默认 20 次） |

### 5. 换线提示
以下接线方式需要单独换线，测试到时会自动弹出确认对话框：

- **3E4WY**（三相四线星形）
- **2E3WD**（两元件三线三角）
- **3E4WY-P**

### 6. 查看结果
- **实时监控**面板（右侧）：每 200ms 刷新，显示当前值/最小/最大/平均及误差%
- **测试结果表格**（底部）：绿色行 = 全部 Pass，红色行 = 有 Failed
- **详细报告 Excel**：测试结束后自动保存至 `results/Precision_Measure_{接线方式}_{时间戳}.xlsx`

---

## 测试点文件格式

与参考脚本 `acuvimseries_test_case.xlsx` 格式一致，关键列（0-based）：

| 列索引 | 字段 | 说明 |
|--------|------|------|
| 0 | Case ID | 用例编号 |
| 1-3 | Ua / Ub / Uc | 相电压（V） |
| 4-6 | Ia / Ib / Ic | 相电流（A） |
| 7-12 | 相角 | 电压/电流相角（°） |
| 15 | 接线方式 | 如 `1E2w1p`、`3E4wY` |
| 16 | 频率 | 50 / 60 Hz |
| 17-22 | 精度阈值 | 电压/电流/相角/有功/无功/视在 |
| 23-24 | 采样次数/间隔 | 次数、间隔（秒） |

---

## 精度判定规则

误差公式：`(测量值 - 期望值) / 期望值`

**Pass 条件**：`|最小误差|`、`|最大误差|`、`|平均误差|` **三者同时** ≤ 精度阈值。

---

## 各接线方式说明

| 接线方式 | 有效测量量 | Excel 置空列 |
|---------|-----------|-------------|
| 1E2W1P | Ua、Ia、Pa、P/Q/S_sys | Uab/Ubc/Uca、In |
| 2E3W1P | Ua、Uc、Ia、Ic、Pa、Pc、Uca | Uab 测量列、Ubc 全部 |
| 2E3WD | Uab/Ubc/Uca、Ia/Ib/Ic、S_sys | 相电压、相角、各相功率 |
| 2E3WN | Ua、Uc、Ia、Ic、P/Q/S_sys | Ub_angle、Ib_angle、Uab 测量列、Ubc 全部 |
| 3E4WY | 全量 | — |
| 3E3WD | Uab/Ubc/Uca、Ia/Ib/Ic、P/Q/S_sys | — |
| 3E4WY-P | 全量 | — |
