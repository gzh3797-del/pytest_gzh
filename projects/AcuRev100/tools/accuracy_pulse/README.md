# 脉冲灯精度测试工具（Python 版）

按测试点表格自动完成"控源 → 写电表接线方式 → 写电表脉冲常数 → 频率计测脉冲周期 → 回填表格"的一条龙测试。**所有设备/寄存器参数集中在 `config.yaml`，跑之前按实际电表改。**

---

## 一、首次使用

需要电脑装了 **Python 3.10+**。在本文件夹里打开命令行（终端），依次：

1. **装依赖**（只需一次）：
   ```
   pip install -r requirements.txt
   ```
   - USB 连频率计还需另装 **NI-VISA**（[官网下载](https://www.ni.com/en/support/downloads/drivers/download.ni-visa.html)）。
2. **查设备**（看有哪些 COM 口 / USB 资源，填进 config）：
   ```
   python list_devices.py
   ```
3. 用记事本 / VSCode 打开 **`config.yaml`**，按你的实际设备和电表填好（见下节）。
4. **运行程序**：
   ```
   python run_autotest.py
   ```

> 在文件夹地址栏输入 `cmd` 回车，即可在本目录打开命令行。

---

## 二、config.yaml 说明（跑前必改）

```yaml
source:                    # CL3021 交流源
  mode: serial            # serial(串口) 或 tcp(网口)
  com: COM3               # 串口时用；波特率 9600
  baud: 9600
  host: 192.168.0.50      # 网口时用
  port: 2404

counter:                   # SDG 2042X 频率计
  mode: usb               # usb 或 lan(网口)
  resource: auto          # USB 自动识别 Siglent 设备，无需手填那串 USB0::...
  host: 192.168.1.100     # 网口时用
  port: 5024

meter:                     # 被测电表 Modbus
  mode: rtu               # rtu(串口) 或 tcp(网口)
  com: COM5               # 串口时用
  baud: 9600
  slave: 1                # 从站号
  host: 192.168.1.10      # 网口时用
  port: 502
  pulse_constant:          # 脉冲常数寄存器（按电表手册改）
    register: 4198         # AcuRev-100(RACG) Energy 脉冲常数 0x1066
    dtype: uint32
    word_order: big
    scale: 1000           # 寄存器存"实际值×1000"时填 1000；否则填 1
  wiring:                  # 接线方式寄存器 + 映射（按电表手册改）
    register: 4162         # AcuRev-100(RACG) Service Configuration 0x1042
    dtype: uint16
    map:                  # 工作表名 -> 写入寄存器的值（AcuRev-100 仅三种接线）
      "1E2W_脉冲灯测点": 0
      "2E3W1P_脉冲灯测点": 1
      "3E4WY_脉冲灯测点": 2

test:                      # 测试参数
  settle_s: 3             # 每点控源后稳定时间(秒)
  n_periods: 10           # 采集周期数（采集时长 = N × 脉冲周期）
  freq: 50                # 输出频率(Hz)
  pf_lagging: true        # 功率因数性质：true=感性(滞后)，false=容性(超前)
```

> **不同电表寄存器地址/格式/映射值不同，务必以该电表 Modbus 手册为准。**
> 本工具已按 **AcuRev-100（内部代号 RACG）** 适配：Energy 脉冲常数寄存器 4198(0x1066, uint32, ×1000)；
> 接线方式寄存器 4162(0x1042, uint16)，值 `0=1E2W(单相), 1=2E3W1P(分相 A+C), 2=3E4WY(三相星形)`——仅此三种。

---

## 三、操作步骤

1. 命令行运行 **`python run_autotest.py`** 打开程序。
2. **配置**：确认已加载 `config.yaml`（状态绿色）。改过配置后点"重新加载"。
3. **选 xlsx**：点浏览选测试点表格；下方"工作表"列表会列出可跑的接线测点页。
4. **多选工作表**：按住 Ctrl / Shift **可多选**要跑的接线方式（会按顺序一个个跑）。
5. **CT类型**：选 `mV` / `mA` / `全部`。
6. 点 **开始测试**。每种接线方式开始前：
   - 程序**自动往电表写该接线方式的寄存器值**；
   - **弹窗提示**"接线方式已设为【XXX】，请确认/改好实际接线后点确定"——
     不用动实际线的直接点确定；需要改线的改好再点确定；点取消则停止。
   - 然后跑该接线的所有点（源正常输出 → 写脉冲常数 → 频率计取周期 → 回填）。
7. 全部跑完，结果写回各工作表的 `mim(s)/max(s)/avg(s)`；误差列由表格公式自动算。

结果表格里每行带"接线"列，区分不同接线方式的数据。中途可点 **停止**。

---

## 四、xlsx 表格要求

- 文件 **.xlsx**；表头含 `Voltage / Current / Power Factor / Pulse Constant`（读取）和 `mim(s)(或min(s)) / max(s) / avg(s)`（回填）。
- 支持多 sheet（按接线方式）、合并的 CT类型/接线方式单元格（自动向下填充）、电流"标签列+数值列"。
- 允许表头上方有合并标题行；误差列公式不会被改动。

---

## 五、安全提示

⚠️ CL3021 会按测试点真实输出电压/电流（可达 400V）。开始测试前务必确认接线与负载正确，建议先用低电压点验证。

---

## 六、常见问题

- **COM 口 / USB 资源不知道填什么**：命令行运行 `python list_devices.py` 查看当前可用设备。
- **CL3021 连不上**：网口确认同网段+端口 2404；串口确认 COM 口与波特率（实测 9600）。
- **频率计 USB 扫不到**：装 NI-VISA、用数据线、必要时重插；config 里 `resource: auto` 会自动识别。
- **电表脉冲常数/接线方式没改对**：核对寄存器地址、数据格式(16/32位)、字节序、倍率、映射值与电表手册；从站号是否正确。
- **卡在"连接…/采集中"**：状态栏会显示当前阶段；采集时长 = N × 脉冲周期，周期长时单点耗时数十秒属正常。
- **某行状态报错**：多为该点无脉冲（检查源输出、电表脉冲常数、频率计接线与触发电平）；出错行不影响其余行。

---

## 七、文件说明

| 文件 | 作用 | 运行方式 |
|---|---|---|
| `run_autotest.py` | 主程序 | `python run_autotest.py` |
| `list_devices.py` | 列出可用 COM 口 / USB 资源 | `python list_devices.py` |
| `config.yaml` | 设备/寄存器/映射/参数配置（跑前改） | 用记事本编辑 |
| `requirements.txt` | Python 依赖清单 | `pip install -r requirements.txt` |
| `src\` | 程序模块源码 | — |

*完整源码仓库：`C:\Users\RenjieJiao\sdg2042x_counter`*
