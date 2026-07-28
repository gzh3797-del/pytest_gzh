# 直流表能量精度测试（XL-9600 控源 + AcuDC 精度测试）

XL-9600（星龙）直流电能表检定装置控源工具，带自动精度测试：
控源 → 等源稳定 → Modbus 读 AcuDC(320/300/260) 电压/电流/功率/能量 → 判精度 → 出 Excel 报告；
可选测能量脉冲误差（表格填「脉冲常数」即启用，脉冲常数自动同步设到 XL9600 和被检表）。

## 运行

```bat
pip install -r requirements.txt
python gui.py
```

- **手动控源页**：连接 XL9600（UDP，默认 192.168.1.105:24433）→ 参数配置 → 源输出（填实际 V/A，自动按额定基准换算成检定点%）→ 误差读取。
- **自动精度测试页**：读表设置（串口/网口，存回 config.json）→ 加载/编辑用例（input.xlsx，可页内改、导出）→ 开始测试 → 报告输出到 `result\`。单点快测可临时验一个点。

## 文件结构

| 文件 | 作用 |
|---|---|
| `xl9600.py` | XL-9600 UDP+GBK 协议驱动（线程安全；控源命令超时容忍，读命令严格） |
| `gui.py` | 主界面（tkinter，两标签页 + 共享日志） |
| `accuracy_engine.py` | 精度测试引擎：XL9600 适配成框架 Source、跑用例/单点、用例读写 |
| `accuracy_tab.py` | 自动精度测试标签页 |
| `fast_accuracy_test/` | 精度测试框架（配置/读表/精度算法/用例执行/报告） |
| `config.json` | 读表连接配置（rtu/tcp、型号、字序等） |
| `test_data/input.xlsx` | 测试用例表 |
| `test_concurrency.py` `test_accuracy_integration.py` | 离线测试（不需硬件）：`python test_accuracy_integration.py` |

## 依赖

Python ≥ 3.11，见 `requirements.txt`（pymodbus / openpyxl / pyserial；GUI 用标准库 tkinter）。

## 打包桌面 exe（可选）

```bat
pip install pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed --name "XL9600控源" ^
  --collect-submodules pymodbus --collect-submodules fast_accuracy_test ^
  --hidden-import serial --hidden-import accuracy_tab --hidden-import accuracy_engine gui.py
```
将 `dist\XL9600控源.exe` 与 `config.json`、`test_data\` 放同一目录使用。
