# AcuRev-100（ACmeter）测试项目

计费级 kWh 子表（内部代号 RACG），型号 AcuRev-101-mA / AcuRev-101-mV；通信仅 RS-485 + USB 双口 Modbus RTU（无北向协议），上位机 Acuview 2，无 Web 页面；接线仅 1E2W / 2E3W1P（A+C）/ 3E4WY。

- 单一配置源 `config.yaml`；素材 `data/`；spec 产物 `spec/`；用例执行证据 `reports/`；转化进度 `PROGRESS.md`（`tools/progress_stats.py` 生成，勿手改）
- 测试数据源 `tests/case_map.yaml` / `tests/case_map_wiring.yaml`：源设定/判据/门禁全在这里，**改判据只改 yaml 不改 .py**；`needs_review` 非空的用例运行时自动拒执行
  - 判据口径溯源见 `case_map.yaml` 顶部 `accuracy_spec` 块（SRS v1.03 §3.5 基础电参量测量）：改容差口径先改该块再同步测点 `lo/hi`
  - 测点判据键：`checks`（source_read=现值区间 / energy_accumulate=窗口 Δ 增量）、`checks_instant`（仅能量类，现值区间，用于频率等非累积量）
- 机制细节看代码注释与知识库：`tests/helpers_*.py`（控源/三级救源/A相护栏/保活退场）、`comm/ctl_acuview/`（共享引擎）、`knowledge/meters/AcuRev100/`（判据依据/决策留痕）

## 执行命令（均在仓库根目录）

> **测量类（002~011）前置**：CL3021 源在位、电表解铅封（拨背面 Dip Switch）、校验口 COM 与 `config.yaml > transport.rtu.port` 一致，不驱动界面。
> **GUI 类（009-GUI/014/015/016/017）前置**：Acuview2 已连表（Add Connection 行0=ACmeter/COM11，行1=ACmeter-USB/COM6）、Tesseract 已装、桌面不锁、CL3021 源保活；**运行期间禁止碰鼠标键盘**。
> 控制台编码与 pymodbus 日志级别由 `conftest.py` 处理，无需 `$env:PYTHONUTF8=1` / `--log-cli-level`；要看 Modbus 协议帧设 `$env:PYMODBUS_DEBUG=1`。
> `-k` 是子串匹配：`case1` 会连带 `case10~19`，选单条请用完整编号或写文件路径。

```powershell
# ── 单条用例（复跑/调试）────────────────────────────────────────────────
pytest projects/AcuRev100 -k test_002_01_case7                       # 路径不可省: testpaths=projects 会连带收集别的项目
pytest projects/AcuRev100 -k test_002_01_case7 -s                    # 加 -s 打印控源过程 [gear]/[guard]/[点N]

# ── 按模块 ─────────────────────────────────────────────────────────────
pytest projects/AcuRev100/tests/002_phase_voltage/ -v                # 002 相电压 7 条
pytest projects/AcuRev100/tests/003_line_voltage/ -v                 # 003 线电压 4 条
pytest projects/AcuRev100/tests/ -k "test_004_01_case8 or test_004_01_case9 or test_004_01_case10 or test_004_01_case11 or test_004_01_case12 or test_004_01_case13 or test_004_01_case14 or test_004_01_case15" -v   # 004 mA 原生 8 条
pytest projects/AcuRev100/tests/ -k "test_005_01_case7 or test_005_01_case8 or test_005_01_case9 or test_005_01_case10" -v   # 005 mA 原生 4 条
pytest projects/AcuRev100/tests/006_phase_angle/ -v                  # 006 相角 4 条
pytest projects/AcuRev100/tests/ -k "test_007_01_case5 or test_007_01_case6 or (test_007_01_case7 and not test_007_01_case7_01) or test_007_01_case8 or test_007_01_case9 or test_007_01_case10" -v   # 007 电能 mA 原生 6 条(增量断言,不清零)
pytest projects/AcuRev100/tests/ -k "test_009" -v                    # 009 接线验证 6 条
pytest projects/AcuRev100/tests/ -k "test_011" -v                    # 011 计时 7 条
pytest projects/AcuRev100/tests/014_firmware/ -k "case7 or case10" -v -s   # 014 只读/校验类, 不刷机
pytest projects/AcuRev100/tests/ -k "test_015 or test_016 or test_017 or test_009_03_case5" -v   # GUI 引擎类(含门禁 skip)
pytest projects/AcuRev100/tests/ -v -s                               # 全量(needs_review 未清空者运行时自动拒执行)

# ── 010 接线检查: 必须按台面实际接线选组(用例会改写 4162, 物理接线须与所选组一致)──
pytest projects/AcuRev100/tests/010_wiring_check/ -k "010_07_case1 or 010_08 or 010_12_case5" -v   # 1E2W 组 4 条
pytest projects/AcuRev100/tests/010_wiring_check/ -k "010_07_case2 or 010_09 or 010_12_case6 or 010_13_case2" -v   # 2E3W1P 组 10 条
pytest projects/AcuRev100/tests/010_wiring_check/ -k "not 010_07_case1 and not 010_07_case2 and not 010_08 and not 010_09 and not 010_12_case5 and not 010_12_case6 and not 010_13_case2" -v   # 3E4WY 组 38 条

# ── 红线操作(会重启/改写电表, 需现场值守)────────────────────────────────
pytest projects/AcuRev100/tests/014_firmware/test_acuview_Function_AC_meter_014_01_case6.py -v -s   # 刷机: 先置 config run.allow_firmware_upgrade=true, 跑完改回 false

# ── 工具与维护 ─────────────────────────────────────────────────────────
python -m comm.ctl_acuview.spec_loader --build --config projects/AcuRev100/config.yaml   # 改地址表/控件 JSON 后重建 spec
python -m comm.ctl_acuview.find_register --config projects/AcuRev100/config.yaml "frequency"   # 按关键词查寄存器
python -m comm.ctl_acuview.find_widget --config projects/AcuRev100/config.yaml "wiring"   # 按关键词查 Acuview 控件
python projects/AcuRev100/tools/source_diag.py                       # CL3021 控源逐动作排查(交互菜单, 人眼盯源输出)
cd projects/AcuRev100/tools/accuracy_pulse ; pip install -r requirements.txt ; python run_autotest.py   # 脉冲灯精度(独立 GUI)
cd projects/AcuRev100/tools/accuracy_quick ; pip install -r requirements.txt ; python precision_tool.py   # 快速精度(独立 GUI)
```

## 模块状态

| 模块 | 可自动化 | 说明 |
|:---:|:---:|---|
| 001 交流频率 | 0 | 转手工×4：源 40Hz 下发被拒、70Hz 输出掉 0 |
| 002 相电压 | 7 | — |
| 003 线电压 | 4 | — |
| 004 电流测量 | 8 | mV(333mV)×9 待接 5A/333mV CT 回路；80mA/RCT×2 待固件确认 CT 枚举 |
| 005 PF·功率 | 4 | case7 Ic 已降 15A（源带载上限） |
| 006 相角 | 4 | 0° 点用 ranges 判据处理 0~360 回绕 |
| 007 电能累计 | 6 | 增量断言（不写 CLEAR_ENERGY）；mV×4 / RCT×3 门禁 |
| 008 脉冲 | 0 | 走 tools/accuracy_pulse，待脉冲台架 |
| 009 接线验证 | 6 | 另 6 条门禁；009_03_case6（CT Primary 越界拒绝）已实测通过 |
| 010 接线检查 | 52/59 | 手工 7 条（Van<10V 自供电冲突）：010_08_case1、010_09_case1/2、010_10_case1/2/4/5 |
| 011 计时功能 | 7 | 门禁 5 条：case3/12 转手工（7 天长时）、case4/8/9 待铅封拨码 |
| 012 LCD/LED | 0 | 目视×5 |
| 013 铅封 | 0 | 人工拨码×5 |
| 014 Firmware | 11/14 | case5/12/15 存根 skip；SN 占位寄存器疑似产测问题 |
| 015 数据重置 | — | case3（恢复出厂）红线未授权 |
| 016 通讯模块 | — | 待专用 runner：SlaveID、case16 端口对调 |
| 017 密码 | — | 待专用 runner：密码三态流程 |

判据依据与裁决留痕见 `knowledge/meters/AcuRev100/`、`knowledge/shared/decisions.md` 及 case_map 各条 `decisions` 字段。
