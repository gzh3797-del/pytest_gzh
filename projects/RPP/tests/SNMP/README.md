# SNMP 测试模块（AcuHMI-1-7）

网关 SNMP 北向协议测试：覆盖 SNMP 页面配置（v2c / v3）、各设备型号 SNMP 数据与 Modbus 跨传输回读比对、Trap、端口/Community/认证等边界与安全场景。

> 本模块的 MIB 文件在**每次测试会话开始时从 HMI 页面自动下载**并解析，无需手动维护 `.MIB` 文件；被测设备的连接信息来自本目录 `devices.yaml`。

## 目录结构

```
SNMP/
├── conftest.py                # session fixtures：自动下载 MIB + 共享浏览器/登录 session
├── configure_snmp.py          # SNMP 页面配置（v2c/v3 参数、设备勾选、登录跳转）
├── mib_manager.py             # 从 HMI 下载 MIB→解压到 mib/→生成 mib_mapping.json
├── snmp_oid_map.py            # 运行时按 mib_mapping.json 动态加载 MIB，构建 OID↔参数映射
├── snmp_utils.py              # SnmpWalk.exe 封装 + pymodbus Modbus 读取工具
├── helpers_data.py            # 数据比对基类：勾选设备→walk→映射 Modbus→逐参数对比
├── helpers_ui.py              # UI 操作基类 SNMPBase（页面交互 / 断言）
├── devices.yaml               # 【本模块个性化配置】被测设备清单：ip/port/unit/model_type
├── SnmpWalk/                  # 第三方 SnmpWalk.exe（SNMP 读取工具，snmp_utils 调用）
├── mib/                       # 运行时生成：下载解压的 MIB 文件（自动创建）
├── mib_mapping.json           # 运行时生成：设备 Template→MIB 文件映射
└── test_Testcase_AcuHMI_SNMP_*.py   # 32 条用例（一函数一用例）
```

## 用例分组（共 32 条）

| 分组 | 文件前缀 | 数量 | 内容 |
|---|---|---|---|
| 设备数据比对 | `..._AcuRev4100 / AcuRev2100 / AcuRev1300 / Acuvim3 / AcuvimIIW / AcuvimIIR / AllDevices` | 7 | 勾选对应型号（支持同型多实例），walk 全量参数与 Modbus 逐参数对比 |
| v2c 端口 | `..._v2c_Port_001~004` | 4 | v2c 端口配置/边界 |
| v3 认证 | `..._v3_Auth_001~006` | 6 | v3 安全名称 / 认证 / 加密 |
| Trap | `..._Trap_001~005` + `..._TrapTarget_001` | 6 | Trap 目标配置与触发（部分需 AcuIOM/监听服务） |
| Community | `..._Community_001~002` | 2 | 读团体名 |
| Persistence | `..._Persistence_001~002` | 2 | 配置持久化 |
| 其它 | `..._Enable / BufferSize / HoldTime / PortBoundary / PortMismatch` | 5 | 使能开关、缓冲区、保持时间、端口边界/不匹配 |

## 前置条件

1. **网关可达**：SNMP 页面 URL / 账号来自 `projects/AcuHMI_1_7/settings.py`（`BASE_URL` / `DEFAULT_USERNAME` / `DEFAULT_PASSWORD`），回退默认见各脚本顶部。
2. **下游设备在线**：按 `devices.yaml` 列出的 ip/port/unit 挂好并可 Modbus 读取（数据比对用例的对比基准）。
3. **Python 依赖**：`playwright`、`pytest`、`pytest-playwright`、`pymodbus`、`pyyaml`。
4. **SnmpWalk.exe**：随模块自带于 `SnmpWalk/`（Windows），`snmp_utils.py` 直接调用，无需单独安装。
5. **MIB**：无需手动准备；会话开始时 `conftest.py` 的 `_mib_setup` 自动从 HMI 下载并生成映射；下载失败时回退已有 `mib/` 文件。

## 配置文件 devices.yaml

被测设备清单（数据比对用例据此勾选设备并读取 Modbus 基准）：

```yaml
devices:
  Acurev4100229:            # 实例名（对应 HMI 上的设备名）
    ip: 192.168.2.29
    port: 502
    unit: 203
    model_type: AcuRev4100  # 型号（同型多实例共享一个 model_type）
```

## 执行命令（仓库根目录）

```bash
# 走框架 run.py（注入报告目录，报告落 reports/AcuHMI_1_7/<时间戳>/）
python run.py AcuHMI_1_7 -k SNMP

# 直接 pytest
pytest projects/AcuHMI_1_7/tests/SNMP/ -v                 # 全部 SNMP 用例
pytest projects/AcuHMI_1_7/tests/SNMP/ -k "v3_Auth" -v    # 按分组过滤
pytest projects/AcuHMI_1_7/tests/SNMP/test_Testcase_AcuHMI_SNMP_AcuRev4100_001.py -v   # 单条用例
```

> UI 用例会打开浏览器（`headless=False`）并全 session 复用一次登录；数据比对用例通过 SnmpWalk 读 SNMP、pymodbus 读 Modbus，按容差 `TOLERANCE` 逐参数对比。
