# 测试组自动化工程

## 运行任意项目（从仓库根执行）
    python run.py <项目> [pytest 参数]
    例：python run.py acuhmi_1_7            # 跑该项目全部用例
        python run.py acuhmi_1_7 -m smoke   # 只跑 smoke 标记

run.py 解析项目名 → 注入报告目录环境变量 → 委托 framework/runner 调用 pytest。

## 配置
- `configs/global.yaml`：全局默认环境（交流源 / 继电器 / 本地 PC / SSH 等）
- `projects/<项目>/config.yaml`：项目特有项与对全局的覆盖；显示名写在本文件的 `project_name`
- 敏感值：复制 `configs/.env.example` 为 `configs/.env` 填写（口令、证书路径等，不入库）
- 合并优先级：configs/global.yaml ← projects/<项目>/config.yaml ← .env
- 旧式项目与 `comm/` 通过根 `modbus_config.py` 取连接参数，同样读 `configs/global.yaml`（单一配置源，ssh 密码走 `configs/.env`）

## 报告（reports/）
- 每次运行落 `reports/<项目>/<时间戳>/`，含 `html/`（pytest-html 自包含报告）、`screenshots/`（仅失败截图）、`logs/`
- 最近一次：本机因系统禁用符号链接（WinError 1314），用 `reports/latest.txt` 记录最近一次目录路径；在允许符号链接的机器上则为 `reports/latest` 软链
- `reports/` 整体 git 忽略

## 目录结构

> 分工：`framework/` 框架基建（产品无关）；`comm/` 通信与台架库（协议客户端 / 源 / 继电器 / SSH，**直接复用原工程**、跨项目共用）；`projects/<项目>/` 各测试项目。

```
<仓库根>/
├── run.py                  # 根入口：python run.py <项目> [pytest 参数]，委托 framework/runner
├── conftest.py             # 仅跨项目通用 fixture（计时 / 异常兜底 / Modbus 连接）
├── pytest.ini              # testpaths 指向 projects/
├── requirements.txt
├── CLAUDE.md               # 给 Claude 的项目指令与知识库导航
├── README.md               # 本文件：怎么跑 / 配置在哪 / 报告在哪
│
├── .claude/                # Claude Code 配置
│   ├── agents/             #   专职子智能体定义
│   ├── skills/             #   斜杠命令技能
│   ├── settings.json       #   团队共享设置（入库）
│   └── settings.local.json #   本地个人设置（不入库）
│
├── framework/              # 测试基建 · 产品无关
│   ├── config/             #   分层配置加载器（configs/global.yaml ← projects/<项目>/config.yaml ← .env）
│   ├── logging/            #   统一日志
│   ├── report/             #   报告 / 截图归一管理
│   ├── retry/              #   重试策略
│   ├── assertion/          #   断言增强
│   └── runner/             #   被 run.py 调用的执行编排
│
├── configs/                # 配置中心
│   ├── global.yaml         #   全局默认环境（交流源 / 继电器 / 本地 PC / SSH 等）
│   └── .env.example        #   敏感值模板（复制为 configs/.env 填写，不入库）
│
├── projects/               # 所有测试项目
│   ├── acuhmi_1_7/         # ← 新结构参考样板（config.yaml + settings.py + framework 配置）
│   │   ├── README.md       #     模块索引（详见各模块 README，如 tests/bacnet/）
│   │   ├── config.yaml     #     项目配置：显示名 project_name、设备 IP 等本项目差异
│   │   ├── conftest.py     #     项目级 fixture
│   │   ├── tests/          #     test_*.py，按模块/协议分子目录（ui/ bacnet/ wiring_check/ ...）
│   │   ├── pages/  helpers/  data/
│   └── <其它旧项目>/       # 从旧 test_case 拷入，待各 owner 迁移（仍用 from comm... 老式导入）
│                           #   ACM_41_WEB2 / AcuRev4100 / AcuRev1320 / AcuDC_300 / AcuDC_320 /
│                           #   AcuDc260 / Acuvim2v3 / AcuvimSeries / AcuvimⅡ / IOM / WEB2_4100 ...
│
├── comm/                   # 通信与台架库（原工程直接复用；import 形如 from comm.xxx）
│   ├── modbus_rtu_tcp.py   #   Modbus RTU/TCP 客户端
│   ├── modbus_get_attr.py / modbus_set_attr.py   #   电表读写
│   ├── source_control.py / Source_control_3021dc.py   #   交流/直流源控制
│   ├── device_reboot.py    #   继电器上下电
│   ├── ssh_cmd.py / dev_uart.py / multi_threads.py / ui_base_page.py
│   └── QT_comm/            #   QT 仪器自动化
├── modbus_config.py        # comm 连接参数（读 configs/global.yaml；ssh 密码走 configs/.env）
│
├── tools/                  # 工具脚本集合（原 Convenient_tools/ 已并入）
│   ├── Protocols/          #   免适配协议比对引擎（从仓库根执行 python tools/Protocols/...；由协议后端工程师维护）
│   └── modbus_msg_gen/     #   原 modbus报文生成
│
├── ci/                     # CI 预留：jenkins/ github_actions/ docker/（按需创建）
├── knowledge/              # 知识库
│
├── reports/                # 报告中心（整体 git 忽略）
│   ├── latest              #   软链 → 最近一次（本机禁用软链时退化为 reports/latest.txt）
│   └── <项目>/<时间戳>/{html, screenshots, logs}
│
└── docs/superpowers/       # 工程文档：MIGRATION_CHECKLIST.md（迁移/适配指南，入库）
```

## 新增/迁移项目
见 `docs/superpowers/MIGRATION_CHECKLIST.md`。已迁移参考样板：`projects/acuhmi_1_7/`。
