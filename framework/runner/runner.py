"""把项目名解析为 pytest 调用，并把报告导向 reports/<项目>/[<模块>/]<时间戳>/。"""
from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path

import pytest

from framework.report.paths import detect_module, make_report_dir, update_latest

_REPO_ROOT = Path(__file__).resolve().parents[2]


def run(project: str, pytest_args: list[str]) -> int:
    project_dir = _REPO_ROOT / "projects" / project
    if not project_dir.is_dir():
        raise SystemExit(f"未知项目：{project}（缺少 projects/{project}/）")

    run_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    # 从 pytest 参数中识别 tests/ 下的单一模块目录；整项目/跨模块运行时为 None（报告退化为不含模块层级）。
    module = detect_module(pytest_args)
    dirs = make_report_dir(project, run_ts, module=module)

    os.environ["PROJECT"] = project
    os.environ["RUN_TS"] = run_ts
    os.environ["REPORT_DIR"] = str(dirs.root)
    os.environ["SCREENSHOT_DIR"] = str(dirs.screenshots)

    args = [
        str(project_dir),
        f"--html={dirs.root / 'report.html'}",
        "--self-contained-html",
        *pytest_args,
    ]
    exit_code = pytest.main(args)
    update_latest(dirs.root)
    return int(exit_code)
