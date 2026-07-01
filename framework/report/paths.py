"""报告目录生成与 latest 指针维护。"""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[2]
_DEFAULT_REPORTS_ROOT = _REPO_ROOT / "reports"

# 从测试路径参数中识别 tests/ 下一级模块目录名（如 projects/AcuHMI_1_7/tests/BacnetIP → BacnetIP）。
_MODULE_PATH_RE = re.compile(r"[\\/]tests[\\/]([^\\/]+)")


@dataclass(frozen=True)
class ReportDirs:
    root: Path
    screenshots: Path


def detect_module(args: object) -> str | None:
    """从测试路径参数中唯一识别 `tests/` 下一级模块目录名；无法唯一确定时返回 None。

    `args` 为命令行/pytest 路径参数序列（支持 `path::node` 形式）。仅当所有参数指向同一个
    模块目录时才返回该名称；跨模块、整项目运行或直接指向 tests/ 下的文件时返回 None，
    使报告目录退化为不含模块层级的旧结构（向后兼容）。
    """
    found: set[str] = set()
    for arg in args or ():
        path_part = str(arg).split("::", 1)[0].replace("\\", "/")
        match = _MODULE_PATH_RE.search(path_part)
        if match:
            name = match.group(1)
            if name.endswith(".py"):  # 直接指向 tests/ 下的文件，无中间模块目录
                continue
            found.add(name)
    return next(iter(found)) if len(found) == 1 else None


def make_report_dir(project: str, run_ts: str, *,
                    module: str | None = None,
                    reports_root: Path | None = None) -> ReportDirs:
    base = reports_root or _DEFAULT_REPORTS_ROOT
    base = base / project
    if module:
        base = base / module
    base = base / run_ts
    # report.html 直接落在 run 根目录（pytest-html --self-contained-html 为单文件）；
    # 失败截图归到 screenshots/ 子目录。不再建 logs/（无任何代码写入，恒为空）。
    dirs = ReportDirs(root=base, screenshots=base / "screenshots")
    for path in (dirs.root, dirs.screenshots):
        path.mkdir(parents=True, exist_ok=True)
    return dirs


def update_latest(run_dir: Path, *, reports_root: Path | None = None) -> None:
    """优先建软链 reports/latest；失败（Windows 权限）退化为 latest.txt。"""
    root = reports_root or _DEFAULT_REPORTS_ROOT
    latest = root / "latest"
    try:
        if latest.exists() or latest.is_symlink():
            latest.unlink()
        latest.symlink_to(run_dir, target_is_directory=True)
    except OSError:
        (root / "latest.txt").write_text(str(run_dir), encoding="utf-8")
