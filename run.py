"""统一入口：python run.py <项目> [pytest 参数...]"""
from __future__ import annotations

import sys

from framework.runner.runner import run


def main() -> int:
    if len(sys.argv) < 2:
        print("用法：python run.py <项目> [pytest 参数...]")
        print("示例：python run.py AcuHMI_1_7 -m smoke")
        return 2
    project = sys.argv[1]
    pytest_args = sys.argv[2:]
    return run(project, pytest_args)


if __name__ == "__main__":
    raise SystemExit(main())
