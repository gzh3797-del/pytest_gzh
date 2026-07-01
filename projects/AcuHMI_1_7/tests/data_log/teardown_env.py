#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
teardown_env.py — 停止 setup_env.py 并清除标志文件

若 .setup_pid 存在，通过 PID 终止 setup_env.py 进程；
否则只清除 .setup_done 和 .setup_pid 标志文件。
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
_SETUP_FLAG = _HERE / ".setup_done"
_PID_FILE   = _HERE / ".setup_pid"


def main():
    if _PID_FILE.exists():
        try:
            pid = int(_PID_FILE.read_text(encoding="utf-8").strip())
            if sys.platform == "win32":
                import subprocess
                result = subprocess.run(
                    ["taskkill", "/F", "/PID", str(pid)],
                    capture_output=True, timeout=5,
                )
                if result.returncode == 0:
                    print(f"已终止 setup_env.py 进程（PID={pid}）")
                else:
                    print(f"终止 PID={pid} 失败（可能已退出）：{result.stderr.decode(errors='replace')}")
            else:
                import signal
                os.kill(pid, signal.SIGTERM)
                print(f"已发送 SIGTERM 给 PID={pid}")
        except Exception as e:
            print(f"终止 setup_env.py 失败：{e}")

    _SETUP_FLAG.unlink(missing_ok=True)
    _PID_FILE.unlink(missing_ok=True)
    print("标志文件已清除")


if __name__ == "__main__":
    main()
