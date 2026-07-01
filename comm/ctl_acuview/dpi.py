"""DPI 感知统一。

必须在导入 pyautogui / pywinauto 之前调用 ensure_dpi_aware()，
让进程工作在物理像素空间(1920x1080)，使截图、点击、窗口矩形坐标三者一致，
避免 125% 缩放下坐标错位。
"""
from __future__ import annotations

import ctypes

_done = False


def ensure_dpi_aware() -> None:
    global _done
    if _done:
        return
    try:
        # PER_MONITOR_AWARE_V2 = -4
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
    except Exception:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PER_MONITOR_AWARE
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass
    _done = True


# 导入即生效
ensure_dpi_aware()
