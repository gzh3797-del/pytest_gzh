"""GUI 坐标标定助手(需解锁桌面后运行)。

Acuview 2 内容区是 Qt 自绘，UIA 不可见，需把 JSON 相对坐标对齐到屏幕物理像素。
本工具：
  1) 连接窗口，打印窗口矩形(物理像素)；
  2) 用当前(或命令行给定)的 content_origin 偏移，把指定页所有控件的"计划点击点"
     画在实时截图上 -> reports/plan_<page>.png，供肉眼核对是否压在控件上；
  3) 反复调整 --ox/--oy 直到点位对齐，然后把值写进 config.yaml 的
     gui_backend.content_origin (method: fixed, fixed_x/fixed_y)。

用法:
    python -m comm.ctl_acuview.calibrate --page Communication --ox 250 --oy 150
"""
from __future__ import annotations

import argparse

from . import dpi  # noqa: F401
from .config import get_config
from .gui_driver import AppDriver, is_session_locked


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--page", default="Communication")
    ap.add_argument("--ox", type=int, default=None, help="内容区原点相对窗口左上的 X 偏移")
    ap.add_argument("--oy", type=int, default=None, help="内容区原点相对窗口左上的 Y 偏移")
    ap.add_argument("--scroll", type=int, default=0, help="页面已向下滚动的像素")
    ap.add_argument("--config", "-c", default=None, help="项目配置 projects/<X>/config_acuview.yaml")
    args = ap.parse_args()

    if is_session_locked():
        raise SystemExit("会话锁屏：请先解锁桌面再标定。")

    cfg = get_config(args.config)
    # 临时覆盖 content_origin 以便试错
    if args.ox is not None and args.oy is not None:
        cfg.gui_backend.content_origin["method"] = "fixed"
        cfg.gui_backend.content_origin["fixed_x"] = args.ox
        cfg.gui_backend.content_origin["fixed_y"] = args.oy

    drv = AppDriver().launch_or_connect()
    rect = drv.window_rect()
    ox, oy = drv.content_origin()
    print(f"窗口矩形(物理px): {rect}")
    print(f"content_origin(屏幕px): ({ox},{oy})  "
          f"= 窗口左上({rect.left},{rect.top}) + 偏移({ox-rect.left},{oy-rect.top})")
    out = drv.plan_overlay(args.page, scroll_y=args.scroll)
    print(f"已生成计划点击点叠加图 -> {out}")
    print("请打开该图核对红/蓝点是否落在对应控件上；若整体偏移，调整 --ox/--oy 重跑。")
    print("对齐后把偏移写入 config.yaml: gui_backend.content_origin "
          "{method: fixed, fixed_x: <ox>, fixed_y: <oy>}")


if __name__ == "__main__":
    main()
