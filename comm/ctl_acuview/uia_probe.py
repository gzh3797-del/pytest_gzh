"""探测 Acuview 2 控件树，决定 GUI 后端。

输出:
  reports/uia_dump.txt   - 控件树文本(带 control_type / name / automation_id / rect)
  reports/uia_summary.json - 按控件类型统计 + 是否暴露可操作控件的结论

结论用于 gui_driver 选择 UIA 还是 坐标+视觉。
Qt5 程序对 UIA 暴露程度不一：若能看到大量 Edit/ComboBox/TreeItem 则 UIA 可用；
若主窗口下几乎只有 1 个 Pane/Custom，则必须走坐标+视觉兜底。
"""
from __future__ import annotations

import json
import sys
import time
from collections import Counter
from pathlib import Path

from .config import get_config


def _connect_or_launch(cfg):
    from pywinauto import Application
    title_re = cfg.app.window_title_re
    # 先尝试连接已运行实例
    try:
        app = Application(backend="uia").connect(title_re=title_re, timeout=3)
        print("[probe] 已连接到正在运行的 Acuview 2")
        return app
    except Exception:
        pass
    print(f"[probe] 启动 {cfg.app.exe_path} ...")
    app = Application(backend="uia").start(f'"{cfg.app.exe_path}"', timeout=10)
    time.sleep(cfg.app.launch_wait_s)
    # 启动器进程可能与主窗口进程不同，改用 connect 抓主窗口
    app = Application(backend="uia").connect(title_re=title_re, timeout=20)
    return app


def _walk(elem, depth, max_depth, lines, counter, controls):
    try:
        info = elem.element_info
        ct = info.control_type
        name = (info.name or "").replace("\n", " ")[:40]
        aid = info.automation_id or ""
        cls = info.class_name or ""
        try:
            r = info.rectangle
            rect = f"({r.left},{r.top},{r.right},{r.bottom})"
        except Exception:
            rect = ""
        counter[ct] += 1
        controls.append({"depth": depth, "type": ct, "name": name,
                         "auto_id": aid, "class": cls, "rect": rect})
        lines.append(f"{'  '*depth}[{ct}] name='{name}' id='{aid}' cls='{cls}' {rect}")
    except Exception as exc:
        lines.append(f"{'  '*depth}<err {exc}>")
        return
    if depth >= max_depth:
        return
    try:
        for child in elem.children():
            _walk(child, depth + 1, max_depth, lines, counter, controls)
    except Exception:
        pass


def probe(max_depth: int = 8) -> dict:
    cfg = get_config()
    cfg.report_dir.mkdir(parents=True, exist_ok=True)

    try:
        app = _connect_or_launch(cfg)
    except Exception as exc:
        result = {"ok": False, "error": f"无法启动/连接 Acuview 2: {exc}",
                  "recommendation": "改用坐标+视觉后端(需可交互桌面会话)"}
        (cfg.report_dir / "uia_summary.json").write_text(
            json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print("[probe] 失败:", exc)
        return result

    win = app.top_window()
    try:
        win.set_focus()
    except Exception:
        pass

    lines, counter, controls = [], Counter(), []
    _walk(win, 0, max_depth, lines, counter, controls)

    (cfg.report_dir / "uia_dump.txt").write_text("\n".join(lines), encoding="utf-8")

    interactive = {ct: counter.get(ct, 0) for ct in
                   ("Edit", "ComboBox", "Spinner", "Slider", "Button",
                    "CheckBox", "TreeItem", "Tree", "List", "ListItem", "TabItem")}
    total_interactive = sum(interactive.values())
    # 启发式结论
    uia_usable = (interactive.get("Edit", 0) + interactive.get("ComboBox", 0)
                  + interactive.get("Button", 0)) >= 5
    recommend = "uia" if uia_usable else "coord"

    summary = {
        "ok": True,
        "window_title": win.window_text(),
        "total_controls": len(controls),
        "by_type": dict(counter),
        "interactive_counts": interactive,
        "total_interactive": total_interactive,
        "uia_usable": uia_usable,
        "recommended_backend": recommend,
        "note": ("UIA 暴露了足够多的可操作控件，优先 UIA。"
                 if uia_usable else
                 "UIA 暴露的可操作控件很少(典型 Qt 自绘界面)，建议坐标+视觉兜底。"),
    }
    (cfg.report_dir / "uia_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"[probe] 窗口: {summary['window_title']}")
    print(f"[probe] 控件总数: {summary['total_controls']}")
    print(f"[probe] 类型分布: {summary['by_type']}")
    print(f"[probe] 可操作控件: {interactive}")
    print(f"[probe] 结论: 推荐后端 = {recommend} ({summary['note']})")
    print(f"[probe] 详细树 -> {cfg.report_dir / 'uia_dump.txt'}")
    return summary


if __name__ == "__main__":
    md = 8
    for a in sys.argv[1:]:
        if a.startswith("--depth="):
            md = int(a.split("=")[1])
    probe(max_depth=md)
