"""updata 用例集 pytest 配置：截图 + Excel/JSON 报告 + Windows 防睡眠。

报告体系与 projects/AcuRev1320/QT_Auto/conftest.py 保持一致：每条用例在
reports/<时间戳>/ 下产出 HTML / Excel / JSON 报告及分层截图。
"""
import os
import sys
import json
import ctypes
import traceback
from datetime import datetime
from pathlib import Path

import pytest
import pandas as pd
import pyautogui

# Windows 控制台默认 GBK，emoji/部分中文 print 会 UnicodeEncodeError；统一切到 UTF-8。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding='utf-8')
    except (AttributeError, ValueError):
        pass

# 路径引导：updata 自身 + QT_Auto（firmware_layout）+ 仓库根（comm / modbus_config）。
_HERE = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(_HERE))))
_QT_AUTO = os.path.join(_REPO_ROOT, 'projects', 'AcuRev1320', 'QT_Auto')
for _p in (_REPO_ROOT, _QT_AUTO, _HERE):
    if _p not in sys.path:
        sys.path.insert(0, _p)


class _ResultCollector:
    """收集测试结果并生成 Excel / JSON 报告。"""

    def __init__(self):
        self.results = []
        self.stats = {"total": 0, "passed": 0, "failed": 0, "skipped": 0, "error": 0, "duration": 0.0}

    def add(self, report, screenshot_path=None):
        if report.when != "call" and not (report.when == "setup" and report.skipped):
            return
        node = report.nodeid
        file_part = node.split("::")[0] if "::" in node else node
        name = "::".join(node.split("::")[1:]) if "::" in node else node
        file_name = os.path.splitext(os.path.basename(file_part))[0]
        if report.passed:
            status = "通过"; self.stats["passed"] += 1
        elif report.failed:
            status = "失败"; self.stats["failed"] += 1
        elif report.skipped:
            status = "跳过"; self.stats["skipped"] += 1
        else:
            status = "错误"; self.stats["error"] += 1
        row = {
            "测试文件": file_name,
            "用例名称": name[5:] if name.startswith("test_") else name,
            "是否通过": status,
            "执行时间(秒)": round(report.duration, 3),
            "错误信息": str(report.longrepr) if (report.failed or report.skipped) else "",
        }
        if screenshot_path:
            row["截图路径"] = str(Path(screenshot_path).resolve())
        self.results.append(row)
        self.stats["total"] += 1
        self.stats["duration"] += report.duration

    def write_reports(self, report_dir):
        if not self.results:
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        df = pd.DataFrame(self.results)
        excel = report_dir / f"test_report_{ts}.xlsx"
        with pd.ExcelWriter(excel, engine='openpyxl') as w:
            df.to_excel(w, sheet_name='测试详情', index=False)
            total = self.stats["total"] or 1
            summary = pd.DataFrame({
                "统计项目": ["总用例数", "通过数", "失败数", "跳过数", "错误数", "执行时间(秒)", "通过率"],
                "数值": [self.stats["total"], self.stats["passed"], self.stats["failed"],
                         self.stats["skipped"], self.stats["error"], round(self.stats["duration"], 3),
                         f"{self.stats['passed'] / total * 100:.2f}%"],
            })
            summary.to_excel(w, sheet_name='测试统计', index=False)
        with open(report_dir / f"test_report_{ts}.json", 'w', encoding='utf-8') as f:
            json.dump({"timestamp": datetime.now().isoformat(), "statistics": self.stats,
                       "test_results": self.results}, f, ensure_ascii=False, indent=2)
        print(f"\n📊 Excel/JSON 报告已生成: {excel}")

    def print_summary(self, report_dir):
        s = self.stats
        total = s["total"] or 1
        print("\n" + "=" * 50)
        print(f"📁 报告目录: {report_dir}")
        print(f"📈 总计:{s['total']}  ✅通过:{s['passed']}  ❌失败:{s['failed']}  "
              f"⏭️跳过:{s['skipped']}  ⚠️错误:{s['error']}")
        print(f"📊 通过率: {s['passed'] / total * 100:.2f}%  ⏱️耗时: {s['duration']:.1f}s")
        print("=" * 50)


class _ScreenshotManager:
    def __init__(self, base_dir):
        self.root = base_dir / "screenshots"
        self.root.mkdir(parents=True, exist_ok=True)
        self.counter = {}

    @staticmethod
    def _safe(name):
        for ch in '<>:"/\\|?*':
            name = name.replace(ch, '_')
        return name.strip().replace(' ', '_')[:100]

    def take(self, node_id, is_failed):
        try:
            file_part = node_id.split("::")[0]
            file_name = self._safe(os.path.splitext(os.path.basename(file_part))[0])
            func = self._safe("_".join(node_id.split("::")[1:]) or "unnamed")
            d = self.root / file_name / func
            d.mkdir(parents=True, exist_ok=True)
            self.counter[node_id] = self.counter.get(node_id, 0) + 1
            status = "FAILED" if is_failed else "PASSED"
            ts = datetime.now().strftime("%H%M%S_%f")[:-3]
            path = d / f"{status}_{ts}_{self.counter[node_id]:03d}.png"
            pyautogui.screenshot().save(str(path))
            return path
        except Exception as e:
            print(f"❌ 截图失败: {e}")
            return None


_collector = _ResultCollector()
_screenshot_mgr = None


@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_makereport(item, call):
    global _screenshot_mgr
    outcome = yield
    report = outcome.get_result()
    if report.when == 'call':
        shot = None
        if _screenshot_mgr is not None and getattr(item, 'instance', None) is not None:
            try:
                shot = _screenshot_mgr.take(item.nodeid, report.failed)
            except Exception:
                traceback.print_exc()
        _collector.add(report, shot)
    elif report.when == 'setup' and report.skipped:
        _collector.add(report)


@pytest.hookimpl(tryfirst=True)
def pytest_configure(config):
    global _screenshot_mgr
    if not config.pluginmanager.hasplugin('html'):
        try:
            config.pluginmanager.import_plugin('pytest_html.plugin')
        except Exception:
            pass
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    config.report_dir = Path("reports") / f"updata_{ts}"
    config.report_dir.mkdir(parents=True, exist_ok=True)
    _screenshot_mgr = _ScreenshotManager(config.report_dir)
    if hasattr(config.option, 'htmlpath'):
        config.option.htmlpath = str(config.report_dir / f"test_report_{ts}.html")
        config.option.self_contained_html = True
    # Windows 防睡眠
    if sys.platform == "win32":
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000 | 0x00000001 | 0x00000002)
        print("🔋 防睡眠已启用")


@pytest.hookimpl(trylast=True)
def pytest_sessionfinish(session, exitstatus):
    try:
        _collector.write_reports(session.config.report_dir)
        _collector.print_summary(session.config.report_dir)
    except Exception:
        traceback.print_exc()
    finally:
        if sys.platform == "win32":
            ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)
