"""跨传输闭环校验 + 报告。

用 *校验通道*(默认与 GUI 不同的传输，规避 COM4 串口独占)从设备读真值，
与"期望值 / GUI 显示值"比对，输出 pass/fail。报告写到 reports/。
"""
from __future__ import annotations

import csv
import json
import time
from dataclasses import dataclass, field, asdict

from .config import get_config
from .modbus_client import MeterClient


@dataclass
class Check:
    label: str
    expected: object
    actual: object
    passed: bool
    detail: str = ""


@dataclass
class Report:
    title: str
    started: str
    checks: list = field(default_factory=list)

    def add(self, c: Check):
        self.checks.append(c)
        flag = "PASS" if c.passed else "FAIL"
        print(f"  [{flag}] {c.label}: expected={c.expected!r} actual={c.actual!r} {c.detail}")

    @property
    def passed(self) -> bool:
        return all(c.passed for c in self.checks) and bool(self.checks)

    def save(self, name: str):
        cfg = get_config()
        cfg.report_dir.mkdir(parents=True, exist_ok=True)
        base = cfg.report_dir / name
        data = {"title": self.title, "started": self.started,
                "passed": self.passed,
                "checks": [asdict(c) for c in self.checks]}
        base.with_suffix(".json").write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        with open(base.with_suffix(".csv"), "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["label", "expected", "actual", "passed", "detail"])
            for c in self.checks:
                w.writerow([c.label, c.expected, c.actual, c.passed, c.detail])
        print(f"[report] {self.title}: {'全部通过' if self.passed else '存在失败'} "
              f"({sum(c.passed for c in self.checks)}/{len(self.checks)})  -> {base}.json")
        return base


def values_match(expected, actual, tol: float) -> tuple[bool, str]:
    """浮点按相对容差比较；整数/字符串按相等。"""
    try:
        ef, af = float(expected), float(actual)
        if abs(ef) < 1e-9:
            ok = abs(af - ef) <= max(tol, 1e-6)
        else:
            ok = abs(af - ef) / abs(ef) <= tol
        return ok, f"(rel_err={abs(af-ef)/abs(ef) if ef else abs(af-ef):.4g})"
    except (TypeError, ValueError):
        return str(expected).strip() == str(actual).strip(), ""


def now() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


class Verifier:
    """封装"用校验通道读寄存器并断言"。"""

    def __init__(self, transport: str | None = None):
        self.cfg = get_config()
        self.client = MeterClient(transport=transport or self.cfg.transport.verify)
        self.tol = float(self.cfg.run.value_tolerance)

    def __enter__(self):
        self.client.connect()
        self.client.calibrate_word_order()
        return self

    def __exit__(self, *exc):
        self.client.close()

    def read_truth(self, name_or_addr):
        return self.client.read(name_or_addr)

    def check(self, report: Report, label: str, expected, name_or_addr):
        actual = self.read_truth(name_or_addr)
        ok, detail = values_match(expected, actual, self.tol)
        report.add(Check(label=label, expected=expected, actual=actual, passed=ok, detail=detail))
        return ok
