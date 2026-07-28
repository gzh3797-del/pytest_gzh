# -*- coding: utf-8 -*-
"""离线验证精度引擎：用假读表器/假源(不需真硬件)跑通整批用例、单点、停止、存回xlsx。"""
import os
import threading

import accuracy_engine as eng
from fast_accuracy_test.core.source_iface import Source
from fast_accuracy_test.core.excel_input import Case

STATE = {"v": 0.0, "i": 0.0}


class FakeReader:
    """读表器替身：读数=源设定值(略加偏差)，让精度判定可通过。"""
    def __init__(self, cfg, client):
        self.cfg = cfg

    def read(self, name):
        v, i = STATE["v"], STATE["i"]
        if name in ("voltage",):
            return v * 1.0001            # 0.01% 偏差
        if name in ("current", "current_1"):
            return i * 1.0001
        if name in ("power", "power_1"):
            return v * i / 1000 * 1.0001
        return 0.0

    def close(self):
        pass


class FakeDev:
    def __init__(self):
        self.config_calls = 0

    def config(self, params):
        self.config_calls += 1

    def source_stop(self):
        pass


def fake_make_source(dev, rated_v, rated_i, base_params=None, pulse_read_count=5):
    def out(voltage, current):
        STATE["v"], STATE["i"] = voltage, current
    return Source(out, lambda: None)


def _patch():
    eng.make_client = lambda cfg: object()
    eng.ModbusReader = FakeReader
    eng.make_source = fake_make_source


def test_run_cases():
    _patch()
    dev = FakeDev()
    cases = [
        Case("c1", 60.0, 1.2, 0.0, 0.0, 0.001, 0.075, 0.075, 3, 0.0),
        Case("c2", 100.0, 5.0, 0.0, 0.0, 0.001, 0.075, 0.075, 3, 0.0),
    ]
    events = []
    path = eng.run_cases(dev, 1000.0, 500.0, {"device_model": "320"}, cases,
                         source_params="dummy",
                         progress=lambda k, p: events.append((k, p)))
    assert dev.config_calls == 1, "应发一次参数配置"
    assert os.path.isfile(path), f"报告未生成: {path}"
    overalls = [p.overall for k, p in events if k == "case_done"]
    assert overalls == ["Passed", "Passed"], overalls
    print("[PASS] run_cases 出报告:", os.path.basename(path), "判定:", overalls)
    os.remove(path)


def test_run_point():
    _patch()
    res = eng.run_point(FakeDev(), 1000.0, 500.0, {"device_model": "320"}, 60.0, 1.2)
    assert res.overall == "Passed", res.overall
    print("[PASS] run_point 单点:", res.overall,
          {m.label: m.result for m in res.metrics})


def test_stop():
    _patch()
    ev = threading.Event()
    ev.set()  # 一开始就请求停止
    cases = [Case("c1", 60.0, 1.2, 0.0, 0.0, 0.001, 0.075, 0.075, 3, 0.0)]
    kinds = []
    path = eng.run_cases(FakeDev(), 1000.0, 500.0, {"device_model": "320"}, cases,
                         progress=lambda k, p: kinds.append(k), stop_event=ev)
    assert "stopped" in kinds, kinds
    assert path == "", "停止且无结果应返回空报告路径"
    print("[PASS] stop 立即停止, kinds=", kinds)


def test_save_xlsx_roundtrip():
    cases = eng.load_cases()
    n = len(cases)
    orig = cases[0].voltage
    cases[0].voltage = 123.45
    eng.save_cases_xlsx(cases)
    reloaded = eng.load_cases()
    assert len(reloaded) == n, (len(reloaded), n)
    assert abs(reloaded[0].voltage - 123.45) < 1e-6, reloaded[0].voltage
    # 还原
    cases[0].voltage = orig
    eng.save_cases_xlsx(cases)
    print("[PASS] save_cases_xlsx 往返: 写入123.45读回", reloaded[0].voltage, "已还原")


if __name__ == "__main__":
    test_run_cases()
    test_run_point()
    test_stop()
    test_save_xlsx_roundtrip()
    print("全部通过 ✓")
