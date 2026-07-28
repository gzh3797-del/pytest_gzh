# fast_accuracy_test/main.py
#
# 用法：
#   python -m fast_accuracy_test.main                  跑 xlsx 所有用例(自动控源)
#   python -m fast_accuracy_test.main --manual         跑 xlsx 所有用例，但你手动设源(脚本提示你设好按回车)
#   python -m fast_accuracy_test.main --point 60 1.2   只测一个点(自动控源)，打印读数与误差
#   python -m fast_accuracy_test.main --point 60 1.2 --manual   只测一个点，且你手动设源
#
# 换控源：改 config.json 的 "source_module"（默认 source_control=科陆），或加 --manual 用手动源。
import logging
import os
import sys
import time

from fast_accuracy_test.core.config import load_config
from fast_accuracy_test.core.excel_input import read_cases, Case
from fast_accuracy_test.core.modbus_reader import make_client, ModbusReader
from fast_accuracy_test.core.source_iface import load_source
from fast_accuracy_test.core.case_runner import run_case
from fast_accuracy_test.core.report import write_report

# 项目根目录(config.json / test_data / result 都在这里)，按代码位置定位，
# 这样无论从哪个目录运行都能找到，不依赖“当前工作目录”。
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _default_config():
    return os.path.join(_ROOT, "config.json")


def _resolve(path, base):
    """相对路径按 base(config.json 所在目录)解析为绝对路径。"""
    return path if os.path.isabs(path) else os.path.normpath(os.path.join(base, path))


def _build_source(cfg, manual):
    """manual=True 用手动源(source_manual)，否则用 config 里的 source_module。"""
    name = "source_manual" if manual else cfg.get("source_module", "source_control")
    return load_source(name)


def _teardown(source, reader):
    try:
        source.output(0, 0)
        source.stop()
    except Exception:
        logging.exception("控源收尾失败")
    try:
        reader.close()
    except Exception:
        logging.exception("读表关闭失败")


def run(config_path=None, *, manual=False, source=None, reader=None):
    """跑 xlsx 全部用例，出 Excel 报告，返回报告路径(无用例返回 "")。"""
    config_path = config_path or _default_config()
    cfg = load_config(config_path)
    base = os.path.dirname(os.path.abspath(config_path))
    input_xlsx = _resolve(cfg["input_xlsx"], base)
    result_dir = _resolve(cfg["result_dir"], base)
    if reader is None:
        reader = ModbusReader(cfg, make_client(cfg))
    if source is None:
        source = _build_source(cfg, manual)

    cases = read_cases(input_xlsx)
    # 本轮只要有任一用例填了脉冲两列，所有行都带“脉冲误差”列(不填的行为 N/A)，保证报告列对齐
    include_pulse = any(c.pulse_const is not None and c.pulse_error_acc is not None for c in cases)
    results = []
    try:
        for c in cases:
            logging.info("测试进度: %s @ %s", c.test_case, time.strftime("%Y-%m-%d %H:%M:%S"))
            print(f"测试进度: {c.test_case}")
            try:
                results.append(run_case(c, reader, source, is_dual=cfg["is_dual"],
                                        settle_s=cfg.get("settle_s", 5.0), include_pulse=include_pulse))
            except Exception:
                logging.exception("用例执行失败: %s", c.test_case)
        return write_report(results, result_dir, cfg["device_model"]) if results else ""
    finally:
        _teardown(source, reader)


def run_point(voltage, current, *, vacc=0.001, iacc=0.075, pacc=0.075,
              config_path=None, manual=False, source=None, reader=None):
    """只测单个电压/电流点，打印各项读数与误差，返回 CaseResult。"""
    cfg = load_config(config_path or _default_config())
    if reader is None:
        reader = ModbusReader(cfg, make_client(cfg))
    if source is None:
        source = _build_source(cfg, manual)

    case = Case(test_case="manual_point", voltage=voltage, current_1=current, current_2=0.0,
                wait_h=0.0, voltage_accuracy=vacc, current_accuracy=iacc, power_accuracy=pacc,
                sample_cnt=20, sample_interval=0.1)
    try:
        res = run_case(case, reader, source, is_dual=cfg["is_dual"], settle_s=cfg.get("settle_s", 5.0))
        print_result(res)
        return res
    finally:
        _teardown(source, reader)


def print_result(res):
    """把单个用例结果按项打印：真值 / min / max / avg / 平均误差% / 判定。"""
    print("=" * 72)
    print("用例: %s   总判定: %s" % (res.case.test_case, res.overall))
    print("-" * 72)
    print("%-10s %10s %10s %10s %10s %9s %9s %s" %
          ("项目", "真值", "min", "max", "avg", "平均误差%", "最大误差%", "判定"))
    for m in res.metrics:
        if m.result == "N/A":
            print("%-10s %10s   (未测/不适用)" % (m.label, _fmt(m.ref)))
        else:
            print("%-10s %10s %10s %10s %10s %9s %9s %s（单位:%s）" %
                  (m.label, _fmt(m.ref), _fmt(m.min), _fmt(m.max), _fmt(m.avg),
                   _fmt(m.err_avg), _fmt(m.err_max), m.result, m.unit))
    print("=" * 72)


def _fmt(x):
    return "-" if x is None else ("%.5g" % x)


def _main(argv):
    logging.basicConfig(level=logging.INFO)
    manual = "--manual" in argv
    rest = [a for a in argv if a != "--manual"]

    if rest[:1] == ["--point"]:
        if len(rest) < 3:
            print("用法: --point <电压> <电流> [电压精度 电流精度 功率精度] [--manual]")
            return
        v, i = float(rest[1]), float(rest[2])
        acc = [float(x) for x in rest[3:6]]
        kw = {}
        if len(acc) == 3:
            kw = dict(vacc=acc[0], iacc=acc[1], pacc=acc[2])
        run_point(v, i, manual=manual, **kw)
        return

    print("==================== Precision Measure Start ====================")
    t0 = time.time()
    out = run(manual=manual)
    print(f"报告: {out}")
    print(f"==================== 总耗时: {time.time() - t0:.1f}s ====================")


if __name__ == "__main__":
    _main(sys.argv[1:])
