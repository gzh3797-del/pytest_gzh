"""
2E3W Network 接线检查自动化测试（HMI1-7，算法与 3E4WY 完全相同）

运行：
    python test_case/AcuHMI_1_7/wiring_check/test_2e3w_network.py
"""
import sys, os, logging, time
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '../../..'))

from playwright.sync_api import sync_playwright
from test_case.AcuHMI_1_7.wiring_check.core import config as cfg
from test_case.AcuHMI_1_7.wiring_check.core.meter_modbus import WiringCheckModbus
from test_case.AcuHMI_1_7.wiring_check.core import signal_driver as src
from test_case.AcuHMI_1_7.wiring_check.core.expected_engine import compute_2e3w_network
from test_case.AcuHMI_1_7.wiring_check.core.wiring_check_page import WiringCheckPage
from test_case.AcuHMI_1_7.wiring_check.core import report as rpt
from test_case.AcuHMI_1_7.wiring_check.core.tc_map import TC_MAP_NET as _TC_MAP
from test_case.AcuHMI_1_7.wiring_check.test_3e4wy import (
    _compare_voltage, _compare_current, DEVICE_NAME,
    ABC_TESTS, ACB_TESTS, ALL_TESTS, VOLTAGE_TESTS, _SPEC_REFS as _SPEC_REFS_NET)

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')

# 2E3W Network 测试用例与 3E4WY 完全一致，仅接线方式寄存器不同


def run_one(tc, modbus, wc_page):
    logging.info('[%s] %s', tc['id'], tc['desc'])
    t0 = time.time()
    s = tc['src']
    modbus.write_phase_order(tc['phase_order'])
    src.output(s['ua'], s['qua'], s['ub'], s['qub'], s['uc'], s['quc'],
               s['ia'], s['qia'], s['ib'], s['qib'], s['ic'], s['qic'])

    expected = compute_2e3w_network(
        s['ua'], s['qua'], s['ub'], s['qub'], s['uc'], s['quc'],
        s['ia'], s['qia'], s['ib'], s['qib'], s['ic'], s['qic'],
        vrate=cfg.NOMINAL_VOLTAGE, phase_order=tc['phase_order'])
    logging.info('  预期: %s', expected)

    actual_v, actual_i = wc_page.run_check()
    v_ok, v_diffs = _compare_voltage(expected['voltage'], actual_v)
    if tc.get('check_current', True):
        i_ok, i_diffs = _compare_current(expected['current'], actual_i,
                                         channel_phase_map=cfg.NETWORK_CHANNEL_PHASE_MAP)
    else:
        i_ok, i_diffs = True, []
    passed = v_ok and i_ok
    detail = '\n'.join(v_diffs + i_diffs) or 'OK'
    logging.info('  → %s', 'PASS' if passed else 'FAIL')
    if not passed:
        logging.warning(detail)
    return dict(pass_=passed, expected=expected, actual_v=actual_v,
                actual_i=actual_i, detail=detail, phase_order=tc['phase_order'], src=s,
                spec_ref=_SPEC_REFS_NET.get(tc['id'], ''),
                tc_id=_TC_MAP.get(tc['id'], ''), elapsed=round(time.time()-t0, 1))


def run_all(tests=ALL_TESTS, headless=False):
    modbus = WiringCheckModbus()
    try:
        modbus.write_service_config(cfg.SERVICE_2E3W_NET)
        modbus.write_nominal_voltage(cfg.NOMINAL_VOLTAGE)
        results = []
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=headless)
            page = browser.new_page(ignore_https_errors=True)
            wc_page = WiringCheckPage(page, device_name=DEVICE_NAME)
            wc_page.login_if_needed()
            wc_page.navigate()
            for tc in tests:
                res = run_one(tc, modbus, wc_page)
                results.append({'id': tc['id'], 'desc': tc['desc'], **res})
            browser.close()

        total  = len(results)
        passed = sum(1 for r in results if r['pass_'])
        print(f'\n2E3W Network  总计:{total}  PASS:{passed}  FAIL:{total-passed}')
        for r in results:
            print(f"  [{'PASS' if r['pass_'] else 'FAIL'}] {r['id']}  {r['desc'][:55]}")
            if not r['pass_']:
                print(r['detail'])
        path = rpt.generate(results, '2E3W Network', DEVICE_NAME, cfg.METER_TCP_IP,
                            active_channels=cfg.ACTIVE_USER_CHANNELS[cfg.SERVICE_2E3W_NET],
                            channel_phases=cfg.CHANNEL_PHASES[cfg.SERVICE_2E3W_NET],
                            channel_phase_map=cfg.NETWORK_CHANNEL_PHASE_MAP)
        print(f'报告：{path}')
        return results
    except KeyboardInterrupt:
        print('[CTRL+C] 检测到中断，正在关源...')
        raise
    finally:
        src.stop()
        modbus.close()


# ── pytest 入口 ───────────────────────────────────────────────────────────────
import pytest


@pytest.fixture(scope='module', autouse=True)
def _setup_2e3w_network(wiring_modbus):
    wiring_modbus.write_service_config(cfg.SERVICE_2E3W_NET)


@pytest.mark.parametrize('tc', ALL_TESTS,
                         ids=lambda t: (f"{_TC_MAP[t['id']]}-" + t['id']) if t['id'] in _TC_MAP else t['id'])
def test_wiring_check(tc, wiring_modbus, wc_page):
    result = run_one(tc, wiring_modbus, wc_page)
    assert result['pass_'], result['detail']


if __name__ == '__main__':
    _ids = sys.argv[1:]
    if _ids:
        _tests = [tc for tc in ALL_TESTS if tc['id'] in _ids]
        if not _tests:
            print(f'[WARN] 未找到匹配 ID，可用：\n  {[t["id"] for t in ALL_TESTS]}')
            sys.exit(1)
    else:
        _tests = ALL_TESTS
    run_all(tests=_tests)
