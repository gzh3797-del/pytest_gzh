"""
1E2W 接线检查自动化测试（HMI1-7）

运行：
    python projects/RPP/tests/Wiring_check/test_1e2w.py
"""
import sys, os, logging, time
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '../../../..'))

from playwright.sync_api import sync_playwright
from projects.RPP.tests.Wiring_check.core import config as cfg
from projects.RPP.tests.Wiring_check.core.meter_modbus import WiringCheckModbus
from projects.RPP.tests.Wiring_check.core import signal_driver as src
from projects.RPP.tests.Wiring_check.core.expected_engine import compute_1e2w
from projects.RPP.tests.Wiring_check.core.wiring_check_page import WiringCheckPage
from projects.RPP.tests.Wiring_check.core import report as rpt
from projects.RPP.tests.Wiring_check.core.tc_map import TC_MAP_1E2W as _TC_MAP
from projects.RPP.tests.Wiring_check.test_3e4wy import (
    _compare_voltage, _compare_current, DEVICE_NAME)

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')

V = cfg.NOMINAL_VOLTAGE
I = cfg.NORMAL_CURRENT


def _s(ua, qua=0, ia=I, qia=0):
    """1E2W 只有 A 相"""
    return dict(ua=ua, qua=qua, ub=0, qub=0, uc=0, quc=0,
                ia=ia, qia=qia, ib=0, qib=0, ic=0, qic=0)


_SPEC_REFS = {
    'PASS': '—',
    'V-01': '1LN(1E2W), 行2',
    'I-01': '1LN(1E2W), 行3',
    'I-02': '1LN(1E2W), 行4',
}

ALL_TESTS = [
    dict(id='PASS',  desc='正常接线 1E2W，Va=230V@0°',
         phase_order=cfg.PHASE_ABC, src=_s(V)),

    dict(id='V-01',  desc='Van<0.1*VRATE → Va Wiring Missing',
         phase_order=cfg.PHASE_ABC, src=_s(0)),

    dict(id='I-01',  desc='Ian<0.1A → Ia Wiring Missing',
         phase_order=cfg.PHASE_ABC, src=_s(V, 0, 0.05, 0)),

    dict(id='I-02',  desc='PF_A=-1（qia=180°）→ Ia Polarity Reversed',
         phase_order=cfg.PHASE_ABC, src=_s(V, 0, I, 180)),
]


def run_one(tc, modbus, wc_page):
    logging.info('[%s] %s', tc['id'], tc['desc'])
    t0 = time.time()
    s = tc['src']
    modbus.write_phase_order(tc['phase_order'])
    src.output(s['ua'], s['qua'], s['ub'], s['qub'], s['uc'], s['quc'],
               s['ia'], s['qia'], s['ib'], s['qib'], s['ic'], s['qic'])

    expected = compute_1e2w(s['ua'], s['qua'], s['ia'], s['qia'],
                            vrate=cfg.NOMINAL_VOLTAGE)
    logging.info('  预期: %s', expected)

    actual_v, actual_i = wc_page.run_check()
    v_ok, v_diffs = _compare_voltage(expected['voltage'], actual_v)
    i_ok, i_diffs = _compare_current(expected['current'], actual_i)
    passed = v_ok and i_ok
    detail = '\n'.join(v_diffs + i_diffs) or 'OK'
    logging.info('  → %s', 'PASS' if passed else 'FAIL')
    if not passed:
        logging.warning(detail)
    return dict(pass_=passed, expected=expected, actual_v=actual_v,
                actual_i=actual_i, detail=detail, phase_order=tc['phase_order'], src=s,
                spec_ref=_SPEC_REFS.get(tc['id'], ''),
                tc_id=_TC_MAP.get(tc['id'], ''), elapsed=round(time.time()-t0, 1))


def run_all(tests=ALL_TESTS, headless=False):
    cfg.ensure_meter_connection(headless=headless)
    modbus = WiringCheckModbus()
    try:
        modbus.write_service_config(cfg.SERVICE_1E2W)
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
        print(f'\n1E2W  总计:{total}  PASS:{passed}  FAIL:{total-passed}')
        for r in results:
            print(f"  [{'PASS' if r['pass_'] else 'FAIL'}] {r['id']}  {r['desc'][:55]}")
            if not r['pass_']:
                print(r['detail'])
        path = rpt.generate(results, '1E2W', DEVICE_NAME, cfg.METER_TCP_IP,
                            active_channels=cfg.ACTIVE_USER_CHANNELS[cfg.SERVICE_1E2W],
                            channel_phases=cfg.CHANNEL_PHASES[cfg.SERVICE_1E2W])
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

# 自定义 HTML 报告参数（pytest 路径下 conftest 的 wc_record 据此归类、生成报告）
REPORT_META = dict(
    wiring_type='1E2W',
    active_channels=cfg.ACTIVE_USER_CHANNELS[cfg.SERVICE_1E2W],
    channel_phases=cfg.CHANNEL_PHASES[cfg.SERVICE_1E2W],
)


@pytest.fixture(scope='module', autouse=True)
def _setup_1e2w(wiring_modbus):
    wiring_modbus.write_service_config(cfg.SERVICE_1E2W)


@pytest.mark.parametrize('tc', ALL_TESTS,
                         ids=lambda t: (f"{_TC_MAP[t['id']]}-" + t['id']) if t['id'] in _TC_MAP else t['id'])
def test_wiring_check(tc, wiring_modbus, wc_page, wc_record):
    result = run_one(tc, wiring_modbus, wc_page)
    wc_record({'id': tc['id'], 'desc': tc['desc'], **result})
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
