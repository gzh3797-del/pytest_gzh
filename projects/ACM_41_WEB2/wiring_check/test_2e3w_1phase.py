"""
2E3W 1Phase 接线检查自动化测试（电压侧6条 + 电流侧6条，A/C两相）

运行：
    python test_case/ACM_41_WEB2/Wiring_check/test_2e3w_1phase.py
"""
import sys, os, logging, time
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '../../..'))

from playwright.sync_api import sync_playwright
from projects.ACM_41_WEB2.wiring_check.core import config as cfg
from projects.ACM_41_WEB2.wiring_check.core.meter_modbus import WiringCheckModbus
from projects.ACM_41_WEB2.wiring_check.core import signal_driver as src
from projects.ACM_41_WEB2.wiring_check.core.expected_engine import compute_2e3w_1phase
from projects.ACM_41_WEB2.wiring_check.core.wiring_check_page import WiringCheckPage
from projects.ACM_41_WEB2.wiring_check.core import report as rpt
from projects.ACM_41_WEB2.wiring_check.core.tc_map import TC_MAP_1PHASE as _TC_MAP
from projects.ACM_41_WEB2.wiring_check.test_3e4wy import (
    _compare_voltage, _compare_current, DEVICE_NAME)

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')

V = cfg.NOMINAL_VOLTAGE   # 230V
I = cfg.NORMAL_CURRENT

# 1Phase 只有 A/C 两相，B 相源输出置 0（不参与检测）
def _s(ua, qua, uc, quc, ia=I, qia=0, ic=I, qic=180):
    """A/C 两相，默认 Ia@0°, Ic@180°（PF=1）。
    ib 角度跟 ic 保持一致，防止台架 B 路电流接 Phase C CT 时产生角度误差。"""
    return dict(ua=ua, qua=qua, ub=0, qub=0, uc=uc, quc=quc,
                ia=ia, qia=qia, ib=I, qib=qic, ic=ic, qic=qic)


ALL_TESTS = [
    dict(id='PASS',   desc='正常接线 1Phase，Va@0°, Vc@180°',
         phase_order=cfg.PHASE_ABC, src=_s(V, 0, V, 180)),

    # 电压缺失
    dict(id='V-01', desc='条件1：Vca<0.1*VRATE → Va&Vc Wiring Missing',
         phase_order=cfg.PHASE_ABC, src=_s(0, 0, 0, 0)),
    dict(id='V-02', desc='条件2：Van<0.8*VRATE → Va Wiring Missing',
         phase_order=cfg.PHASE_ABC, src=_s(int(V*0.75), 0, V, 180)),
    dict(id='V-03', desc='条件3：Vcn<0.8*VRATE → Vc Wiring Missing',
         phase_order=cfg.PHASE_ABC, src=_s(V, 0, int(V*0.75), 180)),

    # 电压反接
    dict(id='V-04', desc='条件4：Van=230@0°, Vcn=400@? → Va-Vn Reversed',
         phase_order=cfg.PHASE_ABC, src=_s(V, 0, 400, 180)),
    dict(id='V-05', desc='条件5：Vcn=230@180°, Van=400@0° → Vc-Vn Reversed',
         phase_order=cfg.PHASE_ABC, src=_s(400, 0, V, 180)),

    # 电流缺失
    dict(id='I-01', desc='条件7：Ian<0.1A → Ia Wiring Missing',
         phase_order=cfg.PHASE_ABC,
         src=_s(V, 0, V, 180, ia=0.05, qia=0, ic=I, qic=180)),
    dict(id='I-02', desc='条件8：Icn<0.1A → Ic Wiring Missing',
         phase_order=cfg.PHASE_ABC,
         src=_s(V, 0, V, 180, ia=I, qia=0, ic=0.05, qic=180)),

    # 电流反接
    dict(id='I-03', desc='条件9：PF_A=-1（qia=180°）→ Ia Polarity Reversed',
         phase_order=cfg.PHASE_ABC,
         src=_s(V, 0, V, 180, ia=I, qia=180, ic=I, qic=180)),
    dict(id='I-04', desc='条件10：PF_C=-1（qic=0°）→ Ic Polarity Reversed',
         phase_order=cfg.PHASE_ABC,
         src=_s(V, 0, V, 180, ia=I, qia=0, ic=I, qic=0)),

    # 电流相位错误
    dict(id='I-05', desc='条件11：PF_A=0（qia=90°）→ Ia Phase Shift',
         phase_order=cfg.PHASE_ABC,
         src=_s(V, 0, V, 180, ia=I, qia=90, ic=I, qic=180)),
    dict(id='I-06', desc='条件12：PF_C=0（qic=90°）→ Ic Phase Shift',
         phase_order=cfg.PHASE_ABC,
         src=_s(V, 0, V, 180, ia=I, qia=0, ic=I, qic=90)),
]

VOLTAGE_TESTS = [{**tc, 'check_current': False}
                 for tc in ALL_TESTS if tc['id'].startswith(('PASS', 'V-'))]

# 用例 → 接线检测总表_ver1.05.xlsx 行号对照（Sheet: 2E3W 1Phase）
_SPEC_REFS = {
    'PASS':  '—',
    'V-01':  '2E3W 1Phase, 行2',
    'V-02':  '2E3W 1Phase, 行3',
    'V-03':  '2E3W 1Phase, 行4',
    'V-04':  '2E3W 1Phase, 行5',
    'V-05':  '2E3W 1Phase, 行6',
    'I-01':  '2E3W 1Phase, 行11',
    'I-02':  '2E3W 1Phase, 行12',
    'I-03':  '2E3W 1Phase, 行13',
    'I-04':  '2E3W 1Phase, 行14',
    'I-05':  '2E3W 1Phase, 行15',
    'I-06':  '2E3W 1Phase, 行16',
}


def run_one(tc, modbus, wc_page):
    logging.info('[%s] %s', tc['id'], tc['desc'])
    t0 = time.time()
    s = tc['src']
    modbus.write_phase_order(tc['phase_order'])
    src.output(s['ua'], s['qua'], s['ub'], s['qub'], s['uc'], s['quc'],
               s['ia'], s['qia'], s['ib'], s['qib'], s['ic'], s['qic'])

    expected = compute_2e3w_1phase(
        s['ua'], s['qua'], s['uc'], s['quc'],
        s['ia'], s['qia'], s['ic'], s['qic'],
        vrate=cfg.NOMINAL_VOLTAGE)
    logging.info('  预期: %s', expected)

    actual_v, actual_i = wc_page.run_check()
    v_ok, v_diffs = _compare_voltage(expected['voltage'], actual_v)
    if tc.get('check_current', True):
        i_ok, i_diffs = _compare_current(expected['current'], actual_i)
    else:
        i_ok, i_diffs = True, []
    passed = v_ok and i_ok
    detail = '\n'.join(v_diffs + i_diffs) or 'OK'
    logging.info('  → %s', 'PASS' if passed else 'FAIL')
    if not passed: logging.warning(detail)
    return dict(pass_=passed, expected=expected, actual_v=actual_v,
                actual_i=actual_i, detail=detail, phase_order=tc['phase_order'], src=s,
                spec_ref=_SPEC_REFS.get(tc['id'], ''),
                tc_id=_TC_MAP.get(tc['id'], ''), elapsed=round(time.time()-t0, 1))


def run_all(tests=ALL_TESTS, headless=False):
    run_start = time.time()
    modbus = WiringCheckModbus()
    try:
        modbus.write_service_config(cfg.SERVICE_2E3W_1PHASE)
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
        print(f'\n2E3W 1Phase  总计:{total}  PASS:{passed}  FAIL:{total-passed}')
        for r in results:
            print(f"  [{'PASS' if r['pass_'] else 'FAIL'}] {r['id']}  {r['desc'][:55]}")
            if not r['pass_']: print(r['detail'])
        path = rpt.generate(results, '2E3W 1Phase', DEVICE_NAME, cfg.METER_TCP_IP,
                            active_channels=cfg.ACTIVE_USER_CHANNELS[cfg.SERVICE_2E3W_1PHASE],
                            channel_phases=cfg.CHANNEL_PHASES[cfg.SERVICE_2E3W_1PHASE],
                            voltage_phases=('A', 'C'))
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
def _setup_2e3w_1phase(wiring_modbus):
    wiring_modbus.write_service_config(cfg.SERVICE_2E3W_1PHASE)


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
