"""
2E3W Delta 接线检查自动化测试（HMI1-7）

运行：
    python projects/AcuHMI_1_7/tests/Wiring_check/test_2e3w_delta.py
"""
import sys, os, logging, time
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '../../../..'))

from playwright.sync_api import sync_playwright
from projects.AcuHMI_1_7.tests.Wiring_check.core import config as cfg
from projects.AcuHMI_1_7.tests.Wiring_check.core.meter_modbus import WiringCheckModbus
from projects.AcuHMI_1_7.tests.Wiring_check.core import signal_driver as src
from projects.AcuHMI_1_7.tests.Wiring_check.core.expected_engine import compute_2e3w_delta
from projects.AcuHMI_1_7.tests.Wiring_check.core.wiring_check_page import WiringCheckPage
from projects.AcuHMI_1_7.tests.Wiring_check.core import report as rpt
from projects.AcuHMI_1_7.tests.Wiring_check.core.tc_map import TC_MAP_DELTA as _TC_MAP
from projects.AcuHMI_1_7.tests.Wiring_check.test_3e4wy import (
    _compare_voltage, _compare_current, _status_match, DEVICE_NAME)

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')

V  = cfg.NOMINAL_VOLTAGE
VL = round(V / 1.732, 1)
I  = cfg.NORMAL_CURRENT


def _s(ua, qua, ub, qub, uc, quc, ia=I, qia=0, ic=I, qic=120):
    """Delta 默认电流绝对角（规格角+Vab偏移）：
    ABC正常：规格Ia@330°+Vab30°=绝对0°，规格Ic@90°+Vab30°=绝对120°"""
    return dict(ua=ua, qua=qua, ub=ub, qub=qub, uc=uc, quc=quc,
                ia=ia, qia=qia, ic=ic, qic=qic)


ABC_TESTS = [
    dict(id='PASS-ABC', desc='正常接线 Delta ABC',
         phase_order=cfg.PHASE_ABC,
         src=_s(VL, 0, VL, 240, VL, 120)),

    # 电压缺失
    dict(id='V-01', desc='条件1：三相线电压全<0.1*VRATE → Va&Vb&Vc Missing',
         phase_order=cfg.PHASE_ABC, src=_s(0, 0, 0, 0, 0, 0)),
    dict(id='V-02', desc='条件2：Vab<0.1*VRATE → Va&Vb Missing',
         phase_order=cfg.PHASE_ABC, src=_s(VL, 0, VL, 0, VL, 120)),
    dict(id='V-03', desc='条件3：Vbc<0.1*VRATE → Vc&Vb Missing',
         phase_order=cfg.PHASE_ABC, src=_s(VL, 0, VL, 240, VL, 240)),
    dict(id='V-04', desc='条件4：Vca<0.1*VRATE → Vc&Va Missing',
         phase_order=cfg.PHASE_ABC, src=_s(VL, 0, VL, 240, VL, 0)),
    dict(id='V-05', desc='条件5：Vbc正常，Vab/Vca偏低 → Va Missing',
         phase_order=cfg.PHASE_ABC,
         src=_s(int(VL*0.5), 0, VL, 240, int(VL*0.5), 120)),
    dict(id='V-06', desc='条件6：Vca正常，Vab/Vbc偏低 → Vb Missing',
         phase_order=cfg.PHASE_ABC,
         src=_s(int(VL*0.5), 0, int(VL*0.5), 240, VL, 120)),
    dict(id='V-07', desc='条件7：Vab正常，Vbc/Vca偏低 → Vc Missing',
         phase_order=cfg.PHASE_ABC,
         src=_s(VL, 0, int(VL*0.5), 240, int(VL*0.5), 120)),

    # 相序错误（Delta 无反接检测）
    dict(id='V-08', desc='条件8：ACB信号接ABC配置 → Phase Order Error',
         phase_order=cfg.PHASE_ABC, src=_s(VL, 0, VL, 120, VL, 240)),

    # 电流缺失
    dict(id='I-01', desc='条件9：Ian<0.1A → Ia Wiring Missing',
         phase_order=cfg.PHASE_ABC,
         src=_s(VL, 0, VL, 240, VL, 120, ia=0.05, qia=0, ic=I, qic=120)),
    dict(id='I-02', desc='条件10：Icn<0.1A → Ic Wiring Missing',
         phase_order=cfg.PHASE_ABC,
         src=_s(VL, 0, VL, 240, VL, 120, ia=I, qia=0, ic=0.05, qic=120)),

    # 电流反接
    dict(id='I-03', desc='条件11：∠IA≈150° → Ia Polarity Reversed',
         phase_order=cfg.PHASE_ABC,
         src=_s(VL, 0, VL, 240, VL, 120, ia=I, qia=180, ic=I, qic=120)),
    dict(id='I-04', desc='条件12：∠IC≈270° → Ic Polarity Reversed',
         phase_order=cfg.PHASE_ABC,
         src=_s(VL, 0, VL, 240, VL, 120, ia=I, qia=0, ic=I, qic=300)),

    # 电流相位错误
    dict(id='I-05', desc='条件13：∠IA∉[310°~350°] → Ia Phase Shift',
         phase_order=cfg.PHASE_ABC,
         src=_s(VL, 0, VL, 240, VL, 120, ia=I, qia=90, ic=I, qic=120)),
    dict(id='I-06', desc='条件14：∠IC∉[70°~110°] → Ic Phase Shift',
         phase_order=cfg.PHASE_ABC,
         src=_s(VL, 0, VL, 240, VL, 120, ia=I, qia=0, ic=I, qic=210)),
]

ACB_TESTS = [
    dict(id='PASS-ACB', desc='正常接线 Delta ACB',
         phase_order=cfg.PHASE_ACB,
         src=_s(VL, 0, VL, 120, VL, 240)),

    dict(id='V-08-ACB', desc='条件8 ACB：ABC信号接ACB配置 → Phase Order Error',
         phase_order=cfg.PHASE_ACB, src=_s(VL, 0, VL, 240, VL, 120)),

    dict(id='I-03-ACB', desc='条件11 ACB：∠IA≈210° → Ia Polarity Reversed',
         phase_order=cfg.PHASE_ACB,
         src=_s(VL, 0, VL, 120, VL, 240, ia=I, qia=180, ic=I, qic=240)),
    dict(id='I-04-ACB', desc='条件12 ACB：∠IC≈90° → Ic Polarity Reversed',
         phase_order=cfg.PHASE_ACB,
         src=_s(VL, 0, VL, 120, VL, 240, ia=I, qia=0, ic=I, qic=60)),
    dict(id='I-05-ACB', desc='条件13 ACB：∠IA∉[10°~50°] → Ia Phase Shift',
         phase_order=cfg.PHASE_ACB,
         src=_s(VL, 0, VL, 120, VL, 240, ia=I, qia=90, ic=I, qic=240)),
    dict(id='I-06-ACB', desc='条件14 ACB：∠IC∉[250°~290°] → Ic Phase Shift',
         phase_order=cfg.PHASE_ACB,
         src=_s(VL, 0, VL, 120, VL, 240, ia=I, qia=0, ic=I, qic=30)),
]

# 条件5/6 ||→&& 防回归（单侧线电压偏低，|| 触发但 && 不触发）
_REG_TESTS = [
    dict(id='V-05-REG', desc='条件5 防回归：Vbc正常，仅Vab偏低（Vca正常） → 不触发条件5',
         phase_order=cfg.PHASE_ABC,
         src=_s(int(VL*0.5), 0, VL, 240, VL, 120)),
    dict(id='V-06-REG', desc='条件6 防回归：Vca正常，仅Vbc偏低（Vab正常） → 不触发条件6',
         phase_order=cfg.PHASE_ABC,
         src=_s(VL, 0, int(VL*0.5), 240, VL, 120)),
]

ALL_TESTS = ABC_TESTS + ACB_TESTS + _REG_TESTS

_SPEC_REFS = {
    'PASS-ABC': '—', 'PASS-ACB': '—',
    'V-01': '2E3WDelta, 行2',
    'V-02': '2E3WDelta, 行3',
    'V-03': '2E3WDelta, 行4',
    'V-04': '2E3WDelta, 行5',
    'V-05': '2E3WDelta, 行6',
    'V-06': '2E3WDelta, 行7',
    'V-07': '2E3WDelta, 行8',
    'V-08': '2E3WDelta, 行9',   'V-08-ACB': '2E3WDelta, 行9',
    'I-01': '2E3WDelta, 行14',
    'I-02': '2E3WDelta, 行15',
    'I-03': '2E3WDelta, 行16',  'I-03-ACB': '2E3WDelta, 行16',
    'I-04': '2E3WDelta, 行17',  'I-04-ACB': '2E3WDelta, 行17',
    'I-05': '2E3WDelta, 行18',  'I-05-ACB': '2E3WDelta, 行18',
    'I-06': '2E3WDelta, 行19',  'I-06-ACB': '2E3WDelta, 行19',
    'V-05-REG': '2E3WDelta, 行6（防回归）',
    'V-06-REG': '2E3WDelta, 行7（防回归）',
}


def run_one(tc, modbus, wc_page):
    logging.info('[%s] %s', tc['id'], tc['desc'])
    t0 = time.time()
    s = tc['src']
    modbus.write_phase_order(tc['phase_order'])

    # Delta：只有 A/C 两相电流，ib 补零
    src.output(s['ua'], s['qua'], s['ub'], s['qub'], s['uc'], s['quc'],
               s['ia'], s['qia'], 0, 0, s['ic'], s['qic'])

    expected = compute_2e3w_delta(
        s['ua'], s['qua'], s['ub'], s['qub'], s['uc'], s['quc'],
        s['ia'], s['qia'], s['ic'], s['qic'],
        vrate=cfg.NOMINAL_VOLTAGE, phase_order=tc['phase_order'])
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
        modbus.write_service_config(cfg.SERVICE_2E3W_DELTA)
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
        print(f'\n2E3W Delta  总计:{total}  PASS:{passed}  FAIL:{total-passed}')
        for r in results:
            print(f"  [{'PASS' if r['pass_'] else 'FAIL'}] {r['id']}  {r['desc'][:55]}")
            if not r['pass_']:
                print(r['detail'])
        path = rpt.generate(results, '2E3W Delta', DEVICE_NAME, cfg.METER_TCP_IP,
                            active_channels=cfg.ACTIVE_USER_CHANNELS[cfg.SERVICE_2E3W_DELTA],
                            channel_phases=cfg.CHANNEL_PHASES[cfg.SERVICE_2E3W_DELTA])
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
    wiring_type='2E3W Delta',
    active_channels=cfg.ACTIVE_USER_CHANNELS[cfg.SERVICE_2E3W_DELTA],
    channel_phases=cfg.CHANNEL_PHASES[cfg.SERVICE_2E3W_DELTA],
)


@pytest.fixture(scope='module', autouse=True)
def _setup_2e3w_delta(wiring_modbus):
    wiring_modbus.write_service_config(cfg.SERVICE_2E3W_DELTA)


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
