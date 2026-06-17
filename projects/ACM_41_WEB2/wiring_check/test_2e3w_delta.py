"""
2E3W Delta 接线检查自动化测试（电压侧8条 + 电流侧6条）

运行：
    python test_case/ACM_41_WEB2/wiring_check/test_2e3w_delta.py
"""
import sys, os, logging, time
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '../../..'))

from playwright.sync_api import sync_playwright
from projects.ACM_41_WEB2.wiring_check.core import config as cfg
from projects.ACM_41_WEB2.wiring_check.core.meter_modbus import WiringCheckModbus
from projects.ACM_41_WEB2.wiring_check.core import signal_driver as src
from projects.ACM_41_WEB2.wiring_check.core.expected_engine import compute_2e3w_delta
from projects.ACM_41_WEB2.wiring_check.core.wiring_check_page import WiringCheckPage
from projects.ACM_41_WEB2.wiring_check.core import report as rpt
from projects.ACM_41_WEB2.wiring_check.core.tc_map import TC_MAP_DELTA as _TC_MAP
from projects.ACM_41_WEB2.wiring_check.test_3e4wy import (
    _compare_voltage, _compare_current, _status_match, DEVICE_NAME)

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')

V  = cfg.NOMINAL_VOLTAGE   # 230V（线电压）
VL = round(V / 1.732, 1)   # ≈133V 对应的相电压
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
    dict(id='V-01', desc='条件1：三相线电压全<0.1*VRATE',
         phase_order=cfg.PHASE_ABC, src=_s(0, 0, 0, 0, 0, 0)),
    dict(id='V-02', desc='条件2：Vab<0.1*VRATE (ua=ub同相)',
         phase_order=cfg.PHASE_ABC, src=_s(VL, 0, VL, 0, VL, 120)),
    dict(id='V-03', desc='条件3：Vbc<0.1*VRATE (ub=uc同相)',
         phase_order=cfg.PHASE_ABC, src=_s(VL, 0, VL, 240, VL, 240)),
    dict(id='V-04', desc='条件4：Vca<0.1*VRATE (uc=ua同相)',
         phase_order=cfg.PHASE_ABC, src=_s(VL, 0, VL, 240, VL, 0)),
    dict(id='V-05', desc='条件5：Va缺失 (ua=0, Vbc正常)',
         phase_order=cfg.PHASE_ABC, src=_s(0, 0, VL, 240, VL, 120)),
    dict(id='V-06', desc='条件6：Vb缺失 (ub=0, Vca正常)',
         phase_order=cfg.PHASE_ABC, src=_s(VL, 0, 0, 0, VL, 120)),
    dict(id='V-07', desc='条件7：Vc缺失 (uc=0, Vab正常)',
         phase_order=cfg.PHASE_ABC, src=_s(VL, 0, VL, 240, 0, 0)),

    # 相序错误
    dict(id='V-08', desc='条件8：ABC配置下输入ACB信号 → Phase Order Error',
         phase_order=cfg.PHASE_ABC, src=_s(VL, 0, VL, 120, VL, 240)),

    # 电流缺失
    dict(id='I-01', desc='条件9：Ia<0.1A → Ia Wiring Missing',
         phase_order=cfg.PHASE_ABC,
         src=_s(VL, 0, VL, 240, VL, 120, ia=0.05, qia=330, ic=I, qic=90)),
    dict(id='I-02', desc='条件10：Ic<0.1A → Ic Wiring Missing',
         phase_order=cfg.PHASE_ABC,
         src=_s(VL, 0, VL, 240, VL, 120, ia=I, qia=330, ic=0.05, qic=90)),

    # 电流反接
    dict(id='I-03', desc='条件11：∠IA_rel=150°（绝对@180°）→ Ia Polarity Reversed',
         phase_order=cfg.PHASE_ABC,
         src=_s(VL, 0, VL, 240, VL, 120, ia=I, qia=180, ic=I, qic=120)),
    dict(id='I-04', desc='条件12：∠IC_rel=270°（绝对@300°）→ Ic Polarity Reversed',
         phase_order=cfg.PHASE_ABC,
         src=_s(VL, 0, VL, 240, VL, 120, ia=I, qia=0, ic=I, qic=300)),

    # 电流相位错误
    dict(id='I-05', desc='条件13：∠IA_rel=90°（绝对@120°）∉ [330°±20°] → Ia Phase Shift',
         phase_order=cfg.PHASE_ABC,
         src=_s(VL, 0, VL, 240, VL, 120, ia=I, qia=120, ic=I, qic=120)),
    dict(id='I-06', desc='条件14：∠IC_rel=180°（绝对@210°）∉ [90°±20°] → Ic Phase Shift',
         phase_order=cfg.PHASE_ABC,
         src=_s(VL, 0, VL, 240, VL, 120, ia=I, qia=0, ic=I, qic=210)),

    # ── 条件5/6 ||→&& 防回归：仅单侧线电压偏低，||触发 && 不触发 ──────────────
    # ang_Vab ≈ 345°（两组信号相同），对应正常电流角：ia@315°, ic@75°
    dict(id='V-05-REG',
         desc='条件5 ||回归：ua@270°→Vab≈69V单侧偏低、Vca≈257V正常 → Va Missing（&&则漏报Pass）',
         phase_order=cfg.PHASE_ABC,
         src=_s(VL, 270, VL, 240, VL, 120, ia=I, qia=315, ic=I, qic=75)),
    dict(id='V-06-REG',
         desc='条件6 ||回归：ub@150°→Vbc≈69V单侧偏低、Vab≈257V正常 → Vb Missing（&&则漏报Pass）',
         phase_order=cfg.PHASE_ABC,
         src=_s(VL, 0, VL, 150, VL, 120, ia=I, qia=315, ic=I, qic=75)),
    # 条件7 无法构造单侧回归：单侧偏低时条件5或6会率先触发（elif优先级），
    # 条件7只在Vbc和Vca同时偏低时可达，&& 与 || 结果相同，无法区分。
]

ACB_TESTS = [
    dict(id='PASS-ACB', desc='正常接线 Delta ACB',
         phase_order=cfg.PHASE_ACB, src=_s(VL, 0, VL, 120, VL, 240, qia=0, qic=240)),
    dict(id='V-08-ACB', desc='条件8 ACB：ACB配置下输入ABC信号 → Phase Order Error',
         phase_order=cfg.PHASE_ACB, src=_s(VL, 0, VL, 240, VL, 120, qia=0, qic=240)),

    # ── 电流侧 ACB（角度阈值与 ABC 不同，需独立用例）──
    dict(id='I-03-ACB', desc='条件11 ACB：∠IA_rel=210°（绝对@180°）→ Ia Polarity Reversed',
         phase_order=cfg.PHASE_ACB,
         src=_s(VL, 0, VL, 120, VL, 240, ia=I, qia=180, ic=I, qic=240)),
    dict(id='I-04-ACB', desc='条件12 ACB：∠IC_rel=90°（绝对@60°）→ Ic Polarity Reversed',
         phase_order=cfg.PHASE_ACB,
         src=_s(VL, 0, VL, 120, VL, 240, ia=I, qia=0, ic=I, qic=60)),
    dict(id='I-05-ACB', desc='条件13 ACB：∠IA_rel=90°（绝对@60°）∉ [30°±20°] → Ia Phase Shift',
         phase_order=cfg.PHASE_ACB,
         src=_s(VL, 0, VL, 120, VL, 240, ia=I, qia=60, ic=I, qic=240)),
    dict(id='I-06-ACB', desc='条件14 ACB：∠IC_rel=180°（绝对@150°）∉ [270°±20°] → Ic Phase Shift',
         phase_order=cfg.PHASE_ACB,
         src=_s(VL, 0, VL, 120, VL, 240, ia=I, qia=0, ic=I, qic=150)),
]

ALL_TESTS     = ABC_TESTS + ACB_TESTS

# 用例 → 接线检测总表_ver1.05.xlsx 行号对照（Sheet: 2LL Delta(2E3W Delta)）
_SPEC_REFS = {
    'PASS-ABC': '—',       'PASS-ACB': '—',
    'V-01': '2LL Delta, 行2',
    'V-02': '2LL Delta, 行3',
    'V-03': '2LL Delta, 行4',
    'V-04': '2LL Delta, 行5',
    'V-05': '2LL Delta, 行6',
    'V-06': '2LL Delta, 行7',
    'V-07': '2LL Delta, 行8',
    'V-08': '2LL Delta, 行9',   'V-08-ACB': '2LL Delta, 行9',
    'I-01': '2LL Delta, 行19',
    'I-02': '2LL Delta, 行20',
    'I-03': '2LL Delta, 行21',
    'I-04': '2LL Delta, 行22',
    'I-05': '2LL Delta, 行23',    'I-05-ACB': '2LL Delta, 行23',
    'I-06': '2LL Delta, 行24',    'I-06-ACB': '2LL Delta, 行24',
    'I-03-ACB': '2LL Delta, 行21',
    'I-04-ACB': '2LL Delta, 行22',
    'V-05-REG': '2LL Delta, 条件5（||防&&回归）',
    'V-06-REG': '2LL Delta, 条件6（||防&&回归）',
}
VOLTAGE_TESTS = [{**tc, 'check_current': False}
                 for tc in ALL_TESTS if tc['id'].startswith(('PASS', 'V-'))]


def run_one(tc, modbus, wc_page):
    logging.info('[%s] %s', tc['id'], tc['desc'])
    t0 = time.time()
    s = tc['src']
    modbus.write_phase_order(tc['phase_order'])
    src.output(s['ua'], s['qua'], s['ub'], s['qub'], s['uc'], s['quc'],
               s['ia'], s['qia'], I, s['qia'],   # ib 不参与 Delta，给默认值
               s['ic'], s['qic'])

    expected = compute_2e3w_delta(
        s['ua'], s['qua'], s['ub'], s['qub'], s['uc'], s['quc'],
        s['ia'], s['qia'], s['ic'], s['qic'],
        vrate=cfg.NOMINAL_VOLTAGE, phase_order=tc['phase_order'])
    logging.info('  预期: %s', expected)

    actual_v, actual_i = wc_page.run_check()
    logging.info('  实测: 电压=%s', actual_v)

    v_ok, v_diffs = _compare_voltage(expected['voltage'], actual_v)
    if tc.get('check_current', True):
        i_ok, i_diffs = _compare_current(expected['current'], actual_i)
    else:
        i_ok, i_diffs = True, []
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
    run_start = time.time()
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
            if not r['pass_']: print(r['detail'])
        path = rpt.generate(results, '2E3W Delta', DEVICE_NAME, cfg.METER_TCP_IP, active_channels=cfg.ACTIVE_USER_CHANNELS[cfg.SERVICE_2E3W_DELTA], channel_phases=cfg.CHANNEL_PHASES[cfg.SERVICE_2E3W_DELTA])
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
def _setup_2e3w_delta(wiring_modbus):
    wiring_modbus.write_service_config(cfg.SERVICE_2E3W_DELTA)


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
