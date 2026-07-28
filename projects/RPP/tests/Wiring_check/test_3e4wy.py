"""
3E4WY 接线检查自动化测试（HMI1-7）

对比逻辑：
  Expected — 由 expected_engine 根据控源参数 + 算法规格实时推导
  Actual   — 由 Playwright 从 HMI 页面 Wiring Status 列读取

覆盖：
  - 电压侧 13 条（条件1-13，ABC + ACB 相序）
  - 电流侧  9 条（条件14-22）
  - Pass 基准用例

运行：
    python projects/RPP/tests/Wiring_check/test_3e4wy.py
"""
import sys
import os
import logging
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '../../../..'))

from playwright.sync_api import sync_playwright

from projects.RPP.tests.Wiring_check.core import config as cfg
from projects.RPP.tests.Wiring_check.core.meter_modbus import WiringCheckModbus
from projects.RPP.tests.Wiring_check.core import signal_driver as src
from projects.RPP.tests.Wiring_check.core.expected_engine import compute_3e4wy
from projects.RPP.tests.Wiring_check.core.wiring_check_page import WiringCheckPage
from projects.RPP.tests.Wiring_check.core import report as rpt
from projects.RPP.tests.Wiring_check.core.tc_map import TC_MAP_3E4WY as _TC_MAP

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')

V = cfg.NOMINAL_VOLTAGE
I = cfg.NORMAL_CURRENT

# ── 设备名称（HMI 页面 Device 选择器中显示的名称）──────────────────────────
DEVICE_NAME = cfg.WIRING_DEVICE_NAME


# ─────────────────────────────────────────────────────────────────────────────
# 用例定义
# ─────────────────────────────────────────────────────────────────────────────

def _s(ua, qua, ub, qub, uc, quc, ia=I, qia=None, ib=I, qib=None, ic=I, qic=None):
    """构造控源参数，电流默认与对应电压同相（PF=1）"""
    return dict(
        ua=ua, qua=qua, ub=ub, qub=qub, uc=uc, quc=quc,
        ia=ia, qia=qua if qia is None else qia,
        ib=ib, qib=qub if qib is None else qib,
        ic=ic, qic=quc if qic is None else qic,
    )


ABC_TESTS = [
    # ── Pass 基准 ──
    dict(id='PASS-ABC',  desc='正常接线 ABC，应全部 Pass',
         phase_order=cfg.PHASE_ABC, src=_s(V, 0, V, 240, V, 120)),

    # ── 电压侧：接线缺失 ──
    dict(id='V-01-ABC',  desc='条件1：三相全<0.1*VRATE → Va&Vb&Vc Wiring Missing',
         phase_order=cfg.PHASE_ABC, src=_s(0, 0, 0, 0, 0, 0)),
    dict(id='V-02-ABC',  desc='条件2：Vab<0.1*VRATE (ua=ub=230@0°) → Va&Vb Wiring Missing',
         phase_order=cfg.PHASE_ABC, src=_s(V, 0, V, 0, V, 120)),
    dict(id='V-03-ABC',  desc='条件3：Vbc<0.1*VRATE (ub=uc=230@240°) → Vb&Vc Wiring Missing',
         phase_order=cfg.PHASE_ABC, src=_s(V, 0, V, 240, V, 240)),
    dict(id='V-04-ABC',  desc='条件4：Vca<0.1*VRATE (ua=uc=230@0°) → Va&Vc Wiring Missing',
         phase_order=cfg.PHASE_ABC, src=_s(V, 0, V, 240, V, 0)),

    # ── 条件2/3/4 guard 负向用例 ──
    dict(id='V-02G-ABC', desc='条件2 guard：Vab≈0 且 Vcn<0.8*VRATE → 条件2不触发，条件7 Vc Missing',
         phase_order=cfg.PHASE_ABC, src=_s(V, 0, V, 0, int(V*0.75), 120)),
    dict(id='V-03G-ABC', desc='条件3 guard：Vbc≈0 且 Van<0.8*VRATE → 条件3不触发，条件5 Va Missing',
         phase_order=cfg.PHASE_ABC, src=_s(int(V*0.75), 0, V, 240, V, 240)),
    dict(id='V-04G-ABC', desc='条件4 guard：Vca≈0 且 Vbn<0.8*VRATE → 条件4不触发，条件6 Vb Missing',
         phase_order=cfg.PHASE_ABC, src=_s(V, 0, int(V*0.75), 240, V, 0)),

    dict(id='V-05-ABC',  desc='条件5：Van=172V ≤ 0.8*VRATE → Va Wiring Missing',
         phase_order=cfg.PHASE_ABC, src=_s(int(V*0.75), 0, V, 240, V, 120)),
    dict(id='V-06-ABC',  desc='条件6：Vbn=172V ≤ 0.8*VRATE → Vb Wiring Missing',
         phase_order=cfg.PHASE_ABC, src=_s(V, 0, int(V*0.75), 240, V, 120)),
    dict(id='V-07-ABC',  desc='条件7：Vcn=172V ≤ 0.8*VRATE → Vc Wiring Missing',
         phase_order=cfg.PHASE_ABC, src=_s(V, 0, V, 240, int(V*0.75), 120)),

    # ── 电压侧：反接 ──
    dict(id='V-08-ABC',  desc='条件8：Van=230@0°,Vbn=400@30°,Vcn=400@330° → Va-Vn Reversed',
         phase_order=cfg.PHASE_ABC, src=_s(V, 0, 400, 30, 400, 330)),
    dict(id='V-09-ABC',  desc='条件9：Van=400@0°,Vbn=230@30°,Vcn=400@60° → Vb-Vn Reversed',
         phase_order=cfg.PHASE_ABC, src=_s(400, 0, V, 30, 400, 60)),
    dict(id='V-10-ABC',  desc='条件10：Van=400@0°,Vbn=400@300°,Vcn=230@330° → Vc-Vn Reversed',
         phase_order=cfg.PHASE_ABC, src=_s(400, 0, 400, 300, V, 330)),

    # ── 电压侧：相位错误 ──
    dict(id='V-11-ABC',  desc='条件11：∠VB=180° ∉ [220°~260°] → Vb Phase Shift',
         phase_order=cfg.PHASE_ABC, src=_s(V, 0, V, 180, V, 120)),
    dict(id='V-12-ABC',  desc='条件12：∠VC=60° ∉ [100°~140°] → Vc Phase Shift',
         phase_order=cfg.PHASE_ABC, src=_s(V, 0, V, 240, V, 60)),
    dict(id='V-13-ABC',  desc='条件13：ACB电压接入ABC配置 → Phase Order Error（+ Phase Shift B/C）',
         phase_order=cfg.PHASE_ABC, src=_s(V, 0, V, 120, V, 240)),

    # ── 电流侧：缺失 ──
    dict(id='I-01-ABC',  desc='条件14：Ian<0.1A → Ia Wiring Missing',
         phase_order=cfg.PHASE_ABC,
         src=_s(V, 0, V, 240, V, 120, ia=0.05, qia=0, ib=I, qib=240, ic=I, qic=120)),
    dict(id='I-02-ABC',  desc='条件15：Ibn<0.1A → Ib Wiring Missing',
         phase_order=cfg.PHASE_ABC,
         src=_s(V, 0, V, 240, V, 120, ia=I, qia=0, ib=0.05, qib=240, ic=I, qic=120)),
    dict(id='I-03-ABC',  desc='条件16：Icn<0.1A → Ic Wiring Missing',
         phase_order=cfg.PHASE_ABC,
         src=_s(V, 0, V, 240, V, 120, ia=I, qia=0, ib=I, qib=240, ic=0.05, qic=120)),

    # ── 电流侧：反接 ──
    dict(id='I-04-ABC',  desc='条件17：PF_A=-1（qia=180°）→ Ia Polarity Reversed',
         phase_order=cfg.PHASE_ABC,
         src=_s(V, 0, V, 240, V, 120, ia=I, qia=180, ib=I, qib=240, ic=I, qic=120)),
    dict(id='I-05-ABC',  desc='条件18：PF_B=-1（qib=60°）→ Ib Polarity Reversed',
         phase_order=cfg.PHASE_ABC,
         src=_s(V, 0, V, 240, V, 120, ia=I, qia=0, ib=I, qib=60, ic=I, qic=120)),
    dict(id='I-06-ABC',  desc='条件19：PF_C=-1（qic=300°）→ Ic Polarity Reversed',
         phase_order=cfg.PHASE_ABC,
         src=_s(V, 0, V, 240, V, 120, ia=I, qia=0, ib=I, qib=240, ic=I, qic=300)),

    # ── 电流侧：相位错误 ──
    dict(id='I-07-ABC',  desc='条件20：PF_A=0（qia=90°）→ Ia Phase Shift',
         phase_order=cfg.PHASE_ABC,
         src=_s(V, 0, V, 240, V, 120, ia=I, qia=90, ib=I, qib=240, ic=I, qic=120)),
    dict(id='I-08-ABC',  desc='条件21：PF_B=0（qib=330°）→ Ib Phase Shift',
         phase_order=cfg.PHASE_ABC,
         src=_s(V, 0, V, 240, V, 120, ia=I, qia=0, ib=I, qib=330, ic=I, qic=120)),
    dict(id='I-09-ABC',  desc='条件22：PF_C=0（qic=210°）→ Ic Phase Shift',
         phase_order=cfg.PHASE_ABC,
         src=_s(V, 0, V, 240, V, 120, ia=I, qia=0, ib=I, qib=240, ic=I, qic=210)),
]

ACB_TESTS = [
    dict(id='PASS-ACB',  desc='正常接线 ACB，应全部 Pass',
         phase_order=cfg.PHASE_ACB, src=_s(V, 0, V, 120, V, 240)),
    dict(id='V-11-ACB',  desc='条件11 ACB：∠VB=60° ∉ [100°~140°] → Vb Phase Shift',
         phase_order=cfg.PHASE_ACB, src=_s(V, 0, V, 60, V, 240)),
    dict(id='V-12-ACB',  desc='条件12 ACB：∠VC=180° ∉ [220°~260°] → Vc Phase Shift',
         phase_order=cfg.PHASE_ACB, src=_s(V, 0, V, 120, V, 180)),
    dict(id='V-13-ACB',  desc='条件13 ACB：ABC电压接入ACB配置 → Phase Order Error',
         phase_order=cfg.PHASE_ACB, src=_s(V, 0, V, 240, V, 120)),

    dict(id='I-05-ACB',  desc='条件18 ACB：PF_B=-1（qib=300°, Vb@120°）→ Ib Polarity Reversed',
         phase_order=cfg.PHASE_ACB,
         src=_s(V, 0, V, 120, V, 240, ia=I, qia=0, ib=I, qib=300, ic=I, qic=240)),
    dict(id='I-06-ACB',  desc='条件19 ACB：PF_C=-1（qic=60°, Vc@240°）→ Ic Polarity Reversed',
         phase_order=cfg.PHASE_ACB,
         src=_s(V, 0, V, 120, V, 240, ia=I, qia=0, ib=I, qib=120, ic=I, qic=60)),
    dict(id='I-08-ACB',  desc='条件21 ACB：PF_B=0（qib=210°, Vb@120°）→ Ib Phase Shift',
         phase_order=cfg.PHASE_ACB,
         src=_s(V, 0, V, 120, V, 240, ia=I, qia=0, ib=I, qib=210, ic=I, qic=240)),
    dict(id='I-09-ACB',  desc='条件22 ACB：PF_C=0（qic=330°, Vc@240°）→ Ic Phase Shift',
         phase_order=cfg.PHASE_ACB,
         src=_s(V, 0, V, 120, V, 240, ia=I, qia=0, ib=I, qib=120, ic=I, qic=330)),
]

ALL_TESTS     = ABC_TESTS + ACB_TESTS

_SPEC_REFS = {
    'PASS-ABC': '—',        'PASS-ACB': '—',
    'V-01-ABC': '3LN接线, 行2',
    'V-02-ABC': '3LN接线, 行3',
    'V-03-ABC': '3LN接线, 行4',
    'V-04-ABC': '3LN接线, 行5',
    'V-02G-ABC': '3LN接线, 行3（guard：Vcn偏低时条件2不触发）',
    'V-03G-ABC': '3LN接线, 行4（guard：Van偏低时条件3不触发）',
    'V-04G-ABC': '3LN接线, 行5（guard：Vbn偏低时条件4不触发）',
    'V-05-ABC': '3LN接线, 行6',
    'V-06-ABC': '3LN接线, 行7',
    'V-07-ABC': '3LN接线, 行8',
    'V-08-ABC': '3LN接线, 行9',
    'V-09-ABC': '3LN接线, 行10',
    'V-10-ABC': '3LN接线, 行11',
    'V-11-ABC': '3LN接线, 行12',  'V-11-ACB': '3LN接线, 行12',
    'V-12-ABC': '3LN接线, 行13',  'V-12-ACB': '3LN接线, 行13',
    'V-13-ABC': '3LN接线, 行14',  'V-13-ACB': '3LN接线, 行14',
    'I-01-ABC': '3LN接线, 行24',
    'I-02-ABC': '3LN接线, 行25',
    'I-03-ABC': '3LN接线, 行26',
    'I-04-ABC': '3LN接线, 行27',
    'I-05-ABC': '3LN接线, 行28',
    'I-06-ABC': '3LN接线, 行29',
    'I-07-ABC': '3LN接线, 行30',
    'I-08-ABC': '3LN接线, 行31',    'I-08-ACB': '3LN接线, 行31',
    'I-09-ABC': '3LN接线, 行32',    'I-09-ACB': '3LN接线, 行32',
    'I-05-ACB': '3LN接线, 行28',
    'I-06-ACB': '3LN接线, 行29',
}
VOLTAGE_TESTS = [{**tc, 'check_current': False}
                 for tc in ALL_TESTS if tc['id'].startswith(('PASS', 'V-'))]
CURRENT_TESTS = [tc for tc in ABC_TESTS if tc['id'].startswith(('PASS', 'I-'))]


# ─────────────────────────────────────────────────────────────────────────────
# 对比逻辑
# ─────────────────────────────────────────────────────────────────────────────

def _compare_voltage(expected: dict, actual: dict) -> tuple[bool, list[str]]:
    diffs = []
    for key in ('A', 'B', 'C', 'order'):
        exp = expected.get(key, '')
        act = actual.get(key, '')
        if not _status_match(exp, act):
            diffs.append(f'  电压 Phase {key}: 预期="{exp}"  实际="{act}"')
    return len(diffs) == 0, diffs


def _compare_current(expected: dict, actual_list: list[dict],
                     channel_phase_map=None) -> tuple[bool, list[str]]:
    diffs = []
    for idx, actual in enumerate(actual_list, start=1):
        if channel_phase_map and idx <= len(channel_phase_map):
            phases = channel_phase_map[idx - 1]
        else:
            phases = ('A', 'B', 'C')
        for phase in phases:
            exp = expected.get(phase, '')
            act = actual.get(phase, '')
            if not act or act in ('—', '-', 'N/A'):
                continue
            if not _status_match(exp, act):
                diffs.append(f'  电流 User{idx} Phase {phase}: 预期="{exp}"  实际="{act}"')
    return len(diffs) == 0, diffs


def _status_match(expected: str, actual: str) -> bool:
    e = expected.lower().strip()
    a = actual.lower().strip()
    if not e or e == 'n/a':
        return True
    if e == a:
        return True
    if e in ('pass',) and a in ('pass', 'ok', '—', '-', ''):
        return True
    if 'missing' in e and 'missing' in a:
        return True
    if 'reversed' in e and 'reversed' in a:
        return True
    if 'phase shift' in e and ('phase shift' in a or 'shift' in a):
        return True
    if 'phase order error' in e and ('error' in a or 'phase order' in a):
        return True
    return False


# ─────────────────────────────────────────────────────────────────────────────
# Runner
# ─────────────────────────────────────────────────────────────────────────────

def run_one(tc: dict, modbus: WiringCheckModbus, wc_page: WiringCheckPage) -> dict:
    logging.info('[%s] %s', tc['id'], tc['desc'])
    t0 = time.time()
    s = tc['src']

    modbus.write_phase_order(tc['phase_order'])

    src.output(s['ua'], s['qua'], s['ub'], s['qub'], s['uc'], s['quc'],
               s['ia'], s['qia'], s['ib'], s['qib'], s['ic'], s['qic'])

    expected = compute_3e4wy(
        s['ua'], s['qua'], s['ub'], s['qub'], s['uc'], s['quc'],
        s['ia'], s['qia'], s['ib'], s['qib'], s['ic'], s['qic'],
        vrate=cfg.NOMINAL_VOLTAGE,
        phase_order=tc['phase_order'],
    )
    logging.info('  预期: 电压=%s  电流=%s', expected['voltage'], expected['current'])

    actual_v, actual_i = wc_page.run_check()
    logging.info('  实测: 电压=%s  电流[0]=%s', actual_v, actual_i[0] if actual_i else '[]')

    v_ok, v_diffs = _compare_voltage(expected['voltage'], actual_v)
    if tc.get('check_current', True):
        i_ok, i_diffs = _compare_current(expected['current'], actual_i)
    else:
        i_ok, i_diffs = True, []
    passed = v_ok and i_ok

    all_diffs = v_diffs + i_diffs
    detail = '\n'.join(all_diffs) if all_diffs else 'OK'
    logging.info('  → %s', 'PASS' if passed else 'FAIL')
    if not passed:
        logging.warning(detail)

    flag  = 'PASS' if passed else 'FAIL'
    exp_v = expected['voltage']
    v_str = '  '.join(
        f"{ph}: {exp_v.get(ph, '')!r}→{actual_v.get(ph, '')!r}"
        for ph in ('A', 'B', 'C')
    )
    print(f'[{flag}] {tc["id"]}  {tc["desc"][:50]}')
    print(f'       电压  {v_str}')
    if tc.get('check_current', True) and actual_i:
        exp_i = expected['current']
        for idx, ai in enumerate(actual_i, start=1):
            i_str = '  '.join(
                f"{ph}: {exp_i.get(ph, '')!r}→{ai.get(ph, '')!r}"
                for ph in ('A', 'B', 'C') if ph in ai
            )
            print(f'       电流U{idx}  {i_str}')
    if not passed:
        print(f'       差异: {detail}')
    print()

    return dict(pass_=passed, expected=expected, actual_v=actual_v,
                actual_i=actual_i, detail=detail, phase_order=tc['phase_order'], src=s,
                spec_ref=_SPEC_REFS.get(tc['id'], ''),
                tc_id=_TC_MAP.get(tc['id'], ''), elapsed=round(time.time()-t0, 1))


def run_all(tests: list[dict] = ALL_TESTS, headless: bool = False):
    cfg.ensure_meter_connection(headless=headless)
    modbus = WiringCheckModbus()
    try:
        modbus.write_service_config(cfg.SERVICE_3E4WY)
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
        failed = total - passed
        print('\n' + '='*65)
        print(f'3E4WY Wiring Check  总计:{total}  PASS:{passed}  FAIL:{failed}')
        print('='*65)
        for r in results:
            flag = 'PASS' if r['pass_'] else 'FAIL'
            print(f"  [{flag}] {r['id']}  {r['desc'][:55]}")
            if not r['pass_']:
                print(r['detail'])
        print('='*65)

        report_path = rpt.generate(
            results=results,
            wiring_type='3E4WY',
            device_name=DEVICE_NAME,
            meter_ip=cfg.METER_TCP_IP,
            active_channels=cfg.ACTIVE_USER_CHANNELS[cfg.SERVICE_3E4WY],
            channel_phases=cfg.CHANNEL_PHASES[cfg.SERVICE_3E4WY],
        )
        print(f'报告已生成：{report_path}')
        return results
    except KeyboardInterrupt:
        print('[CTRL+C] 检测到中断，正在关源...')
        raise
    finally:
        src.stop()
        modbus.write_phase_order(cfg.PHASE_ABC)
        modbus.close()


# ── pytest 入口 ───────────────────────────────────────────────────────────────
import pytest

# 自定义 HTML 报告参数（pytest 路径下 conftest 的 wc_record 据此归类、生成报告）
REPORT_META = dict(
    wiring_type='3E4WY',
    active_channels=cfg.ACTIVE_USER_CHANNELS[cfg.SERVICE_3E4WY],
    channel_phases=cfg.CHANNEL_PHASES[cfg.SERVICE_3E4WY],
)


@pytest.fixture(scope='module', autouse=True)
def _setup_3e4wy(wiring_modbus):
    wiring_modbus.write_service_config(cfg.SERVICE_3E4WY)


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
    run_all(tests=_tests, headless=False)
