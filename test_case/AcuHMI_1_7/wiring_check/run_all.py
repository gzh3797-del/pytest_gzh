"""
一键运行所有接线方式的接线检查测试（HMI1-7），逐个生成 HTML 报告。

运行：
    python test_case/AcuHMI_1_7/wiring_check/run_all.py

可选参数（直接修改下方 WIRING_TYPES 列表来跳过某种接线方式）
"""
import sys
import os
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '../../..'))

from test_case.AcuHMI_1_7.wiring_check.test_3e4wy       import run_all as run_3e4wy,      ALL_TESTS as T_3E4WY
from test_case.AcuHMI_1_7.wiring_check.test_2e3w_delta   import run_all as run_delta,       ALL_TESTS as T_DELTA
from test_case.AcuHMI_1_7.wiring_check.test_2e3w_network import run_all as run_network,     ALL_TESTS as T_NETWORK
from test_case.AcuHMI_1_7.wiring_check.test_2e3w_1phase  import run_all as run_1phase,      ALL_TESTS as T_1PHASE
from test_case.AcuHMI_1_7.wiring_check.test_1e2w         import run_all as run_1e2w,        ALL_TESTS as T_1E2W

# 要执行的接线方式（注释掉不想跑的）
WIRING_TYPES = [
    ('3E4WY',        run_3e4wy,   T_3E4WY),
    ('2E3W Delta',   run_delta,   T_DELTA),
    ('2E3W Network', run_network, T_NETWORK),
    ('2E3W 1Phase',  run_1phase,  T_1PHASE),
    ('1E2W',         run_1e2w,    T_1E2W),
]

if __name__ == '__main__':
    summary = []
    grand_start = time.time()

    for name, runner, tests in WIRING_TYPES:
        print(f'\n{"="*60}')
        print(f'开始：{name}')
        print('='*60)
        t0 = time.time()
        try:
            results = runner(tests=tests, headless=False)
            elapsed = round(time.time() - t0, 1)
            total  = len(results)
            passed = sum(1 for r in results if r['pass_'])
            summary.append((name, total, passed, total - passed, elapsed))
        except Exception as e:
            elapsed = round(time.time() - t0, 1)
            print(f'[ERROR] {name} 运行异常：{e}')
            summary.append((name, 0, 0, -1, elapsed))

    grand_elapsed = round(time.time() - grand_start, 1)

    print(f'\n{"="*65}')
    print('全部接线方式汇总')
    print('='*65)
    total_all = passed_all = failed_all = 0
    for name, total, passed, failed, elapsed in summary:
        status = 'ERROR' if failed == -1 else f'PASS:{passed}  FAIL:{failed}'
        print(f'  {name:<18} 总计:{total:>3}  {status:<20} {elapsed}s')
        if failed != -1:
            total_all  += total
            passed_all += passed
            failed_all += failed
    print('-'*65)
    print(f'  {"合计":<18} 总计:{total_all:>3}  PASS:{passed_all}  FAIL:{failed_all}  总耗时:{grand_elapsed}s')
    print('='*65)
    print(f'报告已输出至：test_case/AcuHMI_1_7/wiring_check/reports/')
