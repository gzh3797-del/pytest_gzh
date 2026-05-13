"""
接线检查脚本 ID → Excel 用例编号映射表

格式：TestCase_AcuHMI17_WRI_<子模块>_<序号>
子模块代码：3E4=3E4WY, DLT=2E3W Delta, NET=2E3W Network, 1PH=2E3W 1Phase, 1E2=1E2W
"""

_P = 'TestCase_AcuHMI17_WRI'


def _tc(sub3: str, seq: int) -> str:
    return f'{_P}_{sub3}_{seq:03d}'


# ── 3E4WY（test_3e4wy.py）────────────────────────────────────────────────────
TC_MAP_3E4WY: dict[str, str] = {
    'PASS-ABC':  _tc('3E4',  1),
    'V-01-ABC':  _tc('3E4',  2),
    'V-02-ABC':  _tc('3E4',  3),
    'V-03-ABC':  _tc('3E4',  4),
    'V-04-ABC':  _tc('3E4',  5),
    'V-05-ABC':  _tc('3E4',  6),
    'V-06-ABC':  _tc('3E4',  7),
    'V-07-ABC':  _tc('3E4',  8),
    'V-08-ABC':  _tc('3E4',  9),
    'V-09-ABC':  _tc('3E4', 10),
    'V-10-ABC':  _tc('3E4', 11),
    'V-11-ABC':  _tc('3E4', 12),
    'V-12-ABC':  _tc('3E4', 13),
    'V-13-ABC':  _tc('3E4', 14),
    'I-01-ABC':  _tc('3E4', 15),
    'I-02-ABC':  _tc('3E4', 16),
    'I-03-ABC':  _tc('3E4', 17),
    'I-04-ABC':  _tc('3E4', 18),
    'I-05-ABC':  _tc('3E4', 19),
    'I-06-ABC':  _tc('3E4', 20),
    'I-07-ABC':  _tc('3E4', 21),
    'I-08-ABC':  _tc('3E4', 22),
    'I-09-ABC':  _tc('3E4', 23),
    'PASS-ACB':  _tc('3E4', 24),
    'V-11-ACB':  _tc('3E4', 25),
    'V-12-ACB':  _tc('3E4', 26),
    'V-13-ACB':  _tc('3E4', 27),
    'I-05-ACB':  _tc('3E4', 28),
    'I-06-ACB':  _tc('3E4', 29),
    'I-08-ACB':  _tc('3E4', 30),
    'I-09-ACB':  _tc('3E4', 31),
    'V-02G-ABC': _tc('3E4', 32),
    'V-03G-ABC': _tc('3E4', 33),
    'V-04G-ABC': _tc('3E4', 34),
}

# ── 2E3W Delta（test_2e3w_delta.py）──────────────────────────────────────────
TC_MAP_DELTA: dict[str, str] = {
    'PASS-ABC':  _tc('DLT',  1),
    'V-01':      _tc('DLT',  2),
    'V-02':      _tc('DLT',  3),
    'V-03':      _tc('DLT',  4),
    'V-04':      _tc('DLT',  5),
    'V-05':      _tc('DLT',  6),
    'V-06':      _tc('DLT',  7),
    'V-07':      _tc('DLT',  8),
    'V-08':      _tc('DLT',  9),
    'I-01':      _tc('DLT', 10),
    'I-02':      _tc('DLT', 11),
    'I-03':      _tc('DLT', 12),
    'I-04':      _tc('DLT', 13),
    'I-05':      _tc('DLT', 14),
    'I-06':      _tc('DLT', 15),
    'PASS-ACB':  _tc('DLT', 16),
    'V-08-ACB':  _tc('DLT', 17),
    'I-03-ACB':  _tc('DLT', 18),
    'I-04-ACB':  _tc('DLT', 19),
    'I-05-ACB':  _tc('DLT', 20),
    'I-06-ACB':  _tc('DLT', 21),
    'V-05-REG':  _tc('DLT', 22),
    'V-06-REG':  _tc('DLT', 23),
}

# ── 2E3W Network（test_2e3w_network.py）──────────────────────────────────────
TC_MAP_NET: dict[str, str] = {
    'PASS-ABC':  _tc('NET',  1),
    'V-01-ABC':  _tc('NET',  2),
    'V-02-ABC':  _tc('NET',  3),
    'V-03-ABC':  _tc('NET',  4),
    'V-04-ABC':  _tc('NET',  5),
    'V-05-ABC':  _tc('NET',  6),
    'V-06-ABC':  _tc('NET',  7),
    'V-07-ABC':  _tc('NET',  8),
    'V-08-ABC':  _tc('NET',  9),
    'V-09-ABC':  _tc('NET', 10),
    'V-10-ABC':  _tc('NET', 11),
    'V-11-ABC':  _tc('NET', 12),
    'V-12-ABC':  _tc('NET', 13),
    'V-13-ABC':  _tc('NET', 14),
    'I-01-ABC':  _tc('NET', 15),
    'I-02-ABC':  _tc('NET', 16),
    'I-03-ABC':  _tc('NET', 17),
    'I-04-ABC':  _tc('NET', 18),
    'I-05-ABC':  _tc('NET', 19),
    'I-06-ABC':  _tc('NET', 20),
    'I-07-ABC':  _tc('NET', 21),
    'I-08-ABC':  _tc('NET', 22),
    'I-09-ABC':  _tc('NET', 23),
    'PASS-ACB':  _tc('NET', 24),
    'V-11-ACB':  _tc('NET', 25),
    'V-12-ACB':  _tc('NET', 26),
    'V-13-ACB':  _tc('NET', 27),
    'I-05-ACB':  _tc('NET', 28),
    'I-06-ACB':  _tc('NET', 29),
    'I-08-ACB':  _tc('NET', 30),
    'I-09-ACB':  _tc('NET', 31),
    'V-02G-ABC': _tc('NET', 32),
    'V-03G-ABC': _tc('NET', 33),
    'V-04G-ABC': _tc('NET', 34),
}

# ── 2E3W 1Phase（test_2e3w_1phase.py）────────────────────────────────────────
TC_MAP_1PHASE: dict[str, str] = {
    'PASS':  _tc('1PH',  1),
    'V-01':  _tc('1PH',  2),
    'V-02':  _tc('1PH',  3),
    'V-03':  _tc('1PH',  4),
    'V-04':  _tc('1PH',  5),
    'V-05':  _tc('1PH',  6),
    'I-01':  _tc('1PH',  7),
    'I-02':  _tc('1PH',  8),
    'I-03':  _tc('1PH',  9),
    'I-04':  _tc('1PH', 10),
    'I-05':  _tc('1PH', 11),
    'I-06':  _tc('1PH', 12),
}

# ── 1E2W（test_1e2w.py）───────────────────────────────────────────────────────
TC_MAP_1E2W: dict[str, str] = {
    'PASS':  _tc('1E2', 1),
    'V-01':  _tc('1E2', 2),
    'I-01':  _tc('1E2', 3),
    'I-02':  _tc('1E2', 4),
}
