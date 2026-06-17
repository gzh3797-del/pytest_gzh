"""
控源封装
set_ac 签名（注意 C/B/A 逆序）：
    set_ac(quc, qub, qua, qic, qib, qia, uc, ub, ua, ic, ib, ia, f)
"""
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '../../../..'))
from comm.source_control import set_ac, up_source_ac

FREQ = 50.0


def output(ua: float, qua: float,
           ub: float, qub: float,
           uc: float, quc: float,
           ia: float, qia: float,
           ib: float, qib: float,
           ic: float, qic: float):
    """
    输出三相交流信号。
    参数均以"A相优先"顺序传入（ua/qua → ub/qub → uc/quc），
    内部自动转为 set_ac 所需的 C/B/A 逆序。

    电压单位：V（如 230）
    电流单位：A（如 1.0）
    角度单位：度（0-360）
    """
    set_ac(quc, qub, qua,   # 电压相角 C/B/A
           qic, qib, qia,   # 电流相角 C/B/A
           uc,  ub,  ua,    # 电压幅值 C/B/A
           ic,  ib,  ia,    # 电流幅值 C/B/A
           FREQ)


def stop():
    """关源（各相幅值归零）"""
    up_source_ac()
