import math

import sys
import __future__
from comm.source_control import *


def line_to_line_voltage_calculate(ua: float, ub: float, uc: float, va_angle: float, vb_angle: float, vc_angle: float):
    ret = []
    va_complex = complex(math.cos(va_angle * math.pi / 180) * ua, math.sin(va_angle * math.pi / 180) * ua)
    vb_complex = complex(math.cos(vb_angle * math.pi / 180) * ub, math.sin(vb_angle * math.pi / 180) * ub)
    vc_complex = complex(math.cos(vc_angle * math.pi / 180) * uc, math.sin(vc_angle * math.pi / 180) * uc)
    vab = va_complex - vb_complex
    vbc = vb_complex - vc_complex
    vca = vc_complex - va_complex
    vab = abs(vab)
    vbc = abs(vbc)
    vca = abs(vca)
    ret.extend([vab, vbc, vca])
    return ret


if __name__ == '__main__':
    ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, 2, 2, 2, 50)
    print(ret)
