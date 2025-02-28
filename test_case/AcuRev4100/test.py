import math
from comm.modbus_rtu_tcp import *
from comm.source_control import *
from tools.log import Log
import xlwt
from tools.excel_operate import data_read
from modbus_config import modbus_config
import sys
import __future__
from comm.source_control import *
from AcuRev4100_modbus_get import *
import cmath
import math


def calculate_angle(sequence_component_complex, sequence_component):
    if sequence_component > 0.00001:
        phase = cmath.phase(sequence_component_complex)
        angle = math.degrees(phase)

        if abs(angle) < 0.000001:
            angle = 0

        # 如果角度为负数，调整为 0 到 360 度之间
        if angle < 0:
            angle += 360

        return angle
    else:
        return 0


def sequence_component_calculation(a: float, b: float, c: float, a_angle: float, b_angle: float, c_angle: float):
    ret = []
    va_complex = complex(math.cos(a_angle * math.pi / 180) * a, math.sin(a_angle * math.pi / 180) * a)
    vb_complex = complex(math.cos(b_angle * math.pi / 180) * b, math.sin(b_angle * math.pi / 180) * b)
    vc_complex = complex(math.cos(c_angle * math.pi / 180) * c, math.sin(c_angle * math.pi / 180) * c)
    Rotation_Factor = complex(-1 / 2, math.sqrt(3) / 2)
    Rotation_Factor_square = Rotation_Factor ** 2
    zero_sequence_component_complex = (va_complex + vb_complex + vc_complex) / 3
    zero_sequence_component = abs(zero_sequence_component_complex)
    zero_seq_calculate_angle = round(calculate_angle(zero_sequence_component_complex, zero_sequence_component), 3)

    positive_sequence_component_complex = (va_complex / 3) + ((vb_complex * Rotation_Factor) / 3) + (
            (Rotation_Factor_square * vc_complex) / 3)
    positive_sequence_component = abs(positive_sequence_component_complex)
    positive_seq_calculate_angle = round(
        calculate_angle(positive_sequence_component_complex, positive_sequence_component), 3)

    negative_sequence_component_complex = (va_complex / 3) + ((vb_complex * Rotation_Factor_square) / 3) + (
            (Rotation_Factor * vc_complex) / 3)
    negative_sequence_component = abs(negative_sequence_component_complex)
    negative_seq_calculate_angle = round(
        calculate_angle(negative_sequence_component_complex, negative_sequence_component), 3)
    VUF_CUF = negative_sequence_component / positive_sequence_component

    return zero_sequence_component, zero_seq_calculate_angle, positive_sequence_component, positive_seq_calculate_angle, negative_sequence_component, negative_seq_calculate_angle, VUF_CUF


if __name__ == '__main__':
    # Set_Service_Configuration(0)
    ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, 0.1, 0.1, 0.2, 50)
    # Set_Clear_energy(1)

