from comm.AcuRev4100_modbus_get_attr import *
from tools.log import Log
from comm.source_control import *

Log(str(__file__).split("\\")[-1])

Voltagelist = [69, 120, 230]


def setup_function():
    logging.info('Start：{}'.format(time.strftime('%Y_%m_%d %H:%M:%S')))


def test_Line_to_Neutral_Voltage():
    rtu_result = []
    ret = set_ac(120, 240, 0, 120, 240, 0, Voltagelist[0], Voltagelist[1], Voltagelist[2], 1, 1, 1, 50)
    Phase_A = Read_Phase_A_Voltage(Voltagelist[0], 10)
    Phase_B = Read_Phase_B_Voltage(Voltagelist[1], 10)
    Phase_C = Read_Phase_C_Voltage(Voltagelist[2], 10)
    if Voltagelist[0] * (1 - 0.001) <= Phase_A <= Voltagelist[0] * (1 + 0.001) and Voltagelist[1] * (
            1 - 0.001) <= Phase_B <= Voltagelist[1] * (1 + 0.001) and Voltagelist[2] * (
            1 - 0.001) <= Phase_C <= Voltagelist[2] * (1 + 0.001):
        logging.info(
            f'源输入三项电压：{Voltagelist[0], Voltagelist[1], Voltagelist[2]}V，电表读取Phase_A：{Phase_A}V,电表读取Phase_B：{Phase_B}V,电表读取Phase_C：{Phase_C}V')
        rtu_result.append(True)
    else:
        logging.error(
            f'源输入三项电压：{Voltagelist[0], Voltagelist[1], Voltagelist[2]}V，电表读取Phase_A：{Phase_A}V,电表读取Phase_B：{Phase_B}V,电表读取Phase_C：{Phase_C}V')
        rtu_result.append(False)
    assert rtu_result == [True]


def teardown_function():
    logging.info('End：{}'.format(time.strftime('%Y_%m_%d %H:%M:%S')))
    pass
