from comm.AcuRev4100_modbus_get_attr import *
from tools.log import Log
from comm.source_control import *

Log(str(__file__).split("\\")[-1])

Voltagelist = [40, 68]


# 电压40、68，无精度要求

def setup_function():
    logging.info('Start：{}'.format(time.strftime('%Y_%m_%d %H:%M:%S')))


def test_Line_to_Neutral_Voltage():
    rtu_result = []
    for i in Voltagelist:
        ret = set_ac(120, 240, 0, 120, 240, 0, i, i, i, 1, 1, 1, 50)
        Phase_A = Read_Phase_A_Voltage(i, 10)
        Phase_B = Read_Phase_B_Voltage(i, 10)
        Phase_C = Read_Phase_C_Voltage(i, 10)
        if Phase_A is None or Phase_B is None or Phase_C is None or Phase_A == 0 or Phase_B == 0 or Phase_C == 0:
            logging.error(
                f'源输入三项电压：{i}V，电表读取Phase_A：{Phase_A}V,电表读取Phase_B：{Phase_B}V,电表读取Phase_C：{Phase_C}V')
            rtu_result.append(False)
        else:
            logging.info(
                f'源输入三项电压：{i}V，电表读取Phase_A：{Phase_A}V,电表读取Phase_B：{Phase_B}V,电表读取Phase_C：{Phase_C}V')
            rtu_result.append(True)
    assert rtu_result == [True, True]


def teardown_function():
    logging.info('End：{}'.format(time.strftime('%Y_%m_%d %H:%M:%S')))
    pass
