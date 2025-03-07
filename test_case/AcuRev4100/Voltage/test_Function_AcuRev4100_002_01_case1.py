from comm.AcuRev4100_modbus_get_attr import *
from tools.log import Log
from comm.source_control import *

Log(str(__file__).split("\\")[-1])

Voltagelist = [69, 120]


def setup_function():
    logging.info('Start：{}'.format(time.strftime('%Y_%m_%d %H:%M:%S')))


def test_Line_to_Neutral_Voltage():
    rtu_result = []
    for i in Voltagelist:
        ret = set_ac(120, 240, 0, 120, 240, 0, i, i, i, 1, 1, 1, 50)
        Phase_A = Read_Phase_A_Voltage(i, 10)
        Phase_B = Read_Phase_B_Voltage(i, 10)
        Phase_C = Read_Phase_C_Voltage(i, 10)
        scale_A = abs(Phase_A - i) / i
        scale_B = abs(Phase_B - i) / i
        scale_C = abs(Phase_C - i) / i
        if i * (1 - 0.001) <= Phase_A <= i * (1 + 0.001) and i * (1 - 0.001) <= Phase_B <= i * (1 + 0.001) and i * (
                1 - 0.001) <= Phase_C <= i * (1 + 0.001):
            logging.info(
                f'源输入三项电压：{i}V，电表读取Phase_A：{Phase_A}V,电表读取Phase_B：{Phase_B}V,电表读取Phase_C：{Phase_C}V')
            rtu_result.append(True)
        else:
            logging.error(
                f'源输入三项电压：{i}V，电表读取Phase_A：{Phase_A}V,实际精度{scale_A:.2%},电表读取Phase_B：{Phase_B}V,实际精度{scale_B:.2%},电表读取Phase_C：{Phase_C}V,实际精度{scale_C:.2%}')
            rtu_result.append(False)
    assert rtu_result == [True, True]


def teardown_function():
    logging.info('End：{}'.format(time.strftime('%Y_%m_%d %H:%M:%S')))
    pass
