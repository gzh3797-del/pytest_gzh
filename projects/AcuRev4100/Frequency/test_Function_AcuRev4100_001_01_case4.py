from comm.AcuRev4100_modbus_get_attr import *
from tools.log import Log
from comm.source_control import *

Log(str(__file__).split("\\")[-1])

Frelist = [40, 70]
# 频率40、70，无精度要求


def setup_function():
    logging.info('Start：{}'.format(time.strftime('%Y_%m_%d %H:%M:%S')))


def test_frequency_0x2000():
    rtu_result = []
    for i in Frelist:
        # ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, 1, 1, 1, i)
        fre = read_frequency(i, 5)
        if fre is None or fre == 0:
            logging.error(f'源输入频率：{i}，电表读取频率：{fre}')
            rtu_result.append(False)
        else:
            logging.info(f'源输入频率：{i}，电表读取频率：{fre}')
            rtu_result.append(True)
    assert rtu_result == [True, True]


def teardown_function():
    logging.info('End：{}'.format(time.strftime('%Y_%m_%d %H:%M:%S')))
    pass
